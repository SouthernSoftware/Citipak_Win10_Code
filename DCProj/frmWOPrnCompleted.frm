VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmWOListPrint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Work Orders List By Book"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmWOPrnCompleted.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboWOStatus 
      Height          =   348
      Left            =   5376
      TabIndex        =   4
      Top             =   4864
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
      ColDesigner     =   "frmWOPrnCompleted.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5376
      TabIndex        =   7
      Top             =   6336
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
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
      ColDesigner     =   "frmWOPrnCompleted.frx":0C30
   End
   Begin LpLib.fpCombo fpCombo1 
      Height          =   348
      Left            =   5376
      TabIndex        =   6
      Top             =   5844
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
      ColDesigner     =   "frmWOPrnCompleted.frx":0F96
   End
   Begin LpLib.fpCombo fpcboCustStatus 
      Height          =   348
      Left            =   5376
      TabIndex        =   5
      Top             =   5354
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
      ColDesigner     =   "frmWOPrnCompleted.frx":12C5
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
      TabIndex        =   9
      Top             =   7464
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
      TabIndex        =   8
      Top             =   7464
      Width           =   1332
   End
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5376
      TabIndex        =   1
      Top             =   3394
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
      OnFocusNoSelect =   0   'False
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5376
      TabIndex        =   0
      Top             =   2904
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
      OnFocusNoSelect =   0   'False
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
      TabIndex        =   13
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
            TextSave        =   "4:53 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/17/2005"
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5376
      TabIndex        =   2
      Top             =   3884
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   5376
      TabIndex        =   3
      Top             =   4374
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
      Caption         =   "To Date:"
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
      Left            =   3840
      TabIndex        =   19
      Top             =   4416
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Date:"
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
      Left            =   3744
      TabIndex        =   18
      Top             =   3936
      Width           =   1476
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order Status:"
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
      Left            =   2820
      TabIndex        =   17
      Top             =   4884
      Width           =   2388
   End
   Begin VB.Label Label2 
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
      Left            =   3420
      TabIndex        =   16
      Top             =   5844
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
      Left            =   3132
      TabIndex        =   15
      Top             =   5376
      Width           =   2076
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type:"
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
      Left            =   2820
      TabIndex        =   14
      Top             =   6360
      Width           =   2388
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4428
      Left            =   2364
      Top             =   2616
      Width           =   7428
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3222
      Top             =   1200
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order List By Book"
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
      Index           =   0
      Left            =   3300
      TabIndex        =   12
      Top             =   1368
      Width           =   5628
   End
   Begin VB.Label LabelB1 
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
      Left            =   3732
      TabIndex        =   11
      Top             =   2940
      Width           =   1476
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Book:"
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
      Left            =   3828
      TabIndex        =   10
      Top             =   3408
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3222
      Top             =   1080
      Width           =   5772
   End
End
Attribute VB_Name = "frmWOListPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub cmdPrint_Click()
  If ValidDate Then
    PrintWOList
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



Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub
Private Sub txtDate1_LostFocus()
  If CheckValDate(txtDate1) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate1.SetFocus
  End If
End Sub
Private Sub txtDate2_LostFocus()
  If CheckValDate(txtDate2) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboWOStatus.SetFocus
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
        fpCombo1.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboWOStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboWOStatus.ListDown = True
  End If
  If fpcboWOStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboCustStatus.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
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
        fpcboWOStatus.SetFocus
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

Private Sub fptxtCopies_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate1.SetFocus
  End If
End Sub
Private Sub fptxtCopies_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub
Private Sub cmdExit_Click()
  frmUBWorkOrderMenu.Show
  Unload frmWOPrintBook
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via WOPrintBook by " + PWUser$
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  fptxtRoute1 = 0
  fptxtRoute2 = 99
  fpcboWOStatus.AddItem "Open"
  fpcboWOStatus.AddItem "Completed"
  fpcboWOStatus.AddItem "Both"
  fpcboWOStatus.ListIndex = 0
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
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
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

Private Sub PrintWOList()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim Dash As String, PrintSingleFlag As Boolean, Copies As Integer
  Dim ReportFile As String, RptHandle As Integer, IdxName As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, IdxNumOfRecs As Long
  Dim NumOfRecs As Long, Handle As Integer, UBCustF As Integer
  Dim UBWOFile As Integer, lcnt As Long, Book As Integer, cnt As Long
  Dim BegRoute As Integer, EndRoute As Integer, Acct As Long
  Dim Title As String, CopyCnt As Integer, MtrCnt As Integer
  Dim Rem1 As String, Rem2 As String, Rem3 As String, Rem4 As String
  Dim Rem5 As String, Rem6 As String, ToPrint As String
  Dim graphicflag As Boolean, UseStatus As Boolean, AcctLu As Long
  Dim RStatus As String, Stat As String, UsingAcct As Boolean
  Dim UsingBook As Boolean, UsingName As Boolean, Page As Integer
  Dim WOStatus As Integer, FromDate As Integer, ThruDate As Integer
  Dim WorkDate As Integer
  FromDate = Date2Num(txtDate1)
  ThruDate = Date2Num(txtDate2)

  ToPrint$ = ""
  MaxLines = 50
  FF$ = Chr$(12)
  'Open Report File
  ReportFile$ = UBPath$ + "UBOPNWRK.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  WOStatus = fpcboWOStatus.ListIndex
  
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
 
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer
 
  Select Case fpCombo1.ListIndex
   Case 0
    UsingBook = True
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
    Title$ = "WorkOrders in Location Order."
   Case 1
    UsingAcct = True
    IdxName$ = ""
    Title$ = "WorkOrders in Account Order."
   Case 2
    UsingName = True
    IdxName$ = UBPath$ + "UBCUSTNM.IDX"
    Title$ = "WorkOrders in Name Order."
  End Select
  If Not UsingAcct Then
    IdxFileSize& = FileSize(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    NumOfRecs = IdxNumOfRecs
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
  UBCustF = FreeFile
  Open UBCustFile For Random Shared As UBCustF Len = UBCustRecLen

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  Rem1$ = ""
  Rem2$ = ""
  Rem3$ = ""
  Rem4$ = ""
  Rem5$ = ""
  Rem6$ = ""
  If fpcboRptType.ListIndex = 0 Then
    graphicflag = True
  Else
    graphicflag = False
  End If
  DeActivateControls Me
  FrmShowPctComp.Label1 = "Creating Work Order Listing"
  FrmShowPctComp.Show , Me

  If graphicflag = True Then
    Dash$ = String$(83, "_")
  Else
    Dash$ = String$(79, "_")
  End If
  ToPrint$ = ""
  FF$ = Chr$(12)
  BegRoute = Val(fptxtRoute1)
  EndRoute = Val(fptxtRoute2)

skipthis:

  'Open Report File
  ReportFile$ = UBPath$ + "WORKORDL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  UBWOFile = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWOFile Len = WorkOrderRecLen
    If graphicflag = False Then GoSub PrintReadHeading
    cnt& = 1
    'ShowProcessingScrn "Processing Work Orders"
    For lcnt& = 1 To NumOfRecs
      FrmShowPctComp.ShowPctComp lcnt, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        ActivateControls Me
        GoTo ExitHere
      End If
      If Not UsingAcct Then
        AcctLu = IdxBuff(lcnt).RecNum
      Else
        AcctLu = lcnt
      End If
      Get #UBCustF, AcctLu, UBCustRec(1)
      If UBCustRec(1).DelFlag <> 0 Then
        GoTo DelSkip
      End If
      If UseStatus Then
        If UBCustRec(1).Status <> RStatus$ Then
          GoTo DelSkip
        End If
      End If
      Book = Val(UBCustRec(1).Book)
      If Book >= BegRoute And Book <= EndRoute Then
        If UBCustRec(1).WOLastTrans > 0 Then
          Get #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
          If WOStatus = 0 Then
            If WorkOrderRec(1).CompletedDate <= 0 Then
              GoSub PrintThemOne
            End If
          ElseIf WOStatus = 1 Then
            If WorkOrderRec(1).CompletedDate > 0 Then
              GoSub PrintThemOne
            End If
          ElseIf WOStatus = 2 Then
            GoSub PrintThemOne
          End If
        End If
      End If
      'ShowPctComp lcnt&, IdxNumOfRecs
DelSkip:
    Next

  'PRINT #RptHandle, FF$

  Close
  Erase UBCustRec, WorkOrderRec, IdxBuff

  'Header$ = "Customer Work Orders "
  'PrintRptFile Header$, ReportFile$, LPTPort, RetCode, EntryPoint
  Erase IdxBuff
  Title$ = Title$ + " " + QPTrim$(fpcboWOStatus.Text)
  If graphicflag = False Then
    ViewPrint ReportFile$, Title$
    ActivateControls Me
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmWOListPrint
    If WOStatus = 0 Then
      ARptWorkOrdersOpen.txtDate = Now
      ARptWorkOrdersOpen.txtTown = TOWNNAME$
      ARptWorkOrdersOpen.Title = Title$
      ARptWorkOrdersOpen.GetName ReportFile$
      ARptWorkOrdersOpen.startrpt
    Else
      ARptWorkOrdersCompleted.txtDate = Now
      ARptWorkOrdersCompleted.txtTown = TOWNNAME$
      ARptWorkOrdersCompleted.Title = Title$
      ARptWorkOrdersCompleted.GetName ReportFile$
      ARptWorkOrdersCompleted.startrpt
    End If
  End If
ExitHere:
  Close
  Erase IdxBuff
  Exit Sub

PrintThemOne:
  WorkDate = WorkOrderRec(1).ENTRYDATE
  If WorkDate >= FromDate And WorkDate <= ThruDate Then
    If graphicflag = False Then
      Print #RptHandle, Using("######", AcctLu);
      Print #RptHandle, Tab(12); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; Tab(28); QPTrim$(UBCustRec(1).CustName);
      If WOStatus > 0 Then
        Print #RptHandle, Tab(65); Num2Date$(WorkOrderRec(1).CompletedDate)
      Else
        Print #RptHandle,
      End If
      Print #RptHandle, QPTrim$(UBCustRec(1).ServAddr); Tab(50); UBCustRec(1).WOLastTrans; Tab(65); Num2Date$(WorkOrderRec(1).ENTRYDATE)
      Print #RptHandle, Dash$
      LineCnt = LineCnt + 3
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintReadHeading
      End If
    Else
      ToPrint$ = Using("######", AcctLu) + "~"
      ToPrint$ = ToPrint$ + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + QPTrim$(UBCustRec(1).CustName)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ServAddr) + "~" + Str(UBCustRec(1).WOLastTrans) + "~" + Num2Date$(WorkOrderRec(1).ENTRYDATE)
      If WOStatus > 0 Then ToPrint$ = ToPrint$ + "~" + Num2Date$(WorkOrderRec(1).CompletedDate)
      Print #RptHandle, ToPrint$
      ToPrint$ = ""
    End If
   End If
 Return
        
PrintReadHeading:
  Page = Page + 1
  Print #RptHandle, Tab(30); Title$
  Print #RptHandle, "Date: "; Date$; Tab(70); "Page #"; Page
  Print #RptHandle, "Acct No.   Location        Customer Name"
  Print #RptHandle, "Service Address                               Work Order #      Order Dates"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5

  Return

End Sub
''Public Sub PrintOpenWorkOrderRpt(graphicflag As Boolean)
''  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer, Dash As String
''  Dim ReportFile As String, RptHandle As Integer, CustName As String
''  Dim IdxRecLen As Integer, IdxFileSize As Long, IdxNumOfRecs As Long
''  Dim NumOfRecs As Long, Handle As Integer, cnt As Long, UBWOFile As Integer
''  Dim lcnt As Long, Header As String, Page As Integer, IdxName As String
''  Dim UBCustF As Integer, ToPrint As String
''  ReDim UBCustRec(1) As NewUBCustRecType
''  UBCustRecLen = Len(UBCustRec(1))
''
''  ReDim WorkOrderRec(1) As WorkOrderRecType
''  WorkOrderRecLen = Len(WorkOrderRec(1))
''  ToPrint$ = ""
''  Dash$ = String$(79, "-")
''
''  MaxLines = 50
''  FF$ = Chr$(12)
''
''  'Open Report File
''  ReportFile$ = UBPath$ + "UBOPNWRK.RPT"
''  RptHandle = FreeFile
''  Open ReportFile$ For Output As #RptHandle
''  CustName$ = Space$(30)
''
''  ' Location Order ********************************************************
''  IdxName$ = UBPath$ + "UBCUSTBK.IDX"
''  IdxRecLen = 4 'we are using a long integer
''  IdxFileSize& = FileSize&(IdxName$)
''  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
''
''  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
''  'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
''  NumOfRecs = IdxNumOfRecs
''  Handle = FreeFile
''  Open IdxName$ For Random Shared As Handle Len = IdxRecLen
''  For cnt& = 1 To IdxNumOfRecs
''    Get #Handle, cnt&, IdxBuff(cnt&)
''  Next
''  Close Handle
''
''  UBCustF = FreeFile
''  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
''
''  UBWOFile = FreeFile
''  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWOFile Len = WorkOrderRecLen
''  FrmShowPctComp.Label1 = "Creating Customer Work Order Listing"
''  FrmShowPctComp.Show , Me
''
''  cnt& = 1
''  If graphicflag = False Then
''    GoSub PrintReadHeading
''    'ShowProcessingScrn "Processing Open Work Orders"
''    For lcnt& = 1 To IdxNumOfRecs
''      FrmShowPctComp.ShowPctComp lcnt&, IdxNumOfRecs
''      If FrmShowPctComp.Out Then
''        Close
''        Unload FrmShowPctComp
''        GoTo ExitHere
''      End If
''
''      Get #UBCustF, IdxBuff(lcnt&).RecNum, UBCustRec(1)
''      If UBCustRec(1).DelFlag <> 0 Then
''        GoTo OpenRptSkip
''      End If
''
''      If UBCustRec(1).WOLastTrans > 0 Then
''        Get #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
''        If WorkOrderRec(1).CompletedDate <= 0 Then
''          Print #RptHandle, Using("######", IdxBuff(lcnt&).RecNum);
''          Print #RptHandle, Tab(12); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; Tab(28); QPTrim$(UBCustRec(1).CustName)
''          Print #RptHandle, QPTrim$(UBCustRec(1).ServAddr); Tab(50); UBCustRec(1).WOLastTrans; Tab(65); Num2Date$(WorkOrderRec(1).ENTRYDATE)
''          Print #RptHandle, Dash$
''          LineCnt = LineCnt + 3
''        End If
''      End If
''      If LineCnt >= MaxLines Then
''        Print #RptHandle, FF$
''        GoSub PrintReadHeading
''      End If
''OpenRptSkip:
''      'ShowPctComp lcnt&, IdxNumOfRecs
''    Next
''    Print #RptHandle, FF$
''  ElseIf graphicflag = True Then 'do the graphics thing
''    For lcnt& = 1 To IdxNumOfRecs
''      FrmShowPctComp.ShowPctComp lcnt&, IdxNumOfRecs
''      If FrmShowPctComp.Out Then
''        Close
''        Unload FrmShowPctComp
''        ActivateControls Me
''        GoTo ExitHere
''      End If
''
''      Get #UBCustF, IdxBuff(lcnt&).RecNum, UBCustRec(1)
''      If UBCustRec(1).DelFlag <> 0 Then
''        GoTo OpenRptSkip2
''      End If
''
''      If UBCustRec(1).WOLastTrans > 0 Then
''        Get #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
''        If WorkOrderRec(1).CompletedDate <= 0 Then
''          ToPrint$ = Using("######", IdxBuff(lcnt&).RecNum) + "~"
''          ToPrint$ = ToPrint$ + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + QPTrim$(UBCustRec(1).CustName)
''          ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ServAddr) + "~" + Str(UBCustRec(1).WOLastTrans) + "~" + Num2Date$(WorkOrderRec(1).ENTRYDATE)
''          Print #RptHandle, ToPrint$
''          ToPrint$ = ""
''        End If
''      End If
''OpenRptSkip2:
''    Next
''  End If
''  Close
''  Erase IdxBuff
''  Header$ = "Open Work Orders Report"
''  'PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
''  If graphicflag = False Then
''    ViewPrint ReportFile$, Header$
''  Else
''    Load frmLoadingRpt
''    frmLoadingRpt.setwherefrom frmUBWorkOrderMenu
''    ARptWorkOrdersOpen.txtDate = Now
''    ARptWorkOrdersOpen.txtTown = TOWNNAME$
''    ARptWorkOrdersOpen.Title = Header$
''    ARptWorkOrdersOpen.GetName ReportFile$
''    ARptWorkOrdersOpen.startrpt
''  End If
''  Exit Sub
''
''PrintReadHeading:
''  Page = Page + 1
''  Print #RptHandle, Tab(30); "Open Work Order Report"
''  Print #RptHandle, "Date: "; Date$; Tab(70); "Page #"; Page
''  Print #RptHandle, "Acct No.   Location        Customer Name"
''  Print #RptHandle, "Service Address                               Work Order #      Order Date"
''  Print #RptHandle, String$(80, "=")
''  LineCnt = 5
''
''  Return
''
''ExitHere:
''  Close
''  Erase IdxBuff
''End Sub
''
