VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmUserSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Setup/Account Structure"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmUserSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAPChk 
      Height          =   405
      Left            =   7035
      TabIndex        =   15
      Top             =   3960
      Width           =   3735
      _Version        =   196608
      _ExtentX        =   6588
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
      EditMarginLeft  =   9
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmUserSetup.frx":08CA
   End
   Begin LpLib.fpCombo fpcboCRBank 
      Height          =   330
      Left            =   7680
      TabIndex        =   19
      Top             =   6240
      Width           =   3105
      _Version        =   196608
      _ExtentX        =   5477
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      EditMarginLeft  =   9
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmUserSetup.frx":0C89
   End
   Begin LpLib.fpCombo fpcboCDBank 
      Height          =   330
      Left            =   7680
      TabIndex        =   20
      Top             =   6570
      Width           =   3105
      _Version        =   196608
      _ExtentX        =   5477
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      EditMarginLeft  =   9
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmUserSetup.frx":1048
   End
   Begin LpLib.fpCombo fpcboChkBank 
      Height          =   330
      Left            =   7680
      TabIndex        =   21
      Top             =   6915
      Width           =   3105
      _Version        =   196608
      _ExtentX        =   5477
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      EditMarginLeft  =   9
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmUserSetup.frx":1407
   End
   Begin VB.ComboBox cboDupInv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      ItemData        =   "frmUserSetup.frx":17C6
      Left            =   10152
      List            =   "frmUserSetup.frx":17C8
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5496
      Width           =   612
   End
   Begin VB.ComboBox cboPSLFlag 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      ItemData        =   "frmUserSetup.frx":17CA
      Left            =   10152
      List            =   "frmUserSetup.frx":17CC
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5112
      Width           =   612
   End
   Begin VB.ComboBox cboPOStop 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      ItemData        =   "frmUserSetup.frx":17CE
      Left            =   10152
      List            =   "frmUserSetup.frx":17D0
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4728
      Width           =   612
   End
   Begin VB.ComboBox txtCDActive 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      ItemData        =   "frmUserSetup.frx":17D2
      Left            =   3840
      List            =   "frmUserSetup.frx":17D4
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   6648
      Width           =   612
   End
   Begin EditLib.fpLongInteger txtDeptCode 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      ToolTipText     =   "Enter the Department Code(Valid Choices 1, 2 or 3)"
      Top             =   3720
      Width           =   615
      _Version        =   196608
      _ExtentX        =   1080
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
      ButtonStyle     =   1
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
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
      Text            =   ""
      MaxValue        =   "3"
      MinValue        =   "1"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpLongInteger txtFundLen 
      Height          =   372
      Left            =   5640
      TabIndex        =   1
      ToolTipText     =   "Enter the Number of Digits in the Fund Code(Valid Choices 1, 2, 3 or 4)."
      Top             =   1920
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1080
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
      ButtonStyle     =   1
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
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
      MaxValue        =   "4"
      MinValue        =   "1"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtTotAcctLen 
      Height          =   372
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Total Number of Digits in GL Account is Calculated from 3 Entries Listed Above."
      Top             =   3192
      Width           =   372
      _Version        =   196608
      _ExtentX        =   656
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
      ButtonStyle     =   0
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   2
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
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtUserName 
      Height          =   372
      Left            =   5640
      TabIndex        =   0
      ToolTipText     =   "Enter the Name of the Town."
      Top             =   1464
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
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
      ButtonStyle     =   0
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
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483629
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
      MaxLength       =   30
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
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   1
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
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
      Left            =   8664
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7488
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Top             =   8265
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   661
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
            TextSave        =   "1:54 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
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
      Height          =   492
      Left            =   10320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7512
      Width           =   1332
   End
   Begin EditLib.fpLongInteger txtAcctLen 
      Height          =   372
      Left            =   5640
      TabIndex        =   2
      ToolTipText     =   "Enter the Number of Digits in Account(Valid Choices 1, 2, 3, 4, 5 or 6)."
      Top             =   2328
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1080
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
      ButtonStyle     =   1
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483643
      InvalidOption   =   2
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
      MaxValue        =   "6"
      MinValue        =   "1"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpLongInteger txtDetLen 
      Height          =   372
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "Enter the Number of Digits in the Department(Valid Choices 1, 2, 3 or 4)."
      Top             =   2760
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1080
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
      ButtonStyle     =   1
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483643
      InvalidOption   =   2
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
      MaxValue        =   "4"
      MinValue        =   "1"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtCashAcct 
      Height          =   372
      Left            =   3840
      TabIndex        =   6
      Top             =   4176
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtCRCashAcct 
      Height          =   372
      Left            =   3840
      TabIndex        =   8
      Top             =   4992
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtCDCashAcct 
      Height          =   372
      Left            =   3840
      TabIndex        =   9
      Top             =   5400
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtFBAcct 
      Height          =   372
      Left            =   3840
      TabIndex        =   10
      Top             =   5832
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtCDDue 
      Height          =   372
      Left            =   3840
      TabIndex        =   14
      Top             =   7464
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtEncAcct 
      Height          =   372
      Left            =   3840
      TabIndex        =   11
      Top             =   6240
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtCDCash 
      Height          =   372
      Left            =   3840
      TabIndex        =   13
      Top             =   7056
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtAPAcct 
      Height          =   372
      Left            =   3840
      TabIndex        =   7
      Top             =   4584
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
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
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Line Line3 
      X1              =   6024
      X2              =   6024
      Y1              =   6216
      Y2              =   7224
   End
   Begin VB.Line Line2 
      X1              =   7272
      X2              =   10848
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Defaults"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   6000
      TabIndex        =   51
      Top             =   6000
      Width           =   2004
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AP Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   6192
      TabIndex        =   50
      Top             =   6912
      Width           =   1356
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   5976
      TabIndex        =   49
      Top             =   6288
      Width           =   1572
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Disb"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   6120
      TabIndex        =   48
      Top             =   6576
      Width           =   1428
   End
   Begin VB.Shape Shape3 
      Height          =   1404
      Left            =   6024
      Top             =   4536
      Width           =   4836
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   6120
      TabIndex        =   47
      Top             =   4584
      Width           =   2004
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duplicate Inv Numbers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   7152
      TabIndex        =   46
      Top             =   5520
      Width           =   2772
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PSL Default "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   7944
      TabIndex        =   45
      Top             =   5184
      Width           =   2052
   End
   Begin VB.Label LblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   360
      TabIndex        =   44
      Top             =   600
      Width           =   1668
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   7536
      TabIndex        =   43
      Top             =   4800
      Width           =   2412
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AP Check Format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   396
      Left            =   5712
      TabIndex        =   42
      Top             =   3648
      Width           =   2148
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmUserSetup.frx":17D6
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1452
      Index           =   1
      Left            =   6480
      TabIndex        =   41
      Top             =   2064
      Width           =   3492
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Due Central Depository "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   360
      TabIndex        =   40
      Top             =   4176
      Width           =   3372
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Payable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   1560
      TabIndex        =   39
      Top             =   4584
      Width           =   2052
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   1920
      TabIndex        =   38
      Top             =   5880
      Width           =   1692
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipts Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   1440
      TabIndex        =   37
      Top             =   4992
      Width           =   2172
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Disbursements Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   600
      TabIndex        =   36
      Top             =   5400
      Width           =   3012
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Setup and Account Structure"
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
      Left            =   3120
      TabIndex        =   24
      Top             =   600
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2160
      Top             =   360
      Width           =   7932
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   972
      Left            =   2160
      Top             =   240
      Width           =   7932
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000013&
      X1              =   696
      X2              =   10896
      Y1              =   3624
      Y2              =   3624
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   1560
      TabIndex        =   34
      Top             =   3768
      Width           =   2052
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Central Depository Active"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   840
      TabIndex        =   33
      Top             =   6696
      Width           =   2772
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Central Depository Due To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   720
      TabIndex        =   32
      Top             =   7512
      Width           =   2892
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Central Depository Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   840
      TabIndex        =   31
      Top             =   7104
      Width           =   2772
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Encumbrance Reserve"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   1080
      TabIndex        =   30
      Top             =   6288
      Width           =   2532
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   4560
      TabIndex        =   29
      Top             =   1464
      Width           =   852
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Length of Account Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   840
      TabIndex        =   28
      Top             =   3216
      Width           =   4572
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Digits in Detail Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Left            =   1560
      TabIndex        =   27
      Top             =   2784
      Width           =   3852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Digits in Account Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   336
      Left            =   1200
      TabIndex        =   26
      Top             =   2352
      Width           =   4212
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Digits in Fund Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   372
      Index           =   0
      Left            =   1560
      TabIndex        =   25
      Top             =   1896
      Width           =   3852
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "&Print Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmUserSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLSetup As GLSetupRecType
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Function Oktogo()
  If txtCDActive.ListIndex = 1 Then
    If txtCDCash = "" Then
      MsgBox "You Must Enter The Central Depository Cash Account Information If Central Depository Is Active.", vbOKOnly, "Required Field"
      txtCDCash.SetFocus
      Oktogo = False
      Exit Function
    End If
    If txtCDDue = "" Then
      MsgBox "You Must Enter The Central Depository Due Account Information If Central Depository Is Active.", vbOKOnly, "Required Field"
      txtCDDue.SetFocus
      Oktogo = False
      Exit Function
    End If
  End If
  If fpcboAPChk.ListIndex = -1 Then
    MsgBox "You Must Select An AP Check Format.", vbOKOnly, "AP Check"
    Oktogo = False
    fpcboAPChk.SetFocus
    Exit Function
  End If
  Oktogo = True
End Function
Private Sub cmdExit_Click()
  If Oktogo = True Then
  If MsgBox("If You Have Made Changes They Will Not Be Saved." & vbCrLf & "'OK' to Exit, 'Cancel' to Remain on Setup Screen.", vbOKCancel, "Exit Without Saving") = vbOK Then
    Call MainLog("Exited User Setup Screen ")
    frmGLConfigUtilMenu.Show
    Unload frmUserSetup
  Else  'Stay on screen
  End If
  End If
End Sub
Private Sub cmdSave_Click()
  Dim SetupFile As Integer
  If Oktogo = True Then
    OpenSetupFile SetupFile
    'Get SetupFile, 1, GLSetup
    Call ValidationRules
    GLSetup.UserName = frmUserSetup.txtUserName
    GLSetup.FundLen = frmUserSetup.txtFundLen
    GLSetup.AcctLen = frmUserSetup.txtAcctLen
    GLSetup.DetLen = frmUserSetup.txtDetLen
    GLSetup.TotAcctLen = frmUserSetup.txtTotAcctLen
    GLSetup.DeptCode = frmUserSetup.txtDeptCode
    GLSetup.CashAcct = frmUserSetup.txtCashAcct
    GLSetup.APAcct = frmUserSetup.txtAPAcct
    GLSetup.CRCashAcct = frmUserSetup.txtCRCashAcct
    GLSetup.CDCashAcct = frmUserSetup.txtCDCashAcct
    GLSetup.FBAcct = frmUserSetup.txtFBAcct
    GLSetup.EncAcct = frmUserSetup.txtEncAcct
    GLSetup.CDCash = frmUserSetup.txtCDCash
    GLSetup.CDDue = frmUserSetup.txtCDDue
    GLSetup.CDActive = frmUserSetup.txtCDActive
    If cboPOStop.ListIndex = 1 Then
      GLSetup.POStop = True
    Else
      GLSetup.POStop = False
    End If
    If cboPSLFlag.ListIndex = 1 Then
      GLSetup.PSLFlag = 1
    Else
      GLSetup.PSLFlag = 0
    End If
    If cboDupInv.ListIndex = 1 Then
      GLSetup.DupInvFlag = 1
    Else
      GLSetup.DupInvFlag = 0
    End If
'(*)#*)@#*)@#*)#@*@)#*)@# PUT Version in !!!!!!!!!!!!!!!!!
      GLSetup.ChkVer = "V205"
      GLSetup.Pad = ""
'&(*&$(*&$(#&#$($#&(#*$&(*&(*&(**&^*&^
    fpcboAPChk.Col = 0
    GLSetup.APChkCode = fpcboAPChk.ColText
    If fpcboCRBank.ListIndex <> -1 Then
      fpcboCRBank.Col = 0
      GLSetup.CRBank = fpcboCRBank.ColText
    Else
      GLSetup.CRBank = 0
    End If
    If fpcboCDBank.ListIndex <> -1 Then
      fpcboCDBank.Col = 0
      GLSetup.CDBank = fpcboCDBank.ColText
    Else
      GLSetup.CDBank = 0
    End If
    If fpcboChkBank.ListIndex <> -1 Then
      fpcboChkBank.Col = 0
      GLSetup.ChkBank = fpcboChkBank.ColText
    Else
      GLSetup.ChkBank = 0
    End If

    Put SetupFile, 1, GLSetup
    Close SetupFile
    MsgBox "Your Information has been saved.", vbOKOnly
    Call MainLog("Saved Setup ")
    frmGLConfigUtilMenu.Show
    Unload frmUserSetup
    Call MainLog("Exit Setup ")
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
      SendKeys "%S"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Dim SetupFile As Integer
  Dim acctmsk As String, fundmsk As String
  Dim detmsk As String
  Dim Num As String
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  txtCDActive.AddItem "N"
  txtCDActive.AddItem "Y"
  cboPOStop.AddItem "N"
  cboPOStop.AddItem "Y"
  cboPSLFlag.AddItem "N"
  cboPSLFlag.AddItem "Y"
  cboDupInv.AddItem "N"
  cboDupInv.AddItem "Y"
  fpcboAPChk.InsertRow = "1" & Chr$(9) & "Blank Top Stub 39 Line"
  fpcboAPChk.InsertRow = "2" & Chr$(9) & "Blank Top Stub 42 Line"
  fpcboAPChk.InsertRow = "3" & Chr$(9) & "Product Code 928"
  fpcboAPChk.InsertRow = "4" & Chr$(9) & "Middle Check Laser"
  fpcboAPChk.InsertRow = "5" & Chr$(9) & "Top Check Laser"
  fpcboAPChk.InsertRow = "6" & Chr$(9) & "Divided Top Stub 42 Line"
  If Exist("GLBank.dat") Then
    GetBankList fpcboCRBank
    GetBankList fpcboCDBank
    GetBankList fpcboChkBank
  Else
    fpcboCRBank.Enabled = False
    fpcboCDBank.Enabled = False
    fpcboChkBank.Enabled = False
  End If
  If Exist("GLSetup.DAT") Then
    OpenSetupFile SetupFile
    Get SetupFile, 1, GLSetup
    Close SetupFile
    LblVersion.Caption = GLSetup.ChkVer
    fundmsk = String(GLSetup.FundLen, "#")
    acctmsk = String(GLSetup.AcctLen, "#") 'bases length of acct mask on what was previously entered
    detmsk = String(GLSetup.DetLen, "#") 'same thing as above but for detail
    frmUserSetup.txtUserName = Trim(GLSetup.UserName)
    frmUserSetup.txtFundLen = GLSetup.FundLen
    frmUserSetup.txtAcctLen = GLSetup.AcctLen
    frmUserSetup.txtDetLen = GLSetup.DetLen
    frmUserSetup.txtTotAcctLen = GLSetup.TotAcctLen
    frmUserSetup.txtDeptCode = GLSetup.DeptCode
    If GLSetup.CDActive = "Y" Then 'for combo box give index value of Y or N
      txtCDActive.ListIndex = 1
    Else
      txtCDActive.ListIndex = 0
    End If
    Select Case GLSetup.APChkCode
      Case 1:
        fpcboAPChk.ListIndex = 0
      Case 2:
        fpcboAPChk.ListIndex = 1
      Case 3:
        fpcboAPChk.ListIndex = 2
      Case 4:
        fpcboAPChk.ListIndex = 3
      Case 5:
        fpcboAPChk.ListIndex = 4
      Case 6:
        fpcboAPChk.ListIndex = 5
      Case Else
        fpcboAPChk.ListIndex = 0
    End Select
    If GLSetup.POStop = True Then
      cboPOStop.ListIndex = 1
    Else
      cboPOStop.ListIndex = 0
    End If
    If GLSetup.TotAcctLen > 0 Then
      txtFundLen.Enabled = False
      txtAcctLen.Enabled = False
      txtDetLen.Enabled = False
      txtTotAcctLen.Enabled = False
    End If
    If GLSetup.PSLFlag = 1 Then
      cboPSLFlag.ListIndex = 1
    Else
      cboPSLFlag.ListIndex = 0
    End If
    If GLSetup.DupInvFlag = 1 Then
      cboDupInv.ListIndex = 1
    Else
      cboDupInv.ListIndex = 0
    End If
     'Set Masks for Account fields
    txtCashAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtAPAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtCRCashAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtCDCashAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtFBAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtEncAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtCDCash.Mask = (fundmsk & "-" & acctmsk & "-" & detmsk)
    txtCDDue.Mask = (fundmsk & "-" & acctmsk & "-")
  'trim fields from file so the validate will be correct
    frmUserSetup.txtCashAcct = Trim(GLSetup.CashAcct)
    frmUserSetup.txtAPAcct = Trim(GLSetup.APAcct)
    frmUserSetup.txtCRCashAcct = Trim(GLSetup.CRCashAcct)
    frmUserSetup.txtCDCashAcct = Trim(GLSetup.CDCashAcct)
    frmUserSetup.txtFBAcct = Trim(GLSetup.FBAcct)
    frmUserSetup.txtEncAcct = Trim(GLSetup.EncAcct)
    frmUserSetup.txtCDCash = Trim(GLSetup.CDCash)
    frmUserSetup.txtCDDue = Trim(GLSetup.CDDue)
  Else
    txtCDActive.ListIndex = 0
    txtDeptCode = 1
  End If
  If GLSetup.CDBank > 0 Then
    fpcboCDBank.SearchText = Trim(GLSetup.CDBank)
    fpcboCDBank.Action = 0
    If fpcboCDBank.SearchIndex <> -1 Then
      fpcboCDBank.ListIndex = fpcboCDBank.SearchIndex
    End If
  End If
  If GLSetup.CRBank > 0 Then
    fpcboCRBank.SearchText = Trim(GLSetup.CRBank)
    fpcboCRBank.Action = 0
    If fpcboCRBank.SearchIndex <> -1 Then
      fpcboCRBank.ListIndex = fpcboCRBank.SearchIndex
    End If
  End If
  If GLSetup.ChkBank > 0 Then
    fpcboChkBank.SearchText = Trim(GLSetup.ChkBank)
    fpcboChkBank.Action = 0
    If fpcboChkBank.SearchIndex <> -1 Then
      fpcboChkBank.ListIndex = fpcboChkBank.SearchIndex
    End If
  End If

  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = GLSetup.UserName
  Me.HelpContextID = hlpAccountSetup
End Sub

Private Sub fpcboAPChk_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAPChk.ListDown = True
  End If
  If fpcboAPChk.ListDown <> True Then
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

Private Sub fpcboAPChk_LostFocus()
  fpcboAPChk.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboCDBank_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCDBank.ListDown = True
  End If
  If fpcboCDBank.ListDown <> True Then
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

Private Sub fpcboCDBank_LostFocus()
  fpcboCDBank.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboCRBank_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCRBank.ListDown = True
  End If
  If fpcboCRBank.ListDown <> True Then
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

Private Sub fpcboCRBank_LostFocus()
  fpcboCRBank.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboChkBank_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboChkBank.ListDown = True
  End If
  If fpcboChkBank.ListDown <> True Then
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

Private Sub fpcboChkBank_LostFocus()
  fpcboChkBank.Action = ActionClearSearchBuffer
End Sub

Private Sub mnuExit_Click()
'Call ValidationRules
  cmdExit_Click
End Sub
Private Sub mnuPrint_Click()
  Print
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtAcctLen_Change()
  SetAcctLen
End Sub
Private Function ValidationRules(Optional ByVal strField As String = "") As Boolean
  Dim blnReturn As Boolean
  Dim ctlLoop As Control
  Dim strMessage As String
  blnReturn = False
  strMessage = ""
  '===    Loop through all controls...
  For Each ctlLoop In Me.Controls
    If (Left(ctlLoop.Name, Len(strField)) = strField) Then
      Select Case ctlLoop.Name
        Case "txtUserName":
          If ((Len(txtUserName) = 0) Or (txtUserName = "  ")) Then
            strMessage = strMessage & vbCrLf & "User Name Can Not Be Left Blank. Previous Value Restored."
            txtUserName = GLSetup.UserName
            blnReturn = True
          End If
        Case "txtFundLen":
          Dim FundLen As Variant
          FundLen = Val(txtFundLen)
          If ((FundLen <= 0) Or (FundLen > 4)) Then
            strMessage = strMessage & vbCrLf & "Invalid Fund Code Value. Previous Value Restored."
              txtFundLen = GLSetup.FundLen
              blnReturn = True
          End If
        Case "txtAcctLen":
          Dim AcctLen As Variant
          AcctLen = Val(txtAcctLen)
          If ((AcctLen <= 0) Or (AcctLen > 6)) Then
            strMessage = strMessage & vbCrLf & "Invalid Account Code Value. Previous Value Restored."
            txtAcctLen = GLSetup.AcctLen
            blnReturn = True
          End If
        Case "txtDetLen":
          'Dim DetLen As Integer
          Dim DetLen As Variant
          DetLen = Val(txtDetLen)
          If ((DetLen <= 0) Or (DetLen > 4)) Then
            strMessage = strMessage & vbCrLf & "Invalid Detail Code Value. Previous Value Restored."
            txtDetLen = GLSetup.DetLen
            blnReturn = True
          End If
        Case "txtDeptCode":
          Dim DeptCode As Variant
          DeptCode = Val(txtDeptCode)
          If ((DeptCode <= 0) Or (DeptCode > 3)) Then
            strMessage = strMessage & vbCrLf & "Invalid Department Code. Default of 1 Will Be Restored if No Prior Value."
            blnReturn = True
            If Not IsNull(GLSetup.DeptCode) Then
              txtDeptCode = GLSetup.DeptCode
            Else
              txtDeptCode = 1
            End If
          End If
        Case Else:
        '=== No validation needed for this control...
        End Select
    End If
    If blnReturn = True Then
        Exit For
    End If
  Next ctlLoop
  If (Len(strMessage) > 0) Then
    MsgBox strMessage, vbOKOnly + vbCritical, "Errors found."
  End If
  ValidationRules = blnReturn
End Function
Private Sub txtAcctLen_Validate(Cancel As Boolean)
  Cancel = ValidationRules("txtAcctLen")
End Sub
Private Sub txtAPAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtAPAcct_LostFocus()
  GetAcctMsk txtAPAcct
End Sub
Private Sub txtCashAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtCashAcct_LostFocus()
  GetAcctMsk txtCashAcct
End Sub

Private Sub txtCDCash_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtCDCash_LostFocus()
  Dim acctmsk As String, detmsk As String, fundmsk As String
  Dim Num As String
  fundmsk = String(txtFundLen, "#")
  acctmsk = String(txtAcctLen, "#")
  detmsk = String(txtDetLen, "#")
  Num = Trim(txtCDCash)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(txtFundLen) + Val(txtAcctLen) + Val(txtDetLen) + 2) Or InstrCount(Num, "-") <> 2 Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtCDCash.Mask = (fundmsk & "-" & acctmsk & "-" & detmsk)
      txtCDCash.SetFocus
    
    End If
  End If
End Sub
Private Sub txtCDCashAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtCDCashAcct_LostFocus()
  GetAcctMsk txtCDCashAcct
End Sub
Private Sub txtCDDue_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtCDDue_LostFocus()
  Dim acctmsk As String, detmsk As String, fundmsk As String
  Dim Num As String
  fundmsk = String(txtFundLen, "#")
  acctmsk = String(txtAcctLen, "#")
  detmsk = String(txtDetLen, "#")
  Num = Trim(txtCDDue)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(txtFundLen) + Val(txtAcctLen) + 2) Or InstrCount(Num, "-") <> 2 Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtCDDue.Mask = (fundmsk & "-" & acctmsk & "-")
      txtCDDue.SetFocus
    
    End If
  End If
End Sub
Private Sub txtCRCashAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtCRCashAcct_LostFocus()
  GetAcctMsk txtCRCashAcct
End Sub
Private Sub txtDeptCode_Validate(Cancel As Boolean)
  Cancel = ValidationRules("txtDeptCode")
End Sub
Private Sub txtDetLen_Change()
  SetAcctLen
End Sub
Private Sub txtDetLen_Validate(Cancel As Boolean)
  Cancel = ValidationRules("txtDetLen")
End Sub
Private Sub txtEncAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtEncAcct_LostFocus()
  GetAcctMsk txtEncAcct
End Sub
Private Sub txtFBAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtFBAcct_LostFocus()
  GetAcctMsk txtFBAcct
End Sub
Private Sub txtFundLen_Change()
  SetAcctLen
End Sub
Private Sub SetAcctLen()
  frmUserSetup.txtTotAcctLen = Val(frmUserSetup.txtAcctLen) + Val(frmUserSetup.txtDetLen) + Val(frmUserSetup.txtFundLen)
End Sub
Private Sub txtFundLen_Validate(Cancel As Boolean)
  Cancel = ValidationRules("txtFundLen")
End Sub


Private Sub txtTotAcctLen_LostFocus()
'once move off totacctlen set up masks for all acct fields
  Dim acctmsk As String, fundmsk As String
  Dim detmsk As String
  'If txtDeptCode = "" Then
    fundmsk = String(txtFundLen, "#")
    acctmsk = String(txtAcctLen, "#")
    detmsk = String(txtDetLen, "#")
    txtCashAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtAPAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtCRCashAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtCDCashAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtFBAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtEncAcct.Mask = ("-" & acctmsk & "-" & detmsk)
    txtCDCash.Mask = (fundmsk & "-" & acctmsk & "-" & detmsk)
    txtCDDue.Mask = (fundmsk & "-" & acctmsk & "-")
    'txtDeptCode = 1
  'End If
End Sub

Private Sub txtUserName_Validate(Cancel As Boolean)
  Cancel = ValidationRules("txtUserName")
End Sub
Private Sub GetAcctMsk(txtField As fpMask)
'When new setup-create account masks on the fly after fill in lengths
  Dim acctmsk As String
  Dim detmsk As String
  Dim Num As String
  
  acctmsk = String(txtAcctLen, "#")
  detmsk = String(txtDetLen, "#")
  Num = Trim(txtField)
  If (Len(Num)) > 1 Then
    If (Len(Num)) <> (Val(txtAcctLen) + Val(txtDetLen) + 2) Or InstrCount(Num, "-") <> 2 Then
      MsgBox "Invalid Code.", vbOKOnly, "Invalid Data!"
      txtField.Mask = ("-" & acctmsk & "-" & detmsk)
      txtField.SetFocus
    End If
  Else
    txtField.Mask = ("-" & acctmsk & "-" & detmsk)
  End If
End Sub
  
   
  
