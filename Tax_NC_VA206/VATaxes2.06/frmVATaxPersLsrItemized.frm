VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPersLsrItemized 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Laser Itemized Tax Bill"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPersLsrItemized.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboLogo 
      Height          =   312
      Left            =   9648
      TabIndex        =   2
      Top             =   1776
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   550
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
      ColDesigner     =   "frmVATaxPersLsrItemized.frx":08CA
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   9862
      TabIndex        =   16
      Top             =   7752
      Width           =   1332
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "F5 &Test Print"
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
      Left            =   7294
      TabIndex        =   15
      Top             =   7752
      Width           =   2250
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
      Left            =   5626
      TabIndex        =   14
      Top             =   7752
      Width           =   1332
   End
   Begin EditLib.fpText fpTxtHead1 
      Height          =   300
      Left            =   630
      TabIndex        =   0
      Top             =   1800
      Width           =   5916
      _Version        =   196608
      _ExtentX        =   10435
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtHead3 
      Height          =   252
      Left            =   2676
      TabIndex        =   10
      Top             =   6276
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   40
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
   Begin EditLib.fpText fpTxtHead2 
      Height          =   300
      Left            =   630
      TabIndex        =   1
      Top             =   2112
      Width           =   5916
      _Version        =   196608
      _ExtentX        =   10435
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpText fpTxtOpt3 
      Height          =   252
      Left            =   636
      TabIndex        =   13
      Top             =   7200
      Width           =   6468
      _Version        =   196608
      _ExtentX        =   11409
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   75
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
   Begin EditLib.fpText fpTxtOpt1 
      Height          =   252
      Left            =   660
      TabIndex        =   8
      Top             =   5376
      Width           =   6468
      _Version        =   196608
      _ExtentX        =   11409
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   75
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
   Begin EditLib.fpText fpTxtOpt2 
      Height          =   252
      Left            =   636
      TabIndex        =   9
      Top             =   5928
      Width           =   4476
      _Version        =   196608
      _ExtentX        =   7895
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   45
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   228
      Index           =   0
      Left            =   636
      TabIndex        =   3
      Top             =   2832
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   228
      Index           =   1
      Left            =   636
      TabIndex        =   4
      Top             =   3072
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   228
      Index           =   2
      Left            =   636
      TabIndex        =   5
      Top             =   3312
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   228
      Index           =   3
      Left            =   636
      TabIndex        =   6
      Top             =   3552
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   228
      Index           =   4
      Left            =   636
      TabIndex        =   7
      Top             =   3792
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtHead4 
      Height          =   252
      Left            =   2676
      TabIndex        =   11
      Top             =   6528
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   40
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
   Begin EditLib.fpText fptxtHead5 
      Height          =   252
      Left            =   2676
      TabIndex        =   12
      Top             =   6792
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   40
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This area is used to display all tax data for the individual taxpayer."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   1
      Left            =   552
      TabIndex        =   27
      Top             =   4560
      Width           =   4980
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   936
      Left            =   2316
      TabIndex        =   26
      Top             =   6216
      Width           =   348
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return Address:"
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
      Height          =   336
      Index           =   0
      Left            =   432
      TabIndex        =   25
      Top             =   6516
      Width           =   1740
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Town Logo on Bill:"
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
      Height          =   336
      Left            =   7104
      TabIndex        =   24
      Top             =   1800
      Width           =   2436
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(7 Lines Available)"
      Height          =   276
      Left            =   4632
      TabIndex        =   23
      Top             =   2616
      Width           =   1416
   End
   Begin VB.Line Line4 
      X1              =   11460
      X2              =   11460
      Y1              =   1404
      Y2              =   7560
   End
   Begin VB.Line Line3 
      X1              =   192
      X2              =   11472
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   294
      X2              =   294
      Y1              =   1356
      Y2              =   7212
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Letter Laser Bill Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   582
      TabIndex        =   22
      Top             =   1176
      Width           =   4860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Personal Itemized Bill Letter Printing"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   576
      Width           =   5976
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   516
      Left            =   2820
      Top             =   480
      Width           =   6252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Paragraph prints above Customer Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Index           =   0
      Left            =   552
      TabIndex        =   20
      Top             =   2592
      Width           =   3372
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Heading Lines 1 and 2:"
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
      Height          =   336
      Index           =   5
      Left            =   186
      TabIndex        =   19
      Top             =   1476
      Width           =   2388
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prints on the Top of Return Stub."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   6
      Left            =   552
      TabIndex        =   18
      Top             =   5160
      Width           =   2496
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prints below Totals Left  Side of Return Stub."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   336
      Index           =   7
      Left            =   588
      TabIndex        =   17
      Top             =   5676
      Width           =   4536
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4110
      X2              =   11286
      Y1              =   1356
      Y2              =   1356
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   636
      Left            =   2820
      Top             =   360
      Width           =   6252
   End
End
Attribute VB_Name = "frmVATaxPersLsrItemized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private Temp_Class As Resize_Class
  Dim Over As clsTextBoxOverRider
  Dim TownName$
Private Sub cmdExit_Click()
  Unload Me
  DoEvents
End Sub

Private Sub cmdSave_Click()
  Call SaveTxDef
End Sub

Private Sub cmdTest_Click()
  Call TestLtrBillPrint
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPersLsrItemized.")
      Call Terminate
      End
    End If
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
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdSave_Click
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      Call cmdTest_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub fpTxtHead1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtHead2.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpcboLogo.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt1.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtHead1.SetFocus
      KeyCode = 0
    End If
  End If
End Sub


Private Sub fpcboLogo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboLogo.ListDown = True
  End If
  If fpcboLogo.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fptxtPgph(0).SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpTxtHead2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpTxtOpt1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt2.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpcboLogo.SetFocus
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fptxtPgph_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
'    If Index = 7 Then
'      fpTxtOpt5.SetFocus
'    Else
      fptxtPgph(Index + 1).SetFocus
'    End If
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
'      If Index = 0 Then
'        fpTxtOpt4.SetFocus
'      Else
        fptxtPgph(Index - 1).SetFocus
'      End If
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpTxtOpt2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtHead4.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt1.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt3_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtHead1.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtHead5.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt4_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtPgph(0).SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt3.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead4_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtHead5.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt2.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead5_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt3.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtHead4.SetFocus
      KeyCode = 0
    End If
  End If
End Sub

Private Sub TestLtrBillPrint()
  Dim RptHandle As Integer
  Dim ReportFile As String
  Dim ThisOpt1Desc$
  Dim ThisOpt2Desc$
  Dim ThisOpt3Desc$
  Dim dlm$
  Dim Head1$, Head2$, Head3$, Head4$, Head5$
  Dim Para0$, Para1$, Para2$, Para3$, Para4$
  Dim Opttxt1$, Opttxt2$, Opttxt3$
  Dim UseLogo As Integer
  
  On Error GoTo ERRORSTUFF
  If fpcboLogo.Text = "Yes" Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxLaserPersItemized.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxLaserPersItemized.Image1.Visible = True
    End If
    UseLogo = 1
  Else
    UseLogo = 0
  End If
  
  If QPTrim$(fpTxtHead1.Text) <> "" Then
    Head1$ = fpTxtHead1.Text
  Else
    Head1$ = "Town of Anytown"
  End If
  
  If QPTrim$(fpTxtHead2.Text) <> "" Then
    Head2$ = fpTxtHead2.Text
  Else
    Head2$ = "2006 Property Tax Notice"
  End If
  
  If QPTrim$(fptxtHead3.Text) <> "" Then
    Head3$ = fptxtHead3.Text
  Else
    Head3$ = "Town of Anytown"
  End If
  
  If QPTrim$(fptxtHead4.Text) <> "" Then
    Head4$ = fptxtHead4.Text
  Else
    Head4$ = "PO Box 1190"
  End If
  
  
  If QPTrim$(fptxtHead5.Text) <> "" Then
    Head5$ = fptxtHead5.Text
  Else
    Head5$ = "Anytown, VA 55555-5555"
  End If
  
  If QPTrim$(fptxtPgph(0).Text) <> "" Then
    Para0 = fptxtPgph(0).Text
  Else
    Para0 = "Taxes are due and payable upon receipt of this notice. Interest begins to accrue on unpaid taxes at the rate of 2%"
  End If
  
  If QPTrim$(fptxtPgph(1).Text) <> "" Then
    Para1 = fptxtPgph(1).Text
  Else
    Para1 = " on January 6th. Beginning on the first day of February and continuing thereafter, interest accrues at the rate of "
  End If
  
  If QPTrim$(fptxtPgph(2).Text) <> "" Then
    Para2 = fptxtPgph(2).Text
  Else
    Para2 = "0.75% per month until all taxes, interest and penalties are paid. Please note that since we bill semi-annually that "
  End If
  
  If QPTrim$(fptxtPgph(3).Text) <> "" Then
    Para3 = fptxtPgph(3).Text
  Else
    Para3 = "property values are displayed accordingly. Please call 555-5555 at the tax office with any questions. Thank you."
  End If
  
  If QPTrim$(fptxtPgph(4).Text) <> "" Then
    Para4 = fptxtPgph(4).Text
  Else
    Para4 = "John Smith, City Tax Collector"
  End If
  
  If QPTrim$(fpTxtOpt1.Text) <> "" Then
    Opttxt1$ = fpTxtOpt1.Text
  Else
    Opttxt1$ = "Please detach and return this portion with your payment."
  End If
  
  If QPTrim$(fpTxtOpt2.Text) <> "" Then
    Opttxt2$ = fpTxtOpt2.Text
  Else
    Opttxt2$ = "Please make checks payable to:"
  End If
  
  If QPTrim$(fpTxtOpt3.Text) <> "" Then
    Opttxt3$ = fpTxtOpt3.Text
  Else
    Opttxt3$ = "If a receipt is desired, send a self addressed stamped envelope."
  End If
  
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxPLsrItem.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
    '                         0                         1
    Print #RptHandle, "100"; dlm; "John J. Neighbor"; dlm;
    '                       2                    3
    Print #RptHandle, "123 Elm St"; dlm; "P.O. Box 1190"; dlm;
    '                       4                5
    Print #RptHandle, "Alltowns, VA"; dlm; "900"; dlm;
    '                   6               7
    Print #RptHandle, 20000; dlm; "Private Vehicle"; dlm;
    '                   8         9       10
    Print #RptHandle, 20000; dlm; 0; dlm; 0; dlm;
    '                    11       12        13
    Print #RptHandle, 20000; dlm; 35#; dlm; 15; dlm;
    '                   14     15      16
    Print #RptHandle, UseLogo; dlm; 0; dlm; 0; dlm;
    '                17       18       19
    Print #RptHandle, 0; dlm; 50; dlm; 15; dlm;
    '                  20        21      22
    Print #RptHandle, 0.25; dlm; 0; dlm; 0; dlm;
    '                 23        24       25
    Print #RptHandle, 0; dlm; 0.25; dlm; 0; dlm;
    '                  26       27        28
    Print #RptHandle, 0.25; dlm; 0; dlm; 0.25; dlm;
    '                 29     30       31
    Print #RptHandle, 0; dlm; 0; dlm; 0; dlm;
    ThisOpt1Desc = "STORM WATER"
    ThisOpt2Desc = ""
    ThisOpt3Desc = ""
    '                        32              33                 34               35             36               37
    Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; 125321; dlm; 55555 - 5555; dlm; UseLogo; dlm;
    '                  38          39           40
    Print #RptHandle, Head1; dlm; Head2; dlm; Opttxt1$; dlm;
    '                   41             42           43
    Print #RptHandle, Opttxt2$; dlm; Opttxt3$; dlm; ""; dlm;
    '                   44         45           46         47
    Print #RptHandle, Para0; dlm; Para1; dlm; Para2; dlm; Para3; dlm;
    '                   48
    Print #RptHandle, Para4; dlm;
    '                           49
    Print #RptHandle, "                        "; dlm;
    '                           50
    Print #RptHandle, "                        "; dlm;
    '                           51
    Print #RptHandle, "                        "; dlm; ""; dlm;
    '                   53
    Print #RptHandle, Head3; dlm;
    '                   54
    Print #RptHandle, Head4; dlm;
    '                   55
    Print #RptHandle, Head5; dlm;
    '                 56       57
    Print #RptHandle, ""; dlm; ""; dlm;
    '                     58                 59              60
    Print #RptHandle, "17626DDT55X"; dlm; "FORD/F150"; dlm; 20000; dlm;
    '                61      62      63       64      65      66         67
    Print #RptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 1; dlm; ""; dlm; "0.00"
  
  Close
  arVATaxLaserPersItemized.Show
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersLsrItemized", "TestLtrBillPrint", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub

Private Sub SaveTxDef()
  Dim TBRec As TxBillLaserItemized
  Dim TBHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  If fpcboLogo.Text = "Yes" Then
    If Not Exist("towntaxlogo.bmp") Then
      Call TaxMsg(600, "The town's logo cannot be displayed until it has been saved in the Citipak folder as 'towntaxlogo.bmp'. Please create and place this file in the Citipak folder. Call Southern Software at 1-800-842-8190 for assistance if needed.")
      fpcboLogo.Text = "No"
    End If
  End If
  
  OpenLaserPersItemized TBHandle
    TBRec.TxtHead1 = QPTrim(fpTxtHead1)
    TBRec.TxtHead2 = QPTrim(fpTxtHead2)
    TBRec.txtOpt1 = QPTrim(fpTxtOpt1)
    TBRec.TxtOpt2 = QPTrim(fpTxtOpt2)
    TBRec.TxtOpt3 = QPTrim(fpTxtOpt3)
    TBRec.txtPgph0 = QPTrim(fptxtPgph(0))
    TBRec.txtPgph1 = QPTrim(fptxtPgph(1))
    TBRec.txtPgph2 = QPTrim(fptxtPgph(2))
    TBRec.txtPgph3 = QPTrim(fptxtPgph(3))
    TBRec.txtPgph4 = QPTrim(fptxtPgph(4))
    TBRec.txtHead3 = QPTrim(fptxtHead3)
    TBRec.txtHead4 = QPTrim(fptxtHead4)
    TBRec.txtHead5 = QPTrim(fptxtHead5)
    If fpcboLogo.Text = "Yes" Then
      TBRec.dologo = 1
    Else
      TBRec.dologo = 0
    End If
  Put #TBHandle, 1, TBRec
  Close TBHandle
  
  Call Savemsg(900, "Tax bill information was saved successfully.")
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersLsrItemized", "SaveTxDef", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

  
End Sub
   
Private Sub LoadMe()
  Dim TBRec As TxBillLaserItemized
  Dim TBHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  On Error GoTo ERRORSTUFF
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)

  fpcboLogo.AddItem "No"
  fpcboLogo.AddItem "Yes"
  
  If Exist(LaserPersItemBill) Then
    OpenLaserPersItemized TBHandle
    Get #TBHandle, 1, TBRec
    Close TBHandle
    If TBRec.dologo = 1 Then
      fpcboLogo.Text = "Yes"
    Else
      fpcboLogo.Text = "No"
    End If
    fpcboLogo.ListIndex = TBRec.dologo
    fpTxtHead1 = QPTrim(TBRec.TxtHead1)
    fpTxtHead2 = QPTrim(TBRec.TxtHead2)
    fpTxtOpt1 = QPTrim(TBRec.txtOpt1)
    fpTxtOpt2 = QPTrim(TBRec.TxtOpt2)
    fpTxtOpt3 = QPTrim(TBRec.TxtOpt3)
    fptxtPgph(0) = QPTrim(TBRec.txtPgph0)
    fptxtPgph(1) = QPTrim(TBRec.txtPgph1)
    fptxtPgph(2) = QPTrim(TBRec.txtPgph2)
    fptxtPgph(3) = QPTrim(TBRec.txtPgph3)
    fptxtPgph(4) = QPTrim(TBRec.txtPgph4)
    fptxtHead3 = QPTrim(TBRec.txtHead3)
    fptxtHead4 = QPTrim(TBRec.txtHead4)
    fptxtHead5 = QPTrim(TBRec.txtHead5)
  Else
    fpcboLogo.Text = "No"
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersLtrItemized", "LoadMe", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub


