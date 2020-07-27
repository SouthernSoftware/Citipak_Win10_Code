VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPBillGLSetUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax GL-Interface Personal Account Setup"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPBillGLSetUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListYear 
      Height          =   3792
      Left            =   1200
      TabIndex        =   0
      Top             =   2616
      Width           =   972
      _Version        =   196608
      _ExtentX        =   1714
      _ExtentY        =   6689
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
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
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
      ColDesigner     =   "frmVATaxPBillGLSetUp.frx":08CA
   End
   Begin EditLib.fpText fptxtPersDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   1
      Top             =   2496
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtPersCredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   2
      Top             =   2496
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMTDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   3
      Top             =   2900
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMTCredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   4
      Top             =   2900
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMCDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   5
      Top             =   3316
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMCCredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   6
      Top             =   3316
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtFEDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   7
      Top             =   3728
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtFECredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   8
      Top             =   3728
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOR1Debit 
      Height          =   372
      Left            =   6000
      TabIndex        =   9
      Top             =   5952
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOR1Credit 
      Height          =   372
      Left            =   8160
      TabIndex        =   10
      Top             =   5952
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOR2Debit 
      Height          =   372
      Left            =   6000
      TabIndex        =   11
      Top             =   6384
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOR2Credit 
      Height          =   372
      Left            =   8160
      TabIndex        =   12
      Top             =   6384
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOR3Debit 
      Height          =   372
      Left            =   6000
      TabIndex        =   13
      Top             =   6816
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOR3Credit 
      Height          =   372
      Left            =   8160
      TabIndex        =   14
      Top             =   6816
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   624
      Left            =   1579
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7791
      Width           =   2388
      _Version        =   131072
      _ExtentX        =   4212
      _ExtentY        =   1101
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmVATaxPBillGLSetUp.frx":0B56
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   630
      Left            =   7695
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7785
      Width           =   2385
      _Version        =   131072
      _ExtentX        =   4207
      _ExtentY        =   1111
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmVATaxPBillGLSetUp.frx":0D35
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   492
      Left            =   1032
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1428
      _Version        =   131072
      _ExtentX        =   2519
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmVATaxPBillGLSetUp.frx":0F12
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLTest 
      Height          =   624
      Left            =   4644
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7791
      Width           =   2376
      _Version        =   131072
      _ExtentX        =   4191
      _ExtentY        =   1101
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmVATaxPBillGLSetUp.frx":10F1
   End
   Begin EditLib.fpText fptxtMHDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   19
      Top             =   4156
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMHCredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   20
      Top             =   4156
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtIntDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   21
      Top             =   4570
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtIntCredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   22
      Top             =   4560
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtPenDebit 
      Height          =   372
      Left            =   6000
      TabIndex        =   23
      Top             =   4990
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtPenCredit 
      Height          =   372
      Left            =   8160
      TabIndex        =   24
      Top             =   4990
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3408
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
      AlignTextH      =   1
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Opt Rev 3:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2940
      TabIndex        =   41
      Top             =   6936
      Width           =   2892
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Opt Rev 2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2940
      TabIndex        =   40
      Top             =   6504
      Width           =   2892
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   10785
      X2              =   2640
      Y1              =   5592
      Y2              =   5592
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Optional Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   324
      Left            =   2640
      TabIndex        =   39
      Top             =   5592
      Width           =   2292
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Opt Rev 1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2940
      TabIndex        =   38
      Top             =   6072
      Width           =   2892
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Equipment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   37
      Top             =   3812
      Width           =   2412
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1259
      Top             =   450
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Tax GL Interface Account Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2760
      TabIndex        =   36
      Top             =   612
      Width           =   5892
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1320
      TabIndex        =   35
      Top             =   2256
      Width           =   732
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type: BILLING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3539
      TabIndex        =   34
      Top             =   1290
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   33
      Top             =   2616
      Width           =   1932
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Machine Tools:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   32
      Top             =   3020
      Width           =   2412
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Merchant Capital:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   31
      Top             =   3416
      Width           =   2652
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6240
      TabIndex        =   30
      Top             =   2136
      Width           =   1572
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   8400
      TabIndex        =   29
      Top             =   2136
      Width           =   1572
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5400
      Left            =   840
      Top             =   2040
      Width           =   9972
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   2040
      Y2              =   7420
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   3720
      TabIndex        =   28
      Top             =   1680
      Width           =   3372
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Homes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   27
      Top             =   4240
      Width           =   2412
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   26
      Top             =   4644
      Width           =   2412
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3000
      TabIndex        =   25
      Top             =   5074
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1259
      Top             =   345
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxPBillGLSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Public GThisYear As String
  Dim TempPersDBAcct As String
  Dim TempPersCRAcct As String
  Dim TempMTDBAcct As String
  Dim TempMTCRAcct As String
  Dim TempMCDBAcct As String
  Dim TempMCCRAcct As String
  Dim TempFEDBAcct As String
  Dim TempFECRAcct As String
  Dim TempMHDBAcct As String
  Dim TempMHCRAcct As String
  Dim TempIntDBAcct As String
  Dim TempIntCRAcct As String
  Dim TempPenDBAcct As String
  Dim TempPenCRAcct As String
  Dim TempOpt1DBAcct As String
  Dim TempOpt1CRAcct As String
  Dim TempOpt2DBAcct As String
  Dim TempOpt2CRAcct As String
  Dim TempOpt3DBAcct As String
  Dim TempOpt3CRAcct As String
  Dim TempYear As Integer
  Dim Exit2Bill As Boolean
  Dim Exit2Int As Boolean
  Dim Exit2Man As Boolean
  Dim Fund As Integer
  Dim Dept As Integer
  Dim Detail As Integer
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$

Private Sub cmdExit_Click()
  If Check4Changes = True Then
    Exit Sub
  End If
  Call LogSaves
  KillFile "C:\CPWork\taxpayGL.dat"
  If Exit2Bill = True Then
    If Exist("C:\CPWork\revrglbill.dat") Then KillFile "C:\CPWork\revrglbill.dat"
    frmVATaxPrebilling.Show
    DoEvents
    Unload Me
    Exit Sub
  ElseIf Exit2Int = True Then
    If Exist("C:\CPWork\revglint.dat") Then KillFile "C:\CPWork\revglint.dat"
    frmVATaxCalcInterest.Show
    DoEvents
    Unload Me
    Exit Sub
  ElseIf Exit2Man = True Then
    If Exist("C:\CPWork\revglman.dat") Then KillFile "C:\CPWork\revglman.dat"
    frmVATaxManualBillEntry.Show
    DoEvents
    Unload Me
    Exit Sub
  End If
  KillFile "C:\CPWork\taxpbillGL.dat"
  Unload frmVATaxGLList
  frmVATaxBillSetUpMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdGLList_Click()
  frmVATaxGLList.Show ' vbModal
End Sub

Private Sub cmdGLTest_Click()
   Dim IdxRec As JGLAcctIdxType
   Dim GLIdxNum$
   Dim IdxHandle As Integer
   Dim IdxCnt As Integer
   Dim x As Integer, y As Integer
   Dim GLRec As GLAcctRecType
   Dim GLHandle As Integer
   Dim GLCnt As Integer
   Dim ListCnt As Integer
   Dim EmptyCnt As Integer
   On Error GoTo ERRORSTUFF
   
   GLCnt = 0
   OpenGLIdxFile IdxHandle, IdxCnt
   
   If IdxCnt = 0 Then
     MsgBox "ERROR: No General Ledger index file could be found. General Ledger list cannot be displayed."
     Close IdxHandle
     Exit Sub
   End If
   ReDim IdxRecs(1 To IdxCnt) As Integer
   For x = 1 To IdxCnt
     Get IdxHandle, x, IdxRec
     IdxRecs(x) = IdxRec.RecNo
   Next x
   Close IdxHandle
   
   OpenGLAcctFile GLHandle, GLCnt
   If GLCnt = 0 Then
     frmVATaxMsg.Label1.Caption = "ERROR: No General Ledger file could be found. The General Ledger list cannot be loaded."
     frmVATaxMsg.Label1.Top = 900
     frmVATaxMsg.Show vbModal
     Close GLHandle
     Exit Sub
   End If
   
   If GLCnt < IdxCnt Then
     frmVATaxMsg.Label1.Caption = "ERROR: The GL index count is greater than the GL file count."
     frmVATaxMsg.Label1.Top = 900
     frmVATaxMsg.Show vbModal
   End If
   
   ListCnt = 20
   If Opt1Desc <> "" Then ListCnt = ListCnt + 2
   If Opt2Desc <> "" Then ListCnt = ListCnt + 2
   If Opt3Desc <> "" Then ListCnt = ListCnt + 2
   ReDim GTestOK(1 To ListCnt) As Boolean
   ReDim GTestNums(1 To ListCnt) As String
   ReDim GTestDbCrt(1 To ListCnt) As String
   ReDim GTestDesc(1 To ListCnt) As String
   For x = 1 To ListCnt
     GTestNums(x) = ""
     GTestDbCrt(x) = ""
     GTestDesc(x) = ""
   Next x
   
   For x = 1 To ListCnt
     GTestOK(x) = False
     Select Case x
       Case 1
         GTestNums(x) = QPTrim$(fptxtPersDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Personal"
       Case 2
         GTestNums(x) = QPTrim$(fptxtPersCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Personal"
       Case 3
         GTestNums(x) = QPTrim$(fptxtMTDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Machine Tools"
       Case 4
         GTestNums(x) = QPTrim$(fptxtMTCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Machine Tools"
       Case 5
         GTestNums(x) = QPTrim$(fptxtMCDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Merchant Capital"
       Case 6
         GTestNums(x) = QPTrim$(fptxtMCCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Merchant Capital"
       Case 7
         GTestNums(x) = QPTrim$(fptxtFEDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Farm Equipment"
       Case 8
         GTestNums(x) = QPTrim$(fptxtFECredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Farm Equipment"
       Case 9
         GTestNums(x) = QPTrim$(fptxtMHDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Mobile Homes"
       Case 10
         GTestNums(x) = QPTrim$(fptxtMHCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Mobile Homes"
       Case 11
         GTestNums(x) = QPTrim$(fptxtIntDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Interest"
       Case 12
         GTestNums(x) = QPTrim$(fptxtIntCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Interest"
       Case 13
         GTestNums(x) = QPTrim$(fptxtPenDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Penalty"
       Case 14
         GTestNums(x) = QPTrim$(fptxtPenCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Penalty"
       Case 15
         GTestNums(x) = QPTrim$(fptxtOR1Debit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = Opt1Desc
       Case 16
         GTestNums(x) = QPTrim$(fptxtOR1Credit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = Opt1Desc
       Case 17
         GTestNums(x) = QPTrim$(fptxtOR2Debit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = Opt2Desc
       Case 18
         GTestNums(x) = QPTrim$(fptxtOR2Credit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = Opt2Desc
       Case 19
         GTestNums(x) = QPTrim$(fptxtOR3Debit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = Opt3Desc
       Case 20
         GTestNums(x) = QPTrim$(fptxtOR3Credit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = Opt3Desc
     End Select
   Next x
   
   If EmptyCnt = 0 Then
     Call TaxMsg(900, "There are no GL numbers entered. Verification not necessary.")
     Exit Sub
   End If
   
   For x = 1 To IdxCnt
     If IdxRecs(x) <> 0 Then
       Get GLHandle, IdxRecs(x), GLRec
       If GLRec.Deleted Then GoTo SkipIt
       For y = 1 To ListCnt
         If GTestOK(y) = False Then
           If GTestNums(y) = QPTrim$(GLRec.Num) Then
             GTestOK(y) = True
           End If
         End If
       Next y
    End If
SkipIt:
   Next x
   Close GLHandle
   
   For x = 1 To ListCnt
     If GTestOK(x) = False And GTestNums(x) <> "" Then
       frmVATaxBadGLList.Show vbModal
       Exit For
     End If
   Next x
     
   If x > ListCnt Then
     frmVATaxMsg.Label1.Caption = "All G/L numbers entries have been verified."
     frmVATaxMsg.Label1.Top = 900
     frmVATaxMsg.Show vbModal
   End If
   
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "cmdGLTest_Click", Erl)
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
    frmVATaxBillSetUpMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim RevRec As TaxPAcctsType
  Dim PPHandle As Integer
  Dim x As Integer
  Dim StartYear As Integer
  
  On Error GoTo ERRORSTUFF
  
  If VerifyGLNum(QPTrim$(fptxtPersDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Personal Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax personal debit number " + QPTrim$(fptxtPersDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtPersDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtPersCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Personal Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax personal credit number " + QPTrim$(fptxtPersCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtPersCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtMTDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Machine Tools Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax machine tools debit number " + QPTrim$(fptxtIntCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtMTDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtMTCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Machine Tools Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax machine tools credit number " + QPTrim$(fptxtMTCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtMTCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtMCDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Merchant Capital Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax merchant capital debit number " + QPTrim$(fptxtMCDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtMCDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtMCCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Merchant Capital Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax merchant capital credit number " + QPTrim$(fptxtMCCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtMCCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtFEDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Farm Equipment Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax farm equipment debit number " + QPTrim$(fptxtFEDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtFEDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtFECredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Farm Equipment Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax farm equipment credit number " + QPTrim$(fptxtFECredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtFECredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtMHDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Mobile Homes Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax mobile homes debit number " + QPTrim$(fptxtMHDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtMHDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtFECredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Mobile Homes Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax mobile homes credit number " + QPTrim$(fptxtMHCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtMHCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtIntDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Interest Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax interest debit number " + QPTrim$(fptxtIntDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtIntDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtIntCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Interest Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax interest credit number " + QPTrim$(fptxtIntCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtIntCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtPenDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Penalty Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax penalty debit number " + QPTrim$(fptxtPenDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtPenDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtPenCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Penalty Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax penalty credit number " + QPTrim$(fptxtPenCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtPenCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If fptxtOR1Debit.Enabled = False Then GoTo MoveTo2
  If VerifyGLNum(QPTrim$(fptxtOR1Debit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Optional Revenue #1 Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax optional revenue #1 debit number " + QPTrim$(fptxtOR1Debit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtOR1Debit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtOR1Credit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Optional Revenue #1 Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax optional revenue #1 credit number " + QPTrim$(fptxtOR1Credit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtOR1Credit.SetFocus
      Exit Sub
    End If
  End If
  
MoveTo2:
  If fptxtOR2Debit.Enabled = False Then GoTo MoveTo3
  If VerifyGLNum(QPTrim$(fptxtOR2Debit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Optional Revenue #2 Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax optional revenue #2 debit number " + QPTrim$(fptxtOR2Debit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtOR2Debit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtOR2Credit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Optional Revenue #2 Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax optional revenue #2 credit number " + QPTrim$(fptxtOR2Credit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtOR2Credit.SetFocus
      Exit Sub
    End If
  End If
MoveTo3:
  If fptxtOR3Debit.Enabled = False Then GoTo MoveOn
  If VerifyGLNum(QPTrim$(fptxtOR3Debit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Optional Revenue #3 Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax optional revenue #3 debit number " + QPTrim$(fptxtOR3Debit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtOR3Debit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtOR3Credit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Optional Revenue #3 Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax optional revenue #3 credit number " + QPTrim$(fptxtOR3Credit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtOR3Credit.SetFocus
      Exit Sub
    End If
  End If
  
MoveOn:
  ReDim GLYears(1 To 51) As Integer
  StartYear = 1979
  For x = 1 To 51
    StartYear = StartYear + 1
    GLYears(x) = StartYear
  Next x
  OpenPTaxGLInterBill PPHandle
  Get PPHandle, 1, RevRec
  For x = 1 To 51
    If GLYears(x) = CInt(GThisYear) Then
      RevRec.TaxAcct(x).TaxYear = CInt(GThisYear)
      RevRec.TaxAcct(x).PersDBAcct = QPTrim$(fptxtPersDebit.Text)
      RevRec.TaxAcct(x).PersCRAcct = QPTrim$(fptxtPersCredit.Text)
      RevRec.TaxAcct(x).MTDBAcct = QPTrim$(fptxtMTDebit.Text)
      RevRec.TaxAcct(x).MTCRAcct = QPTrim$(fptxtMTCredit.Text)
      RevRec.TaxAcct(x).MCCRAcct = QPTrim$(fptxtMCCredit.Text)
      RevRec.TaxAcct(x).MCDBAcct = QPTrim$(fptxtMCDebit.Text)
      RevRec.TaxAcct(x).FECRAcct = QPTrim$(fptxtFECredit.Text)
      RevRec.TaxAcct(x).FEDBAcct = QPTrim$(fptxtFEDebit.Text)
      RevRec.TaxAcct(x).MHCRAcct = QPTrim$(fptxtMHCredit.Text)
      RevRec.TaxAcct(x).MHDBAcct = QPTrim$(fptxtMHDebit.Text)
      RevRec.TaxAcct(x).IntDBAcct = QPTrim$(fptxtIntDebit.Text)
      RevRec.TaxAcct(x).IntCRAcct = QPTrim$(fptxtIntCredit.Text)
      RevRec.TaxAcct(x).PenDBAcct = QPTrim$(fptxtPenDebit.Text)
      RevRec.TaxAcct(x).PenCRAcct = QPTrim$(fptxtPenCredit.Text)
      RevRec.TaxAcct(x).Opt1CRAcct = QPTrim$(fptxtOR1Credit.Text)
      RevRec.TaxAcct(x).Opt1DBAcct = QPTrim$(fptxtOR1Debit.Text)
      RevRec.TaxAcct(x).Opt2CRAcct = QPTrim$(fptxtOR2Credit.Text)
      RevRec.TaxAcct(x).Opt2DBAcct = QPTrim$(fptxtOR2Debit.Text)
      RevRec.TaxAcct(x).Opt3CRAcct = QPTrim$(fptxtOR3Credit.Text)
      RevRec.TaxAcct(x).Opt3DBAcct = QPTrim$(fptxtOR3Debit.Text)
      Put PPHandle, 1, RevRec
      Exit For
    End If
  Next x
  Close PPHandle
  If x < 52 Then
    TempPersDBAcct = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
    TempPersCRAcct = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
    TempMTDBAcct = QPTrim$(RevRec.TaxAcct(x).MTDBAcct)
    TempMTCRAcct = QPTrim$(RevRec.TaxAcct(x).MTCRAcct)
    TempMCDBAcct = QPTrim$(RevRec.TaxAcct(x).MCDBAcct)
    TempMCCRAcct = QPTrim$(RevRec.TaxAcct(x).MCCRAcct)
    TempFEDBAcct = QPTrim$(RevRec.TaxAcct(x).FEDBAcct)
    TempFECRAcct = QPTrim$(RevRec.TaxAcct(x).FECRAcct)
    TempMHDBAcct = QPTrim$(RevRec.TaxAcct(x).MHDBAcct)
    TempMHCRAcct = QPTrim$(RevRec.TaxAcct(x).MHCRAcct)
    TempIntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
    TempIntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
    TempPenDBAcct = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
    TempPenCRAcct = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
    TempOpt1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
    TempOpt1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
    TempOpt2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
    TempOpt2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
    TempOpt3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
    TempOpt3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
    TempYear = RevRec.TaxAcct(x).TaxYear
  End If
  Unload frmVATaxGLList
  
  Call Savemsg(900, "Your Personal Pay Setup Data has been saved successfully.")
  
  If Exist("C:\CPWork\revpglbill.dat") Then KillFile "C:\CPWork\revpglbill.dat"
  
  If Exit2Bill = True Then
    DoEvents
    Unload Me
  ElseIf Exit2Int = True Then
    DoEvents
    Unload Me
  ElseIf Exit2Man = True Then
    DoEvents
    Unload Me
  End If
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "cmdSave_Click", Erl)
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Unload frmVATaxGLList
      KillFile "C:\CPWork\taxpayGL.dat"
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPBillGLSetup.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxPBillGLSetup.")
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim RevRec As TaxPAcctsType
  Dim PPHandle As Integer
  Dim x As Integer
  Dim One As Integer
  Dim AHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Not Exist("GLACCT.IDX") Or Not Exist("GLACCT.DAT") Then
    Fund = 0
    Dept = 0
    Detail = 0
  Else
    Call GetAcctStruct(CurrCitiPath, Fund, Dept, Detail)
  End If
  
  Me.HelpContextID = hlpTaxGLAccountsSetupP
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Exit2Bill = False
  Exit2Int = False
  Exit2Man = False
  
  If Exist("C:\CPWork\revpglbill.dat") Then
    Exit2Bill = True
  ElseIf Exist("C:\CPWork\revglint.dat") Then
    Exit2Int = True
  ElseIf Exist("C:\CPWork\revglman.dat") Then
    Exit2Man = True
  End If
  Opt1Desc = ""
  Opt2Desc = ""
  Opt3Desc = ""
  If QPTrim$(TaxMasterRec.POptRev1) <> "" Then
    Label10.Caption = QPTrim$(TaxMasterRec.POptRev1)
    Opt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  Else
    Label10.Caption = "NO OPTION 1 SAVED"
    fptxtOR1Debit.Enabled = False
    fptxtOR1Credit.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.POptRev2) <> "" Then
    Label12.Caption = QPTrim$(TaxMasterRec.POptRev2)
    Opt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  Else
    Label12.Caption = "NO OPTION 2 SAVED"
    fptxtOR2Debit.Enabled = False
    fptxtOR2Credit.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.POptRev3) <> "" Then
    Label13.Caption = QPTrim$(TaxMasterRec.POptRev3)
    Opt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  Else
    Label13.Caption = "NO OPTION 3 SAVED"
    fptxtOR3Debit.Enabled = False
    fptxtOR3Credit.Enabled = False
  End If
  
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\taxpbillGL.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  OpenPTaxGLInterBill PPHandle
  If Exist(TxPGLInterBill) Then
    Get PPHandle, 1, RevRec
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = 0 Then
        fpListYear.AddItem 1979 + x
      Else
        fpListYear.AddItem RevRec.TaxAcct(x).TaxYear
      End If
      If x = 1 Then
        lblYear.Caption = "For Year " + CStr(RevRec.TaxAcct(x).TaxYear)
        TempPersDBAcct = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
        TempPersCRAcct = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
        TempMTDBAcct = QPTrim$(RevRec.TaxAcct(x).MTDBAcct)
        TempMTCRAcct = QPTrim$(RevRec.TaxAcct(x).MTCRAcct)
        TempMCDBAcct = QPTrim$(RevRec.TaxAcct(x).MCDBAcct)
        TempMCCRAcct = QPTrim$(RevRec.TaxAcct(x).MCCRAcct)
        TempFEDBAcct = QPTrim$(RevRec.TaxAcct(x).FEDBAcct)
        TempFECRAcct = QPTrim$(RevRec.TaxAcct(x).FECRAcct)
        TempMHDBAcct = QPTrim$(RevRec.TaxAcct(x).MHDBAcct)
        TempMHCRAcct = QPTrim$(RevRec.TaxAcct(x).MHCRAcct)
        TempIntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
        TempIntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
        TempPenDBAcct = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
        TempPenCRAcct = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
        TempOpt1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
        TempOpt1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
        TempOpt2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
        TempOpt2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
        TempOpt3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
        TempOpt3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
        TempYear = RevRec.TaxAcct(x).TaxYear
        fptxtPersDebit.Text = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
        fptxtPersCredit.Text = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
        fptxtMTDebit.Text = QPTrim$(RevRec.TaxAcct(x).MTDBAcct)
        fptxtMTCredit.Text = QPTrim$(RevRec.TaxAcct(x).MTCRAcct)
        fptxtMCDebit.Text = QPTrim$(RevRec.TaxAcct(x).MCDBAcct)
        fptxtMCCredit.Text = QPTrim$(RevRec.TaxAcct(x).MCCRAcct)
        fptxtFEDebit.Text = QPTrim$(RevRec.TaxAcct(x).FEDBAcct)
        fptxtFECredit.Text = QPTrim$(RevRec.TaxAcct(x).FECRAcct)
        fptxtMHDebit.Text = QPTrim$(RevRec.TaxAcct(x).MHDBAcct)
        fptxtMHCredit.Text = QPTrim$(RevRec.TaxAcct(x).MHCRAcct)
        fptxtIntDebit.Text = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
        fptxtIntCredit.Text = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
        fptxtPersDebit.Text = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
        fptxtPersCredit.Text = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
        fptxtOR1Debit.Text = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
        fptxtOR1Credit.Text = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
        fptxtOR2Debit.Text = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
        fptxtOR2Credit.Text = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
        fptxtOR3Debit.Text = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
        fptxtOR3Credit.Text = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
      End If
    Next x
  Else
    For x = 1 To 51
      TempPersDBAcct = ""
      TempPersCRAcct = ""
      TempMTDBAcct = ""
      TempMTCRAcct = ""
      TempMCDBAcct = ""
      TempMCCRAcct = ""
      TempFEDBAcct = ""
      TempFECRAcct = ""
      TempMHDBAcct = ""
      TempMHCRAcct = ""
      TempIntDBAcct = ""
      TempIntCRAcct = ""
      TempPenDBAcct = ""
      TempPenCRAcct = ""
      TempOpt1DBAcct = ""
      TempOpt1CRAcct = ""
      TempOpt2DBAcct = ""
      TempOpt2CRAcct = ""
      TempOpt3DBAcct = ""
      TempOpt3CRAcct = ""
      TempYear = 0
      RevRec.TaxAcct(x).TaxYear = 1979 + x
      fpListYear.AddItem RevRec.TaxAcct(x).TaxYear
      RevRec.TaxAcct(x).PersDBAcct = ""
      RevRec.TaxAcct(x).PersCRAcct = ""
      RevRec.TaxAcct(x).MTDBAcct = ""
      RevRec.TaxAcct(x).MTCRAcct = ""
      RevRec.TaxAcct(x).MCCRAcct = ""
      RevRec.TaxAcct(x).MCDBAcct = ""
      RevRec.TaxAcct(x).FECRAcct = ""
      RevRec.TaxAcct(x).FEDBAcct = ""
      RevRec.TaxAcct(x).MHCRAcct = ""
      RevRec.TaxAcct(x).MHDBAcct = ""
      RevRec.TaxAcct(x).IntDBAcct = ""
      RevRec.TaxAcct(x).IntCRAcct = ""
      RevRec.TaxAcct(x).PersDBAcct = ""
      RevRec.TaxAcct(x).PersCRAcct = ""
      RevRec.TaxAcct(x).Opt1CRAcct = ""
      RevRec.TaxAcct(x).Opt1DBAcct = ""
      RevRec.TaxAcct(x).Opt2CRAcct = ""
      RevRec.TaxAcct(x).Opt2CRAcct = ""
      RevRec.TaxAcct(x).Opt3CRAcct = ""
      RevRec.TaxAcct(x).Opt3CRAcct = ""
    Next
    Put PPHandle, 1, RevRec
  End If
  Close PPHandle
  fpListYear.ListIndex = 0
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "LoadMe", Erl)
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

Private Sub fpListYear_Click()
  Dim RevRec As TaxPAcctsType
  Dim PPHandle As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Exist("C:\CPWork\revpglbill.dat") Then
    KillFile "C:\CPWork\revpglbill.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  ElseIf Exist("C:\CPWork\revglint.dat") Then
    KillFile "C:\CPWork\revglint.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  ElseIf Exist("C:\CPWork\revglman.dat") Then
    KillFile "C:\CPWork\revglman.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  Else
    If fpListYear.ListIndex = -1 Then fpListYear.ListIndex = 0
    fpListYear.Row = fpListYear.ListIndex
    GThisYear = fpListYear.Text
  End If
  If QPTrim$(GThisYear) = "" Then
    Close
    Exit Sub
  End If
  
  lblYear.Caption = "For Year " + GThisYear
  
  OpenPTaxGLInterBill PPHandle
  Get PPHandle, 1, RevRec
  Close PPHandle
  
  For x = 1 To 51
    If RevRec.TaxAcct(x).TaxYear = CInt(GThisYear) Then
      TempPersDBAcct = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
      TempPersCRAcct = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
      TempMTDBAcct = QPTrim$(RevRec.TaxAcct(x).MTDBAcct)
      TempMTCRAcct = QPTrim$(RevRec.TaxAcct(x).MTCRAcct)
      TempMCDBAcct = QPTrim$(RevRec.TaxAcct(x).MCDBAcct)
      TempMCCRAcct = QPTrim$(RevRec.TaxAcct(x).MCCRAcct)
      TempFEDBAcct = QPTrim$(RevRec.TaxAcct(x).FEDBAcct)
      TempFECRAcct = QPTrim$(RevRec.TaxAcct(x).FECRAcct)
      TempMHDBAcct = QPTrim$(RevRec.TaxAcct(x).MHDBAcct)
      TempMHCRAcct = QPTrim$(RevRec.TaxAcct(x).MHCRAcct)
      TempIntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
      TempIntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
      TempPenDBAcct = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
      TempPenCRAcct = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
      TempOpt1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
      TempOpt1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
      TempOpt2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
      TempOpt2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
      TempOpt3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
      TempOpt3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
      TempYear = RevRec.TaxAcct(x).TaxYear
      fptxtPersDebit.Text = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
      fptxtPersCredit.Text = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
      fptxtMTDebit.Text = QPTrim$(RevRec.TaxAcct(x).MTDBAcct)
      fptxtMTCredit.Text = QPTrim$(RevRec.TaxAcct(x).MTCRAcct)
      fptxtMCDebit.Text = QPTrim$(RevRec.TaxAcct(x).MCDBAcct)
      fptxtMCCredit.Text = QPTrim$(RevRec.TaxAcct(x).MCCRAcct)
      fptxtFEDebit.Text = QPTrim$(RevRec.TaxAcct(x).FEDBAcct)
      fptxtFECredit.Text = QPTrim$(RevRec.TaxAcct(x).FECRAcct)
      fptxtMHDebit.Text = QPTrim$(RevRec.TaxAcct(x).MHDBAcct)
      fptxtMHCredit.Text = QPTrim$(RevRec.TaxAcct(x).MHCRAcct)
      fptxtIntDebit.Text = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
      fptxtIntCredit.Text = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
      fptxtPenDebit.Text = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
      fptxtPenCredit.Text = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
      fptxtOR1Debit.Text = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
      fptxtOR1Credit.Text = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
      fptxtOR2Debit.Text = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
      fptxtOR2Credit.Text = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
      fptxtOR3Debit.Text = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
      fptxtOR3Credit.Text = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
      Exit For
    End If
  Next x
  
  If x > 51 Then
    TempPersDBAcct = ""
    TempPersCRAcct = ""
    TempMTDBAcct = ""
    TempMTCRAcct = ""
    TempMCDBAcct = ""
    TempMCCRAcct = ""
    TempFEDBAcct = ""
    TempFECRAcct = ""
    TempMHDBAcct = ""
    TempMHCRAcct = ""
    TempIntDBAcct = ""
    TempIntCRAcct = ""
    TempPenDBAcct = ""
    TempPenCRAcct = ""
    TempOpt1DBAcct = ""
    TempOpt1CRAcct = ""
    TempOpt2DBAcct = ""
    TempOpt2CRAcct = ""
    TempOpt3DBAcct = ""
    TempOpt3CRAcct = ""
    TempYear = 0
    fptxtPersDebit.Text = ""
    fptxtPersCredit.Text = ""
    fptxtMTDebit.Text = ""
    fptxtMTCredit.Text = ""
    fptxtMCDebit.Text = ""
    fptxtMCCredit.Text = ""
    fptxtFEDebit.Text = ""
    fptxtFECredit.Text = ""
    fptxtMHDebit.Text = ""
    fptxtMHCredit.Text = ""
    fptxtIntDebit.Text = ""
    fptxtIntCredit.Text = ""
    fptxtPenDebit.Text = ""
    fptxtPenCredit.Text = ""
    fptxtOR1Debit.Text = ""
    fptxtOR1Credit.Text = ""
    fptxtOR2Debit.Text = ""
    fptxtOR2Credit.Text = ""
    fptxtOR3Debit.Text = ""
    fptxtOR3Credit.Text = ""
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "fpListYear_Click", Erl)
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

Private Sub fptxtMTCredit_DblClick(Button As Integer)
  fptxtMTCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtMTCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtMTCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtMTCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtMTCredit.Text = ReplaceString(fptxtMTCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtMTCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtMTCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtMTCredit.Text = AddDashesToGLNumber(fptxtMTCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtMTDebit_DblClick(Button As Integer)
  fptxtMTDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtMTDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtMTDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtMTDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtMTDebit.Text = ReplaceString(fptxtMTDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtMTDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtMTDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtMTDebit.Text = AddDashesToGLNumber(fptxtMTDebit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtIntCredit_DblClick(Button As Integer)
  fptxtIntCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtIntCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtIntCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtIntCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtIntCredit.Text = ReplaceString(fptxtIntCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtIntCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtIntCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtIntCredit.Text = AddDashesToGLNumber(fptxtIntCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtIntDebit_DblClick(Button As Integer)
  fptxtIntDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtIntDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtIntDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtIntDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtIntDebit.Text = ReplaceString(fptxtIntDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtIntDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtIntDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtIntDebit.Text = AddDashesToGLNumber(fptxtIntDebit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtMCCredit_DblClick(Button As Integer)
  fptxtMCCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtMCCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtMCCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtMCCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtMCCredit.Text = ReplaceString(fptxtMCCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtMCCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtMCCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtMCCredit.Text = AddDashesToGLNumber(fptxtMCCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtMCDebit_DblClick(Button As Integer)
  fptxtMCDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtMCDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtMCDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtMCDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtMCDebit.Text = ReplaceString(fptxtMCDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtMCDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtMCDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtMCDebit.Text = AddDashesToGLNumber(fptxtMCDebit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtOR1Credit_DblClick(Button As Integer)
  fptxtOR1Credit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR1Credit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtOR1Credit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtOR1Credit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtOR1Credit.Text = ReplaceString(fptxtOR1Credit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtOR1Credit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOR1Credit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtOR1Credit.Text = AddDashesToGLNumber(fptxtOR1Credit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtOR1Debit_DblClick(Button As Integer)
  fptxtOR1Debit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR1Debit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtOR1Debit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtOR1Debit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtOR1Debit.Text = ReplaceString(fptxtOR1Debit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtOR1Debit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOR1Debit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtOR1Debit.Text = AddDashesToGLNumber(fptxtOR1Debit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtOR2Credit_DblClick(Button As Integer)
  fptxtOR2Credit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR2Credit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtOR2Credit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtOR2Credit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtOR2Credit.Text = ReplaceString(fptxtOR2Credit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtOR2Credit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOR2Credit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtOR2Credit.Text = AddDashesToGLNumber(fptxtOR2Credit.Text, Fund, Dept, Detail)
End Sub

Private Sub fptxtOR2Debit_DblClick(Button As Integer)
  fptxtOR2Debit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR2Debit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtOR2Debit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtOR2Debit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtOR2Debit.Text = ReplaceString(fptxtOR2Debit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtOR2Debit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOR2Debit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtOR2Debit.Text = AddDashesToGLNumber(fptxtOR2Debit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtOR3Credit_DblClick(Button As Integer)
  fptxtOR3Credit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR3Credit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtOR3Credit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtOR3Credit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtOR3Credit.Text = ReplaceString(fptxtOR3Credit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtOR3Credit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOR3Credit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtOR3Credit.Text = AddDashesToGLNumber(fptxtOR3Credit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtOR3Debit_DblClick(Button As Integer)
  fptxtOR3Debit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR3Debit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtOR3Debit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtOR3Debit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtOR3Debit.Text = ReplaceString(fptxtOR3Debit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtOR3Debit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOR3Debit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtOR3Debit.Text = AddDashesToGLNumber(fptxtOR3Debit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtPenCredit_DblClick(Button As Integer)
  fptxtPenCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtPenDebit_DblClick(Button As Integer)
  fptxtPenDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtPersCredit_DblClick(Button As Integer)
  fptxtPersCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtPersCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtPersCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtPersCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtPersCredit.Text = ReplaceString(fptxtPersCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtPersCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPersCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtPersCredit.Text = AddDashesToGLNumber(fptxtPersCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtPersDebit_DblClick(Button As Integer)
  fptxtPersDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtFECredit_DblClick(Button As Integer)
  fptxtFECredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtFECredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtFECredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtFECredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtFECredit.Text = ReplaceString(fptxtFECredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtFECredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtFECredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtFECredit.Text = AddDashesToGLNumber(fptxtFECredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtFEDebit_DblClick(Button As Integer)
  fptxtFEDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub
Private Sub fptxtFEDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtFEDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtFEDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtFEDebit.Text = ReplaceString(fptxtFEDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtFEDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtFEDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtFEDebit.Text = AddDashesToGLNumber(fptxtFEDebit.Text, Fund, Dept, Detail)

End Sub
Private Sub fptxtMHCredit_DblClick(Button As Integer)
  fptxtMHCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtMHCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtMHCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtMHCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtMHCredit.Text = ReplaceString(fptxtMHCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtMHCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtMHCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtMHCredit.Text = AddDashesToGLNumber(fptxtMHCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtMHDebit_DblClick(Button As Integer)
  fptxtMHDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtMHDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtMHDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtMHDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtMHDebit.Text = ReplaceString(fptxtMHDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtMHDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtMHDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtMHDebit.Text = AddDashesToGLNumber(fptxtMHDebit.Text, Fund, Dept, Detail)

End Sub

Private Function Check4Changes() As Boolean
  Dim RevRec As TaxPAcctsType
  Dim PPHandle As Integer
  Dim x As Integer
  Dim ThisControl As Control
  Dim PersDBAcct As String
  Dim PersCRAcct As String
  Dim MTDBAcct As String
  Dim MTCRAcct As String
  Dim MCDBAcct As String
  Dim MCCRAcct As String
  Dim FEDBAcct As String
  Dim FECRAcct As String
  Dim MHDBAcct As String
  Dim MHCRAcct As String
  Dim IntDBAcct As String
  Dim IntCRAcct As String
  Dim PenDBAcct As String
  Dim PenCRAcct As String
  Dim OR1DBAcct As String
  Dim OR1CRAcct As String
  Dim OR2DBAcct As String
  Dim OR2CRAcct As String
  Dim OR3DBAcct As String
  Dim OR3CRAcct As String
  Dim ThisStr As String
  Dim Thisx As Integer
  Dim choice As String
  Dim NewDesc As String
  
  On Error GoTo ERRORSTUFF
  
  Check4Changes = False
  If Exist(TxPGLInterBill) Then
    OpenPTaxGLInterBill PPHandle
    Get PPHandle, 1, RevRec
  Else
    frmVATaxMsgWOpts.Label1.Caption = "Are you sure you want to exit without saving?"
    frmVATaxMsgWOpts.Label1.Top = 900
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPersDebit.SetFocus
      Check4Changes = True
      Exit Function
    Else
      Unload frmVATaxMsgWOpts
    End If
  End If
  
  For x = 1 To 51
    If RevRec.TaxAcct(x).TaxYear = CInt(GThisYear) Then
      PersDBAcct = QPTrim$(RevRec.TaxAcct(x).PersDBAcct)
      PersCRAcct = QPTrim$(RevRec.TaxAcct(x).PersCRAcct)
      MTDBAcct = QPTrim$(RevRec.TaxAcct(x).MTDBAcct)
      MTCRAcct = QPTrim$(RevRec.TaxAcct(x).MTCRAcct)
      MCDBAcct = QPTrim$(RevRec.TaxAcct(x).MCDBAcct)
      MCCRAcct = QPTrim$(RevRec.TaxAcct(x).MCCRAcct)
      FEDBAcct = QPTrim$(RevRec.TaxAcct(x).FEDBAcct)
      FECRAcct = QPTrim$(RevRec.TaxAcct(x).FECRAcct)
      MHDBAcct = QPTrim$(RevRec.TaxAcct(x).MHDBAcct)
      MHCRAcct = QPTrim$(RevRec.TaxAcct(x).MHCRAcct)
      IntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
      IntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
      PenDBAcct = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
      PenCRAcct = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
      OR1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
      OR1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
      OR2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
      OR2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
      OR3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
      OR3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
      Thisx = x
      Exit For
    End If
  Next x
  
  Set ThisControl = fptxtPersDebit
  ThisStr = PersDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Personal Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).PersDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Tax Personal Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtPersCredit
  ThisStr = PersCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Tax Personal Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).PersCRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Tax Personal Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtMTDebit
  ThisStr = MTDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Machine Tools Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).MTDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Machine Tools Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtMTCredit
  ThisStr = MTCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Machine Tools Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).MTCRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Machine Tools Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtMCDebit
  ThisStr = MCDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Merchant Capital Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).MCDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Merchant Capital Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtMCCredit
  ThisStr = MCCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Merchant Capital Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).MCCRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Merchant Capital Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtFEDebit
  ThisStr = FEDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Farm Equipment Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).FEDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Farm Equipment Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtFECredit
  ThisStr = FECRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Farm Equipment Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).FECRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Farm Equipment Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtMHDebit
  ThisStr = MHDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Mobile Homes Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).MHDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Mobile Homes Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtMHCredit
  ThisStr = MHCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Mobile Homes Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).MHCRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Mobile Homes Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtIntDebit
  ThisStr = IntDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Interest Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).IntDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Interest Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtIntCredit
  ThisStr = IntCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Interest Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).IntCRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Interest Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtPenDebit
  ThisStr = PenDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalty Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).PenDBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Penalty Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtPenCredit
  ThisStr = PenCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalty Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).PenCRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "Penalty Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtOR1Debit
  ThisStr = OR1DBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The '" + Label10.Caption + "' Debit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).Opt1DBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label10.Caption + "' Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtOR1Credit
  ThisStr = OR1CRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The '" + Label10.Caption + "' Credit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).Opt1CRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label10.Caption + "' Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtOR2Debit
  ThisStr = OR2DBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The '" + Label12.Caption + "' Debit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).Opt2DBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label12.Caption + "' Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtOR2Credit
  ThisStr = OR2CRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The '" + Label12.Caption + "' Credit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).Opt2CRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label12.Caption + "' Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtOR3Debit
  ThisStr = OR3DBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The '" + Label13.Caption + "' Debit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).Opt3DBAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label13.Caption + "' Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtOR3Credit
  ThisStr = OR3CRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The '" + Label13.Caption + "' Credit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).Opt3CRAcct = QPTrim$(ThisControl.Text)
      Put PPHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label13.Caption + "' Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  Close PPHandle
  
  Exit Function

HandleChoice:
    Select Case choice
      Case "abandon"
        Close PPHandle
        frmVATaxBillSetUpMenu.Show
        DoEvents
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        ThisControl.SetFocus
        Close PPHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "Check4Changes", Erl)
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

End Function

Private Sub LogSaves()
  Dim ThisDesc$
  
  On Error GoTo ERRORSTUFF
  
  If InStr(lblYear.Caption, CStr(TempYear)) = 0 Then Exit Sub
  
  If QPTrim$(TempPersDBAcct) = "" Then TempPersDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtPersDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempPersDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Tax Personal Debit was changed from " + QPTrim$(TempPersDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempPersCRAcct) = "" Then TempPersCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtPersCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempPersCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Tax Personal Credit was changed from " + QPTrim$(TempPersCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempMTDBAcct) = "" Then TempMTDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtMTDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempMTDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Machine Tools Debit was changed from " + QPTrim$(TempMTDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempMTCRAcct) = "" Then TempMTCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtMTCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempMTCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Machine Tools Credit was changed from " + QPTrim$(TempMTCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempMCDBAcct) = "" Then TempMCDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtMCDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempMCDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Merchant Capital Debit was changed from " + QPTrim$(TempMCDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempMCCRAcct) = "" Then TempMCCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtMCCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempMCCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Merchant Capital Credit was changed from " + QPTrim$(TempMCCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempFEDBAcct) = "" Then TempFEDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtFEDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempFEDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Farm Equipment Debit was changed from " + QPTrim$(TempMCDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempFECRAcct) = "" Then TempFECRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtFECredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempFECRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Farm Equipment Credit was changed from " + QPTrim$(TempMCCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempMHDBAcct) = "" Then TempMHDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtMHDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempMHDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Mobile Homes Debit was changed from " + QPTrim$(TempMCDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempMHCRAcct) = "" Then TempFECRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtMHCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempMHCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Mobile Homes Credit was changed from " + QPTrim$(TempMCCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempIntDBAcct) = "" Then TempIntDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtIntDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempIntDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Interest Debit was changed from " + QPTrim$(TempIntDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempIntCRAcct) = "" Then TempIntCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtIntCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempIntCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Interest Credit was changed from " + QPTrim$(TempIntCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempPenDBAcct) = "" Then TempPenDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtPenDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempPenDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Penalty Debit was changed from " + QPTrim$(TempPenDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempPenCRAcct) = "" Then TempPenCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtPenCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempPenCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Penalty Credit was changed from " + QPTrim$(TempPenCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt1DBAcct) = "" Then TempOpt1DBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR1Debit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt1DBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label10.Caption + " Debit was changed from " + QPTrim$(TempOpt1DBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt1CRAcct) = "" Then TempOpt1CRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR1Credit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt1CRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label10.Caption + " Credit was changed from " + QPTrim$(TempOpt1CRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempOpt2DBAcct) = "" Then TempOpt2DBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR2Debit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt2DBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label12.Caption + " Debit was changed from " + QPTrim$(TempOpt2DBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt2CRAcct) = "" Then TempOpt2CRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR2Credit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt2CRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label12.Caption + " Credit was changed from " + QPTrim$(TempOpt2CRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempOpt3DBAcct) = "" Then TempOpt3DBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR3Debit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt3DBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label13.Caption + " Debit was changed from " + QPTrim$(TempOpt3DBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt3CRAcct) = "" Then TempOpt3CRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR3Credit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt3CRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label13.Caption + " Credit was changed from " + QPTrim$(TempOpt3CRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "LogSaves", Erl)
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

Private Sub fptxtPersDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  
  If QPTrim$(fptxtPersDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
  fptxtPersDebit.Text = ReplaceString(fptxtPersDebit.Text, "-", "")
  ThisLen = Len(QPTrim$(fptxtPersDebit.Text))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPersDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      Exit Sub
    End If
  End If
  fptxtPersDebit.Text = AddDashesToGLNumber(fptxtPersDebit.Text, Fund, Dept, Detail)

End Sub
Private Function VerifyGLNum(GLNum$) As Boolean
   Dim IdxRec As JGLAcctIdxType
   Dim GLIdxNum$
   Dim IdxHandle As Integer
   Dim IdxCnt As Integer
   Dim x As Integer, y As Integer
   Dim GLRec As GLAcctRecType
   Dim GLHandle As Integer
   Dim GLCnt As Integer
   Dim CheckThis$
   Dim EmptyCnt As Integer
   On Error GoTo ERRORSTUFF
   
   VerifyGLNum = True
   
   If Not Exist("GLACCT.IDX") Then
     Call TaxMsg(900, "Unable to locate 'GLACCT.IDX'. General Ledger numbers cannot be verified.")
     Exit Function
   End If
   
   OpenGLIdxFile IdxHandle, IdxCnt
   
   ReDim IdxRecs(1 To IdxCnt) As Integer
   If IdxCnt = 0 Then
     Close
     Exit Function
   End If
   
   For x = 1 To IdxCnt
     Get IdxHandle, x, IdxRec
     IdxRecs(x) = IdxRec.RecNo
   Next x
   Close IdxHandle
   
   If Not Exist("GLACCT.DAT") Then
     Call TaxMsg(900, "Unable to locate 'GLACCT.DAT'. General Ledger numbers cannot be verified.")
     Exit Function
   End If
   
   OpenGLAcctFile GLHandle, GLCnt
   If GLCnt = 0 Then
     Close GLHandle
     Exit Function
   End If
   
   CheckThis = QPTrim$(GLNum)
   For x = 1 To IdxCnt
     If IdxRecs(x) <> 0 Then
       Get GLHandle, IdxRecs(x), GLRec
       If GLRec.Deleted Then GoTo SkipIt
       If CheckThis = QPTrim$(GLRec.Num) Then
         Exit For
       End If
     End If
SkipIt:
   Next x
   Close GLHandle
   
   If x > IdxCnt Then
     VerifyGLNum = False
   End If
   
   Exit Function
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPBillGLSetup", "VerifyGLNum", Erl)
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
    frmVATaxBillSetUpMenu.Show
    DoEvents
    Unload Me

End Function




