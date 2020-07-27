VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPayGLSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax GL-Interface Real Account Setup"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxPayGLSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListYear 
      Height          =   3792
      Left            =   1620
      TabIndex        =   0
      Top             =   2772
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
      ColDesigner     =   "frmVATaxPayGLSetup.frx":08CA
   End
   Begin EditLib.fpText fptxtTPDebit 
      Height          =   372
      Left            =   6420
      TabIndex        =   1
      Top             =   2772
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtTPCredit 
      Height          =   372
      Left            =   8580
      TabIndex        =   2
      Top             =   2772
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtIDebit 
      Height          =   372
      Left            =   6420
      TabIndex        =   3
      Top             =   3252
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtICredit 
      Height          =   372
      Left            =   8580
      TabIndex        =   4
      Top             =   3252
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtACDebit 
      Height          =   372
      Left            =   6420
      TabIndex        =   5
      Top             =   3732
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtACCredit 
      Height          =   372
      Left            =   8580
      TabIndex        =   6
      Top             =   3732
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtLLDebit 
      Height          =   372
      Left            =   6420
      TabIndex        =   7
      Top             =   4236
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtLLCredit 
      Height          =   372
      Left            =   8580
      TabIndex        =   8
      Top             =   4236
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   6420
      TabIndex        =   11
      Top             =   5760
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   8580
      TabIndex        =   12
      Top             =   5760
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   6420
      TabIndex        =   13
      Top             =   6240
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   8580
      TabIndex        =   14
      Top             =   6240
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   6420
      TabIndex        =   15
      Top             =   6720
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   8580
      TabIndex        =   16
      Top             =   6720
      Width           =   1932
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Left            =   1680
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7716
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
      ButtonDesigner  =   "frmVATaxPayGLSetup.frx":0B56
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   624
      Left            =   7860
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7716
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
      ButtonDesigner  =   "frmVATaxPayGLSetup.frx":0D35
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   492
      Left            =   1300
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1548
      _Version        =   131072
      _ExtentX        =   2730
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
      ButtonDesigner  =   "frmVATaxPayGLSetup.frx":0F12
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLTest 
      Height          =   624
      Left            =   4800
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7716
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
      ButtonDesigner  =   "frmVATaxPayGLSetup.frx":10F1
   End
   Begin EditLib.fpText fptxtPenDebit 
      Height          =   372
      Left            =   6420
      TabIndex        =   9
      Top             =   4740
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
      Left            =   8580
      TabIndex        =   10
      Top             =   4740
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
   Begin VB.Label Label14 
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
      Left            =   3420
      TabIndex        =   35
      Top             =   4824
      Width           =   1932
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
      Left            =   3360
      TabIndex        =   30
      Top             =   6840
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
      Left            =   3360
      TabIndex        =   29
      Top             =   6360
      Width           =   2892
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11163
      X2              =   3048
      Y1              =   5400
      Y2              =   5400
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
      Left            =   3060
      TabIndex        =   28
      Top             =   5400
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
      Left            =   3360
      TabIndex        =   27
      Top             =   5880
      Width           =   2892
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Late Listing:"
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
      Left            =   3420
      TabIndex        =   26
      Top             =   4320
      Width           =   1932
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1680
      Top             =   375
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Tax GL Interface Account Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3330
      TabIndex        =   25
      Top             =   540
      Width           =   5295
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
      Left            =   1740
      TabIndex        =   24
      Top             =   2412
      Width           =   732
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type: PAYMENT"
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
      Left            =   3960
      TabIndex        =   23
      Top             =   1215
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Principle:"
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
      Left            =   3420
      TabIndex        =   22
      Top             =   2892
      Width           =   1932
   End
   Begin VB.Label Label5 
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
      Left            =   3420
      TabIndex        =   21
      Top             =   3372
      Width           =   1572
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Adv/Collect:"
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
      Left            =   3420
      TabIndex        =   20
      Top             =   3852
      Width           =   1692
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
      Left            =   6660
      TabIndex        =   19
      Top             =   2412
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
      Left            =   8820
      TabIndex        =   18
      Top             =   2412
      Width           =   1572
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5160
      Left            =   1080
      Top             =   2295
      Width           =   10095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2280
      Y2              =   7440
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
      Height          =   375
      Left            =   4140
      TabIndex        =   17
      Top             =   1710
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1680
      Top             =   270
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxPayGLSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Public GThisYear As String
  Dim TempTaxDBAcct As String
  Dim TempTaxCRAcct As String
  Dim TempIntDBAcct As String
  Dim TempIntCRAcct As String
  Dim TempAdvDBAcct As String
  Dim TempAdvCRAcct As String
  Dim TempLtLstDBAcct As String
  Dim TempLtLstCRAcct As String
  Dim TempPenDBAcct As String
  Dim TempPenCRAcct As String
  Dim TempOpt1DBAcct As String
  Dim TempOpt1CRAcct As String
  Dim TempOpt2DBAcct As String
  Dim TempOpt2CRAcct As String
  Dim TempOpt3DBAcct As String
  Dim TempOpt3CRAcct As String
  Dim TempYear As Integer
  Dim Exit2Pay As Boolean
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
  If Exit2Pay = True Then
    If Exist("C:\CPWork\revrglpay.dat") Then KillFile "C:\CPWork\revrglpay.dat"
    frmVATaxPayMenu.Show
    DoEvents
    Unload Me
    Exit Sub
  End If
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
   
   ListCnt = 16
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
         GTestNums(x) = QPTrim$(fptxtTPDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Tax Principle"
       Case 2
         GTestNums(x) = QPTrim$(fptxtTPCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Tax Principle"
       Case 3
         GTestNums(x) = QPTrim$(fptxtIDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Interest"
       Case 4
         GTestNums(x) = QPTrim$(fptxtICredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Interest"
       Case 5
         GTestNums(x) = QPTrim$(fptxtACDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Adv/Collect"
       Case 6
         GTestNums(x) = QPTrim$(fptxtACCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Adv/Collect"
       Case 7
         GTestNums(x) = QPTrim$(fptxtLLDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Penalty"
       Case 8
         GTestNums(x) = QPTrim$(fptxtLLCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Penalty"
       Case 9
         GTestNums(x) = QPTrim$(fptxtPenDebit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Penalty"
       Case 10
         GTestNums(x) = QPTrim$(fptxtPenCredit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Penalty"
       Case 11
         GTestNums(x) = QPTrim$(fptxtOR1Debit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = Opt1Desc
       Case 12
         GTestNums(x) = QPTrim$(fptxtOR1Credit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = Opt1Desc
       Case 13
         GTestNums(x) = QPTrim$(fptxtOR2Debit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = Opt2Desc
       Case 14
         GTestNums(x) = QPTrim$(fptxtOR2Credit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = Opt2Desc
       Case 15
         GTestNums(x) = QPTrim$(fptxtOR3Debit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = Opt3Desc
       Case 16
         GTestNums(x) = QPTrim$(fptxtOR3Credit.Text)
         If Len(GTestNums(x)) > 0 Then EmptyCnt = EmptyCnt + 1
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = Opt3Desc
     End Select
   Next x
   
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
   
   If EmptyCnt = 0 Then
     Call TaxMsg(900, "There are no GL numbers entered. Verification not necessary.")
     Exit Sub
   End If
   
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillGLSetUp", "cmdGLTest_Click", Erl)
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
  Dim RevRec As TaxRAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  Dim StartYear As Integer
  
  On Error GoTo ERRORSTUFF
  
  If VerifyGLNum(QPTrim$(fptxtTPDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Principle Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax principle debit number " + QPTrim$(fptxtTPDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtTPDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtTPCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Principle Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax principle credit number " + QPTrim$(fptxtTPCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtTPCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtIDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Interest Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax interest debit number " + QPTrim$(fptxtICredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtIDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtICredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Interest Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax interest credit number " + QPTrim$(fptxtICredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtICredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtACDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Adv/Collect Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax adv/collect debit number " + QPTrim$(fptxtACDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtACDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtACCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Adv/Collect Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax adv/collect credit number " + QPTrim$(fptxtACCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtACCredit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtLLDebit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Late Listing Debit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax late listing debit number " + QPTrim$(fptxtLLDebit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtLLDebit.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtLLCredit.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Tax Late Listing Credit number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the tax late listing credit number " + QPTrim$(fptxtLLCredit.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      fptxtLLCredit.SetFocus
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
  OpenRTaxGLInterPay RRHandle
  Get RRHandle, 1, RevRec
  For x = 1 To 51
    If RevRec.TaxAcct(x).TaxYear = 0 Then 'added 2/17/2011 to accommodate years past 2010
      RevRec.TaxAcct(x).TaxYear = CInt(GThisYear)
    End If
    If GLYears(x) = CInt(GThisYear) Then
      RevRec.TaxAcct(x).TaxYear = CInt(GThisYear)
      RevRec.TaxAcct(x).TaxDBAcct = QPTrim$(fptxtTPDebit.Text)
      RevRec.TaxAcct(x).TaxCRAcct = QPTrim$(fptxtTPCredit.Text)
      RevRec.TaxAcct(x).IntDBAcct = QPTrim$(fptxtIDebit.Text)
      RevRec.TaxAcct(x).IntCRAcct = QPTrim$(fptxtICredit.Text)
      RevRec.TaxAcct(x).AdvDBAcct = QPTrim$(fptxtACDebit.Text)
      RevRec.TaxAcct(x).AdvCRAcct = QPTrim$(fptxtACCredit.Text)
      RevRec.TaxAcct(x).LtLstCRAcct = QPTrim$(fptxtLLCredit.Text)
      RevRec.TaxAcct(x).LtLstDBAcct = QPTrim$(fptxtLLDebit.Text)
      RevRec.TaxAcct(x).PenDBAcct = QPTrim$(fptxtPenDebit.Text)
      RevRec.TaxAcct(x).PenCRAcct = QPTrim$(fptxtPenCredit.Text)
      RevRec.TaxAcct(x).Opt1CRAcct = QPTrim$(fptxtOR1Credit.Text)
      RevRec.TaxAcct(x).Opt1DBAcct = QPTrim$(fptxtOR1Debit.Text)
      RevRec.TaxAcct(x).Opt2CRAcct = QPTrim$(fptxtOR2Credit.Text)
      RevRec.TaxAcct(x).Opt2DBAcct = QPTrim$(fptxtOR2Debit.Text)
      RevRec.TaxAcct(x).Opt3CRAcct = QPTrim$(fptxtOR3Credit.Text)
      RevRec.TaxAcct(x).Opt3DBAcct = QPTrim$(fptxtOR3Debit.Text)
      Put RRHandle, 1, RevRec
      Exit For
    End If
  Next x
  Close RRHandle
  If x < 52 Then
    TempTaxDBAcct = QPTrim$(RevRec.TaxAcct(x).TaxDBAcct)
    TempTaxCRAcct = QPTrim$(RevRec.TaxAcct(x).TaxCRAcct)
    TempIntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
    TempIntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
    TempAdvDBAcct = QPTrim$(RevRec.TaxAcct(x).AdvDBAcct)
    TempAdvCRAcct = QPTrim$(RevRec.TaxAcct(x).AdvCRAcct)
    TempLtLstDBAcct = QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct)
    TempLtLstCRAcct = QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct)
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
  
  Call Savemsg(900, "Your Pay Setup Data has been saved successfully.")
  
  If Exist("C:\CPWork\revrglpay.dat") Then KillFile "C:\CPWork\revrglpay.dat"
  If Exit2Pay = True Then
    Unload frmVATaxPayMenu
    frmVATaxPaymentEntry.Show
    DoEvents
    Unload Me
  End If
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayGLSetup", "cmdSave_Click", Erl)
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPayGLSetUp.")
      Call Terminate
      End
    End If
  End If
End Sub

'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxPayGLSetUp.")
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim RevRec As TaxRAcctsType
  Dim RRHandle As Integer
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
  
  Exit2Pay = False
  If Exist("C:\CPWork\revrglpay.dat") Then
    Exit2Pay = True
  End If

  If QPTrim$(TaxMasterRec.OptRev1) <> "" Then
    Label10.Caption = QPTrim$(TaxMasterRec.OptRev1)
  Else
    Label10.Caption = "NO OPTION 1 SAVED"
    fptxtOR1Debit.Enabled = False
    fptxtOR1Credit.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.OptRev2) <> "" Then
    Label12.Caption = QPTrim$(TaxMasterRec.OptRev2)
  Else
    Label12.Caption = "NO OPTION 2 SAVED"
    fptxtOR2Debit.Enabled = False
    fptxtOR2Credit.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.OptRev3) <> "" Then
    Label13.Caption = QPTrim$(TaxMasterRec.OptRev3)
  Else
    Label13.Caption = "NO OPTION 3 SAVED"
    fptxtOR3Debit.Enabled = False
    fptxtOR3Credit.Enabled = False
  End If
  
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\taxpayGL.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  Opt1Desc = ""
  Opt2Desc = ""
  Opt3Desc = ""
  If QPTrim$(TaxMasterRec.OptRev1) <> "" Then
    Label10.Caption = QPTrim$(TaxMasterRec.OptRev1)
    Opt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  Else
    Label10.Caption = "NO OPTION 1 SAVED"
    fptxtOR1Debit.Enabled = False
    fptxtOR1Credit.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.OptRev2) <> "" Then
    Label12.Caption = QPTrim$(TaxMasterRec.OptRev2)
    Opt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  Else
    Label12.Caption = "NO OPTION 2 SAVED"
    fptxtOR2Debit.Enabled = False
    fptxtOR2Credit.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.OptRev3) <> "" Then
    Label13.Caption = QPTrim$(TaxMasterRec.OptRev3)
    Opt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  Else
    Label13.Caption = "NO OPTION 3 SAVED"
    fptxtOR3Debit.Enabled = False
    fptxtOR3Credit.Enabled = False
  End If
  
  OpenRTaxGLInterPay RRHandle
  If Exist(TxRGLInterPay) Then
    Get RRHandle, 1, RevRec
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = 0 Then
        fpListYear.AddItem 1979 + x
      Else
        fpListYear.AddItem RevRec.TaxAcct(x).TaxYear
      End If
      If x = 1 Then
        lblYear.Caption = "For Year " + CStr(RevRec.TaxAcct(x).TaxYear)
        TempTaxDBAcct = QPTrim$(RevRec.TaxAcct(x).TaxDBAcct)
        TempTaxCRAcct = QPTrim$(RevRec.TaxAcct(x).TaxCRAcct)
        TempIntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
        TempIntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
        TempAdvDBAcct = QPTrim$(RevRec.TaxAcct(x).AdvDBAcct)
        TempAdvCRAcct = QPTrim$(RevRec.TaxAcct(x).AdvCRAcct)
        TempLtLstDBAcct = QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct)
        TempLtLstCRAcct = QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct)
        TempPenDBAcct = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
        TempPenCRAcct = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
        TempOpt1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
        TempOpt1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
        TempOpt2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
        TempOpt2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
        TempOpt3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
        TempOpt3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
        TempYear = RevRec.TaxAcct(x).TaxYear
        fptxtTPDebit.Text = QPTrim$(RevRec.TaxAcct(x).TaxDBAcct)
        fptxtTPCredit.Text = QPTrim$(RevRec.TaxAcct(x).TaxCRAcct)
        fptxtIDebit.Text = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
        fptxtICredit.Text = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
        fptxtACDebit.Text = QPTrim$(RevRec.TaxAcct(x).AdvDBAcct)
        fptxtACCredit.Text = QPTrim$(RevRec.TaxAcct(x).AdvCRAcct)
        fptxtLLDebit.Text = QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct)
        fptxtLLCredit.Text = QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct)
        fptxtPenDebit.Text = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
        fptxtPenCredit.Text = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
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
      TempTaxDBAcct = ""
      TempTaxCRAcct = ""
      TempIntDBAcct = ""
      TempIntCRAcct = ""
      TempAdvDBAcct = ""
      TempAdvCRAcct = ""
      TempLtLstDBAcct = ""
      TempLtLstCRAcct = ""
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
      RevRec.TaxAcct(x).TaxDBAcct = ""
      RevRec.TaxAcct(x).TaxCRAcct = ""
      RevRec.TaxAcct(x).IntDBAcct = ""
      RevRec.TaxAcct(x).IntCRAcct = ""
      RevRec.TaxAcct(x).AdvDBAcct = ""
      RevRec.TaxAcct(x).AdvCRAcct = ""
      RevRec.TaxAcct(x).LtLstCRAcct = ""
      RevRec.TaxAcct(x).LtLstDBAcct = ""
      RevRec.TaxAcct(x).PenDBAcct = ""
      RevRec.TaxAcct(x).PenCRAcct = ""
      RevRec.TaxAcct(x).Opt1CRAcct = ""
      RevRec.TaxAcct(x).Opt1DBAcct = ""
      RevRec.TaxAcct(x).Opt2CRAcct = ""
      RevRec.TaxAcct(x).Opt2CRAcct = ""
      RevRec.TaxAcct(x).Opt3CRAcct = ""
      RevRec.TaxAcct(x).Opt3CRAcct = ""
    Next
    Put RRHandle, 1, RevRec
  End If
  Close RRHandle
  fpListYear.ListIndex = 0
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayGLSetup", "LoadMe", Erl)
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
  Dim RevRec As TaxRAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Exist("C:\CPWork\revrglpay.dat") Then
    KillFile "C:\CPWork\revrglpay.dat"
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
  
  OpenRTaxGLInterPay RRHandle
  Get RRHandle, 1, RevRec
  Close RRHandle
  
  For x = 1 To 51
    If RevRec.TaxAcct(x).TaxYear = CInt(GThisYear) Then
      TempTaxDBAcct = QPTrim$(RevRec.TaxAcct(x).TaxDBAcct)
      TempTaxCRAcct = QPTrim$(RevRec.TaxAcct(x).TaxCRAcct)
      TempIntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
      TempIntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
      TempAdvDBAcct = QPTrim$(RevRec.TaxAcct(x).AdvDBAcct)
      TempAdvCRAcct = QPTrim$(RevRec.TaxAcct(x).AdvCRAcct)
      TempLtLstDBAcct = QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct)
      TempLtLstCRAcct = QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct)
      TempPenDBAcct = QPTrim$(RevRec.TaxAcct(x).PenDBAcct)
      TempPenCRAcct = QPTrim$(RevRec.TaxAcct(x).PenCRAcct)
      TempOpt1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
      TempOpt1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
      TempOpt2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
      TempOpt2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
      TempOpt3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
      TempOpt3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
      TempYear = RevRec.TaxAcct(x).TaxYear
      fptxtTPDebit.Text = QPTrim$(RevRec.TaxAcct(x).TaxDBAcct)
      fptxtTPCredit.Text = QPTrim$(RevRec.TaxAcct(x).TaxCRAcct)
      fptxtIDebit.Text = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
      fptxtICredit.Text = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
      fptxtACDebit.Text = QPTrim$(RevRec.TaxAcct(x).AdvDBAcct)
      fptxtACCredit.Text = QPTrim$(RevRec.TaxAcct(x).AdvCRAcct)
      fptxtLLDebit.Text = QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct)
      fptxtLLCredit.Text = QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct)
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
    TempTaxDBAcct = ""
    TempTaxCRAcct = ""
    TempIntDBAcct = ""
    TempIntCRAcct = ""
    TempAdvDBAcct = ""
    TempAdvCRAcct = ""
    TempLtLstDBAcct = ""
    TempLtLstCRAcct = ""
    TempPenDBAcct = ""
    TempPenCRAcct = ""
    TempOpt1DBAcct = ""
    TempOpt1CRAcct = ""
    TempOpt2DBAcct = ""
    TempOpt2CRAcct = ""
    TempOpt3DBAcct = ""
    TempOpt3CRAcct = ""
    TempYear = 0
    fptxtTPDebit.Text = ""
    fptxtTPCredit.Text = ""
    fptxtIDebit.Text = ""
    fptxtICredit.Text = ""
    fptxtACDebit.Text = ""
    fptxtACCredit.Text = ""
    fptxtLLDebit.Text = ""
    fptxtLLCredit.Text = ""
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayGLSetup", "fpListYear_Click", Erl)
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

Private Sub fptxtACCredit_DblClick(Button As Integer)
  fptxtACCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtACCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtACCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtACCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtACCredit.Text = ReplaceString(fptxtACCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtACCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtACCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtACCredit.Text = AddDashesToGLNumber(fptxtACCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtACDebit_DblClick(Button As Integer)
  fptxtACDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtACDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtACDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtACDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtACDebit.Text = ReplaceString(fptxtACDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtACDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtACDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtACDebit.Text = AddDashesToGLNumber(fptxtACDebit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtICredit_DblClick(Button As Integer)
  fptxtICredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtICredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtICredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtICredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtICredit.Text = ReplaceString(fptxtICredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtICredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtICredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtICredit.Text = AddDashesToGLNumber(fptxtICredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtIDebit_DblClick(Button As Integer)
  fptxtIDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtIDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtIDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtIDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtIDebit.Text = ReplaceString(fptxtIDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtIDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtIDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtIDebit.Text = AddDashesToGLNumber(fptxtIDebit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtLLCredit_DblClick(Button As Integer)
  fptxtLLCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtLLCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtLLCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtLLCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtLLCredit.Text = ReplaceString(fptxtLLCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtLLCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtLLCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtLLCredit.Text = AddDashesToGLNumber(fptxtLLCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtLLDebit_DblClick(Button As Integer)
  fptxtLLDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtLLDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtLLDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtLLDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtLLDebit.Text = ReplaceString(fptxtLLDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtLLDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtLLDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtLLDebit.Text = AddDashesToGLNumber(fptxtLLDebit.Text, Fund, Dept, Detail)

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

Private Sub fptxtPenCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtPenCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPenCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  End If

End Sub

Private Sub fptxtPenDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtPenDebit.Text)
  If ChkGL$ = "" Then Exit Sub
  
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPenDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  End If

End Sub

Private Sub fptxtTPCredit_DblClick(Button As Integer)
  fptxtTPCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtTPCredit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtTPCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtTPCredit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtTPCredit.Text = ReplaceString(fptxtTPCredit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtTPCredit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtTPCredit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtTPCredit.Text = AddDashesToGLNumber(fptxtTPCredit.Text, Fund, Dept, Detail)

End Sub

Private Sub fptxtTPDebit_DblClick(Button As Integer)
  fptxtTPDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Function Check4Changes() As Boolean
  Dim RevRec As TaxRAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  Dim ThisControl As Control
  Dim TaxDBAcct As String
  Dim TaxCRAcct As String
  Dim IntDBAcct As String
  Dim IntCRAcct As String
  Dim AdvDBAcct As String
  Dim AdvCRAcct As String
  Dim LLDBAcct As String
  Dim LLCRAcct As String
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
  If Exist(TxRGLInterPay) Then
    OpenRTaxGLInterPay RRHandle
    Get RRHandle, 1, RevRec
  Else
    frmVATaxMsgWOpts.Label1.Caption = "Are you sure you want to exit without saving?"
    frmVATaxMsgWOpts.Label1.Top = 900
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtTPDebit.SetFocus
      Check4Changes = True
      Exit Function
    Else
      Unload frmVATaxMsgWOpts
    End If
  End If
  
  For x = 1 To 51
    If RevRec.TaxAcct(x).TaxYear = CInt(GThisYear) Then
      TaxDBAcct = QPTrim$(RevRec.TaxAcct(x).TaxDBAcct)
      TaxCRAcct = QPTrim$(RevRec.TaxAcct(x).TaxCRAcct)
      IntDBAcct = QPTrim$(RevRec.TaxAcct(x).IntDBAcct)
      IntCRAcct = QPTrim$(RevRec.TaxAcct(x).IntCRAcct)
      AdvDBAcct = QPTrim$(RevRec.TaxAcct(x).AdvDBAcct)
      AdvCRAcct = QPTrim$(RevRec.TaxAcct(x).AdvCRAcct)
      LLDBAcct = QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct)
      LLCRAcct = QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct)
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
  
  Set ThisControl = fptxtTPDebit
  ThisStr = TaxDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Tax Principle Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).TaxDBAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Tax Principle Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtTPCredit
  ThisStr = TaxCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Tax Principle Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).TaxCRAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Tax Principle Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtIDebit
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
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Interest Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtICredit
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
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Interest Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtACDebit
  ThisStr = AdvDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Adv/Collect Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).AdvDBAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Adv/Collect Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtACCredit
  ThisStr = AdvCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Adv/Collect Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).AdvCRAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Adv/Collect Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtLLDebit
  ThisStr = LLDBAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Late Listing Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).LtLstDBAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Late Listing Debit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtLLCredit
  ThisStr = LLCRAcct
  NewDesc = QPTrim$(ThisControl.Text)
  If NewDesc <> ThisStr Then
    If QPTrim$(ThisControl.Text) = "" Then NewDesc = "BLANK"
    If QPTrim$(ThisStr) = "" Then ThisStr = "BLANK"
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Late Listing Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).LtLstCRAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Late Listing Credit has been saved successfully.")
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
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
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "The '" + Label13.Caption + "' Credit has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  Close RRHandle
  
  Exit Function

HandleChoice:
    Select Case choice
      Case "abandon"
        Close RRHandle
        frmVATaxBillSetUpMenu.Show
        DoEvents
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        ThisControl.SetFocus
        Close RRHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayGLSetUp", "Check4Changes", Erl)
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
  
  If QPTrim$(TempTaxDBAcct) = "" Then TempTaxDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtTPDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempTaxDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Tax Principle Debit was changed from " + QPTrim$(TempTaxDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempTaxCRAcct) = "" Then TempTaxCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtTPCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempTaxCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Tax Principle Credit was changed from " + QPTrim$(TempTaxCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempIntDBAcct) = "" Then TempIntDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtIDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempIntDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Interest Debit was changed from " + QPTrim$(TempIntDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempIntCRAcct) = "" Then TempIntCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtICredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempIntCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Interest Credit was changed from " + QPTrim$(TempIntCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempAdvDBAcct) = "" Then TempAdvDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtACDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempAdvDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Adv/Collect Debit was changed from " + QPTrim$(TempAdvDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempAdvCRAcct) = "" Then TempAdvCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtACCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempAdvCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Adv/Collect Credit was changed from " + QPTrim$(TempAdvCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempLtLstDBAcct) = "" Then TempLtLstDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtLLDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempLtLstDBAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Late Listing Debit was changed from " + QPTrim$(TempLtLstDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempLtLstCRAcct) = "" Then TempLtLstCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtLLCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempLtLstCRAcct) <> ThisDesc Then
    MainLog ("frmVATaxBillSetUp: For Year " + CStr(TempYear) + ": Late Listing Credit was changed from " + QPTrim$(TempLtLstCRAcct) + " to " + ThisDesc + " and saved.")
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayGLSetup", "LogSaves", Erl)
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

Private Sub fptxtTPDebit_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ChkGL$
  
  ChkGL$ = QPTrim$(fptxtTPCredit.Text)
  If ChkGL$ = "" Then Exit Sub
  
'  If QPTrim$(fptxtTPDebit.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
'  fptxtTPDebit.Text = ReplaceString(fptxtTPDebit.Text, "-", "")
'  ThisLen = Len(QPTrim$(fptxtTPDebit.Text))
  ChkGL$ = ReplaceString(ChkGL, "-", "")
  ThisLen = Len(QPTrim$(ChkGL$))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtTPDebit.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
'      Exit Sub
    End If
  End If
'  fptxtTPDebit.Text = AddDashesToGLNumber(fptxtTPDebit.Text, Fund, Dept, Detail)

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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayGLSetup", "VerifyGLNum", Erl)
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

Private Sub fptxtPenCredit_DblClick(Button As Integer)
  fptxtPenCredit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtPenDebit_DblClick(Button As Integer)
  fptxtPenDebit.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

