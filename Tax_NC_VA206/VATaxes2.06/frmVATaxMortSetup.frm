VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxMortSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mortgage Codes"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxMortSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   825
      Left            =   3240
      TabIndex        =   24
      Top             =   1800
      Width           =   5055
      _Version        =   196608
      _ExtentX        =   8916
      _ExtentY        =   1455
      TextAlias       =   ""
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
      Columns         =   3
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
      ColDesigner     =   "frmVATaxMortSetup.frx":08CA
   End
   Begin EditLib.fpText fptxtMortCode 
      Height          =   372
      Left            =   4452
      TabIndex        =   0
      Top             =   3360
      Width           =   1560
      _Version        =   196608
      _ExtentX        =   2752
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   8
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
   Begin EditLib.fpText fptxtCmpnyName 
      Height          =   372
      Left            =   4446
      TabIndex        =   1
      Top             =   3840
      Width           =   4692
      _Version        =   196608
      _ExtentX        =   8276
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   32
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
   Begin EditLib.fpText fptxtAdd1 
      Height          =   372
      Left            =   4446
      TabIndex        =   2
      Top             =   4320
      Width           =   4692
      _Version        =   196608
      _ExtentX        =   8276
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   32
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
   Begin EditLib.fpText fptxtAdd2 
      Height          =   372
      Left            =   4446
      TabIndex        =   3
      Top             =   4800
      Width           =   4692
      _Version        =   196608
      _ExtentX        =   8276
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   32
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
   Begin EditLib.fpText fptxtAdd3 
      Height          =   372
      Left            =   4446
      TabIndex        =   4
      Top             =   5280
      Width           =   4692
      _Version        =   196608
      _ExtentX        =   8276
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   32
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
   Begin EditLib.fpText fptxtContact 
      Height          =   372
      Left            =   4446
      TabIndex        =   5
      Top             =   5760
      Width           =   4692
      _Version        =   196608
      _ExtentX        =   8276
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   32
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
   Begin EditLib.fpMask fptxtPhone 
      Height          =   372
      Left            =   4446
      TabIndex        =   6
      Top             =   6240
      Width           =   2076
      _Version        =   196608
      _ExtentX        =   3662
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483639
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "(###)-###-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
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
   Begin EditLib.fpText fptxtXFileName 
      Height          =   375
      Left            =   4446
      TabIndex        =   7
      Top             =   6720
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
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
      MaxLength       =   8
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
      Height          =   630
      Left            =   1485
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7680
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
      ButtonDesigner  =   "frmVATaxMortSetup.frx":0BDA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   636
      Left            =   7764
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7680
      Width           =   2388
      _Version        =   131072
      _ExtentX        =   4212
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxMortSetup.frx":0DB9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   636
      Left            =   4608
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2388
      _Version        =   131072
      _ExtentX        =   4212
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxMortSetup.frx":0F96
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT MODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   5040
      TabIndex        =   20
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "+ Year"
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
      Height          =   375
      Left            =   6606
      TabIndex        =   19
      Top             =   6810
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export File Name:"
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
      Height          =   375
      Left            =   2166
      TabIndex        =   18
      Top             =   6810
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4212
      Left            =   1974
      Top             =   3120
      Width           =   7692
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone #:"
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
      Left            =   2166
      TabIndex        =   17
      Top             =   6336
      Width           =   2052
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
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
      Left            =   2166
      TabIndex        =   16
      Top             =   5856
      Width           =   2052
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City, State, Zip:"
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
      Left            =   2166
      TabIndex        =   15
      Top             =   5376
      Width           =   2052
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line #2:"
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
      Left            =   2166
      TabIndex        =   14
      Top             =   4896
      Width           =   2052
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line #1:"
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
      Left            =   2166
      TabIndex        =   13
      Top             =   4416
      Width           =   2052
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
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
      Left            =   2406
      TabIndex        =   12
      Top             =   3912
      Width           =   1812
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mortgage Code:"
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
      Left            =   2406
      TabIndex        =   11
      Top             =   3444
      Width           =   1812
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF MORTGAGOR"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   1440
      Width           =   3012
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CODE"
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
      Height          =   252
      Left            =   6600
      TabIndex        =   9
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mortgage Codes"
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
      Left            =   3144
      TabIndex        =   8
      Top             =   516
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   864
      Index           =   1
      Left            =   1488
      Top             =   448
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   948
      Left            =   1488
      Top             =   360
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxMortSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim GThisRec As Integer
  Dim TempMORTCODE As String
  Dim TempBName As String
  Dim TempADD1 As String
  Dim TempADD2 As String
  Dim TempADD3 As String
  Dim TempContact As String
  Dim TempPHONE As String
  Dim TempXFileNme As String
  Dim MortName$()
  Dim MortCnt As Integer
  Dim MortRecd() As Integer
  Dim ExitOK As Boolean

Private Sub cmdClear_Click()
  If Check4Changes = True Then Exit Sub
  GThisRec = 0
  fptxtMortCode.Text = ""
  fptxtCmpnyName.Text = ""
  fptxtAdd1.Text = ""
  fptxtAdd2.Text = ""
  fptxtAdd3.Text = ""
  fptxtContact.Text = ""
  fptxtPhone.Text = "(000)-000-0000"
  fptxtXFileName.Text = ""
  Label13.Caption = "ADD MODE"
  If fptxtMortCode.Visible = True Then
    fptxtMortCode.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  If Check4Changes = True Then Exit Sub
  If ExitOK = False Then
    ExitOK = True
    Exit Sub
  End If
  frmVATaxBillSetUpMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  
  If Check4DupFileNames = False Then
    ExitOK = False
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtMortCode.Text)) = 0 Then
    Call TaxMsg(900, "Please enter a mortgage code.")
    fptxtMortCode.SetFocus
    ExitOK = False
    Exit Sub
  ElseIf Check4DupCodes = False Then
    ExitOK = False
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtCmpnyName.Text)) = 0 Then
    Call TaxMsg(900, "Please enter a company name.")
    fptxtCmpnyName.SetFocus
    ExitOK = False
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtAdd1.Text)) = 0 And Len(QPTrim$(fptxtAdd2.Text)) = 0 Then
    Call TaxMsg(900, "Please enter an address.")
    fptxtAdd1.SetFocus
    ExitOK = False
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtAdd3.Text)) = 0 Then
    Call TaxMsg(900, "Please enter a city, state and zip code.")
    fptxtAdd3.SetFocus
    ExitOK = False
    Exit Sub
  End If
   
  If Len(QPTrim$(fptxtXFileName.Text)) = 0 Then
    Call TaxMsg(900, "Please enter a unique export file name.")
    fptxtXFileName.SetFocus
    ExitOK = False
    Exit Sub
  End If
   
  OpenMortCodeFile MHandle, NumOfMCodes
  If GThisRec = 0 Then
     MortRec.MORTCODE = QPTrim$(fptxtMortCode.Text)
     MortRec.BName = QPTrim$(fptxtCmpnyName.Text)
     MortRec.Add1 = QPTrim$(fptxtAdd1.Text)
     MortRec.Add2 = QPTrim$(fptxtAdd2.Text)
     MortRec.Add3 = QPTrim$(fptxtAdd3.Text)
     MortRec.Contact = QPTrim$(fptxtContact.Text)
     MortRec.PHONE = QPTrim$(fptxtPhone.Text)
     MortRec.XFileNme = QPTrim$(fptxtXFileName.Text)
     Put MHandle, NumOfMCodes + 1, MortRec
     GThisRec = NumOfMCodes + 1
  Else
     Get MHandle, GThisRec, MortRec
       MortRec.MORTCODE = QPTrim$(fptxtMortCode.Text)
       MortRec.BName = QPTrim$(fptxtCmpnyName.Text)
       MortRec.Add1 = QPTrim$(fptxtAdd1.Text)
       MortRec.Add2 = QPTrim$(fptxtAdd2.Text)
       MortRec.Add3 = QPTrim$(fptxtAdd3.Text)
       MortRec.Contact = QPTrim$(fptxtContact.Text)
       MortRec.PHONE = QPTrim$(fptxtPhone.Text)
       MortRec.XFileNme = QPTrim$(fptxtXFileName.Text)
     Put MHandle, GThisRec, MortRec
  End If
  Close MHandle
  Call Savemsg(900, "Your data has been saved successfully.")
  Call LoadMe
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
    Case vbKeyF5:
      SendKeys "%F"
      Call cmdClear_Click
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
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxMortSetup.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxMortSetUp.")
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim x As Integer
  Dim ThisPhone As String
  
  fpList1.Clear
  ExitOK = True
'  GThisRec = 0
  Me.HelpContextID = hlpMortgageCodeM
  OpenMortCodeFile MHandle, NumOfMCodes
  If NumOfMCodes = 0 Then
    Label13.Caption = "ADD MODE"
    GThisRec = 0
    Close MHandle
    Exit Sub
  End If
  
  Label13.Caption = "EDIT MODE"
  If NumOfMCodes = 1 Then
    Get MHandle, 1, MortRec
    If MortRec.Deleted = True Then
      fptxtMortCode.Text = ""
      TempMORTCODE = ""
      fptxtCmpnyName.Text = ""
      TempBName = ""
      fptxtAdd1.Text = ""
      TempADD1 = ""
      fptxtAdd2.Text = ""
      TempADD2 = ""
      fptxtAdd3.Text = ""
      TempADD3 = ""
      fptxtContact.Text = ""
      TempContact = ""
      fptxtPhone.Text = ""
      TempPHONE = ""
      TempXFileNme = ""
      fptxtXFileName.Text = ""
    Else
      GThisRec = 1
      fptxtMortCode.Text = QPTrim$(MortRec.MORTCODE)
      TempMORTCODE = QPTrim$(MortRec.MORTCODE)
      fptxtCmpnyName.Text = QPTrim$(MortRec.BName)
      TempBName = QPTrim$(MortRec.BName)
      fptxtAdd1.Text = QPTrim$(MortRec.Add1)
      TempADD1 = QPTrim$(MortRec.Add1)
      fptxtAdd2.Text = QPTrim$(MortRec.Add2)
      TempADD2 = QPTrim$(MortRec.Add2)
      fptxtAdd3.Text = QPTrim$(MortRec.Add3)
      TempADD3 = QPTrim$(MortRec.Add3)
      fptxtContact.Text = QPTrim$(MortRec.Contact)
      TempContact = QPTrim$(MortRec.Contact)
      ThisPhone = GetPhoneNum(MortRec.PHONE)
      If Len(ThisPhone) > 0 Then
        fptxtPhone.Text = ThisPhone
        TempPHONE = ThisPhone
      Else
        fptxtPhone.Text = "(000)-000-0000"
        TempPHONE = "(000)-000-0000"
      End If
      TempXFileNme = QPTrim$(MortRec.XFileNme)
      fptxtXFileName.Text = TempXFileNme
      fpList1.InsertRow = QPTrim$(MortRec.BName) + Chr(9) + QPTrim$(MortRec.MORTCODE) + Chr(9) + CStr(1)
    End If
  Else
    Call SortEm
    fpList1.Clear
    For x = 1 To MortCnt
      Get MHandle, MortRecd(x), MortRec
      If MortRec.Deleted = True Then GoTo ItsDeleted
      If Len(fptxtMortCode.Text) > 0 Then GoTo Loaded
      If QPTrim$(MortRec.MORTCODE) <> "" Then
        GThisRec = MortRecd(x)
        fptxtMortCode.Text = QPTrim$(MortRec.MORTCODE)
        TempMORTCODE = QPTrim$(MortRec.MORTCODE)
        fptxtCmpnyName.Text = QPTrim$(MortRec.BName)
        TempBName = QPTrim$(MortRec.BName)
        fptxtAdd1.Text = QPTrim$(MortRec.Add1)
        TempADD1 = QPTrim$(MortRec.Add1)
        fptxtAdd2.Text = QPTrim$(MortRec.Add2)
        TempADD2 = QPTrim$(MortRec.Add2)
        fptxtAdd3.Text = QPTrim$(MortRec.Add3)
        TempADD3 = QPTrim$(MortRec.Add3)
        fptxtContact.Text = QPTrim$(MortRec.Contact)
        TempContact = QPTrim$(MortRec.Contact)
        ThisPhone = GetPhoneNum(MortRec.PHONE)
        If Len(ThisPhone) > 0 Then
          fptxtPhone.Text = ThisPhone
          TempPHONE = ThisPhone
        Else
          fptxtPhone.Text = "(000)-000-0000"
          TempPHONE = "(000)-000-0000"
        End If
        TempXFileNme = QPTrim$(MortRec.XFileNme)
        fptxtXFileName.Text = TempXFileNme
      End If
Loaded:
      fpList1.InsertRow = QPTrim$(MortRec.BName) + Chr(9) + QPTrim$(MortRec.MORTCODE) + Chr(9) + CStr(MortRecd(x))
ItsDeleted:
    Next x
  End If
  
  Close
End Sub
Private Function Check4Changes() As Boolean
  Dim OldStr As String
  Dim NewStr As String
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisPhone$
  
  On Error GoTo ERRORSTUFF
  
  Check4Changes = False
  ThisPhone = ReplaceString(fptxtPhone.Text, "(", "")
  ThisPhone = ReplaceString(ThisPhone, ")", "")
  ThisPhone = ReplaceString(ThisPhone, "-", "")
  If Val(ThisPhone) = 0 Then
    ThisPhone = ""
  End If
  
  If GThisRec > 0 Then
    OpenMortCodeFile MHandle, NumOfMCodes
    Get MHandle, GThisRec, MortRec
  Else
    If QPTrim$(fptxtMortCode.Text) <> "" Or QPTrim$(fptxtCmpnyName.Text) <> "" _
    Or QPTrim$(fptxtAdd1.Text) <> "" Or QPTrim$(fptxtAdd2.Text) <> "" _
    Or QPTrim$(fptxtAdd3.Text) <> "" Or QPTrim$(fptxtContact.Text) <> "" _
    Or ThisPhone <> "" Or QPTrim$(fptxtXFileName.Text) <> "" Then
      frmVATaxMsgWOpts.Label1.Caption = "Are you sure you want to exit without saving?"
      frmVATaxMsgWOpts.Label1.Top = 1000
      frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
      frmVATaxMsgWOpts.cmdExit.Text = "ESC Don't Save"
      frmVATaxMsgWOpts.Show vbModal
      If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
        Call cmdSave_Click
        Exit Function
      Else
        MainLog ("frmVATaxMortSetup: User elected to exit this screen without saving any data.")
        Exit Function
      End If
    Else
      Exit Function
    End If
  End If
  
  Set ThisControl = fptxtMortCode
  OldStr = QPTrim$(MortRec.MORTCODE)
  NewStr = QPTrim$(fptxtMortCode.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Mortgage Code' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If Check4DupCodes = False Then
        Check4Changes = True
        Close MHandle
        Exit Function
      End If
      MortRec.MORTCODE = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Mortgage Code has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtCmpnyName
  OldStr = QPTrim$(MortRec.BName)
  NewStr = QPTrim$(fptxtCmpnyName.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Company Name' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      MortRec.BName = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Company Name has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtAdd1
  OldStr = QPTrim$(MortRec.Add1)
  NewStr = QPTrim$(fptxtAdd1.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Address #1' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      MortRec.Add1 = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Address #1 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtAdd2
  OldStr = QPTrim$(MortRec.Add2)
  NewStr = QPTrim$(fptxtAdd2.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Address #2' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      MortRec.Add2 = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Address #2 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtAdd3
  OldStr = QPTrim$(MortRec.Add3)
  NewStr = QPTrim$(fptxtAdd3.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Address #3' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      MortRec.Add3 = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Address #3 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtContact
  OldStr = QPTrim$(MortRec.Contact)
  NewStr = QPTrim$(fptxtContact.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Contact' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      MortRec.Contact = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Contact has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtPhone
  OldStr = QPTrim$(ReplaceString(MortRec.PHONE, "-", ""))
  OldStr = QPTrim$(ReplaceString(OldStr, "(", ""))
  OldStr = QPTrim$(ReplaceString(OldStr, ")", ""))
  OldStr = QPTrim$(ReplaceString(OldStr, " ", ""))
  
  NewStr = QPTrim$(ReplaceString(fptxtPhone.Text, "-", ""))
  NewStr = QPTrim$(ReplaceString(NewStr, "(", ""))
  NewStr = QPTrim$(ReplaceString(NewStr, ")", ""))
  NewStr = QPTrim$(ReplaceString(NewStr, " ", ""))
  
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Telephone #' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      MortRec.PHONE = QPTrim$(fptxtPhone.Text)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Telephone # has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtXFileName
  OldStr = QPTrim$(MortRec.XFileNme)
  NewStr = QPTrim$(fptxtXFileName.Text)
  If OldStr <> NewStr Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Export File Name' field has been changed from " + OldStr + " to " + NewStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If Check4DupFileNames = False Then
        Check4Changes = True
        Close MHandle
        Exit Function
      End If
      MortRec.XFileNme = QPTrim$(NewStr)
      Put MHandle, GThisRec, MortRec
      Call Savemsg(900, "The Export File Name has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Close MHandle
  Exit Function
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close MHandle
        frmVATaxBillSetUpMenu.Show
        DoEvents
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        ThisControl.SetFocus
        Close MHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMortSetup", "Check4Changes", Erl)
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

Private Sub fpList1_Click()
  Dim ThisRec
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim ThisPhone$
  
  fpList1.Row = fpList1.ListIndex
  fpList1.Col = 2
  ThisRec = CInt(fpList1.ColText)
  OpenMortCodeFile MHandle, NumOfMCodes
  Get MHandle, ThisRec, MortRec
  Close MHandle
  GThisRec = ThisRec
  fptxtMortCode.Text = QPTrim$(MortRec.MORTCODE)
  fptxtCmpnyName.Text = QPTrim$(MortRec.BName)
  fptxtAdd1.Text = QPTrim$(MortRec.Add1)
  fptxtAdd2.Text = QPTrim$(MortRec.Add2)
  fptxtAdd3.Text = QPTrim$(MortRec.Add3)
  fptxtContact.Text = QPTrim$(MortRec.Contact)
  ThisPhone = GetPhoneNum(MortRec.PHONE)

  fptxtPhone.Text = ThisPhone
  fptxtXFileName.Text = QPTrim$(MortRec.XFileNme)
  Label13.Caption = "EDIT MODE"
  
End Sub

Private Function Check4DupCodes() As Boolean
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim x As Integer
  Dim ThisCode$
  
  Check4DupCodes = True
  ThisCode$ = QPTrim$(fptxtMortCode.Text)
  OpenMortCodeFile MHandle, NumOfMCodes
  If Label13.Caption = "ADD MODE" Then
    For x = 1 To NumOfMCodes
      Get MHandle, x, MortRec
      If QPTrim$(MortRec.MORTCODE) = ThisCode$ Then
        Exit For
      End If
    Next x
  Else
    For x = 1 To NumOfMCodes
      Get MHandle, x, MortRec
      If x = GThisRec Then GoTo SkipIt
      If QPTrim$(MortRec.MORTCODE) = ThisCode$ Then
        Exit For
      End If
SkipIt:
    Next x
  End If
  Close MHandle
  
  If x <= NumOfMCodes Then
    Check4DupCodes = False
    Call TaxMsg(800, "The mortgage code entered is already in use. Please select a different mortgage code.")
    fptxtMortCode.SetFocus
  End If
  
End Function

Private Sub SortEm()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim Big$
  Dim Lil$
  Dim x As Integer
  Dim HoldRec As Integer
  Dim HoldName As String
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim ThisName$
  
  ReDim MortRecd(1 To 1) As Integer
  ReDim MortName(1 To 1) As String
  Big = ""
  MortCnt = 0
  OpenMortCodeFile MHandle, NumOfMCodes
  For x = 1 To NumOfMCodes
    Get MHandle, x, MortRec
    If MortRec.Deleted <> 0 Then GoTo Deleted
    ThisName = QPTrim$(MortRec.BName)
    MortCnt = MortCnt + 1
    ReDim Preserve MortName(1 To MortCnt) As String
    ReDim Preserve MortRecd(1 To MortCnt) As Integer
    MortName(MortCnt) = ThisName
    MortRecd(MortCnt) = x
    If ThisName > Big Then
      Big = ThisName
    End If
Deleted:
  Next x
  
  Close MHandle
  
  Lil = Big + "z"
  Nextx = 1
  Do
    For x = Nextx To MortCnt
      ThisName = MortName(x)
      If ThisName < Lil Then
        Lil = ThisName
        Thisx = x
      End If
    Next x
    HoldName = MortName(Nextx)
    HoldRec = MortRecd(Nextx)
    MortName(Nextx) = MortName(Thisx)
    MortRecd(Nextx) = MortRecd(Thisx)
    MortName(Thisx) = HoldName
    MortRecd(Thisx) = HoldRec
    Lil = Big + "z"
    Nextx = Nextx + 1
    If Nextx > MortCnt Then Exit Do
  Loop
  
End Sub

Private Function Check4DupFileNames() As Boolean
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim x As Integer
  Dim ThisFileName$
  
  Check4DupFileNames = True
  ThisFileName$ = QPTrim$(fptxtXFileName.Text)
  If ThisFileName$ = "" Then Exit Function
  OpenMortCodeFile MHandle, NumOfMCodes
  If Label13.Caption = "ADD MODE" Then
    For x = 1 To NumOfMCodes
      Get MHandle, x, MortRec
      If QPTrim$(MortRec.XFileNme) = ThisFileName$ Then
        Exit For
      End If
    Next x
  Else
    For x = 1 To NumOfMCodes
      Get MHandle, x, MortRec
      If x = GThisRec Then GoTo SkipIt
      If QPTrim$(MortRec.XFileNme) = ThisFileName$ Then
        Exit For
      End If
SkipIt:
    Next x
  End If
  Close MHandle
  
  If x <= NumOfMCodes Then
    Check4DupFileNames = False
    Call TaxMsg(800, "The export file name entered is already in use. Please select a different export file name.")
    fptxtXFileName.SetFocus
  End If

End Function
