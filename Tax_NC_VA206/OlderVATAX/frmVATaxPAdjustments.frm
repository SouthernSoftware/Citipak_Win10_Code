VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPAdjustments 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Adjustments"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   ClipControls    =   0   'False
   Icon            =   "frmVATaxPAdjustments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAdjType 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   5880
      Width           =   4260
      _Version        =   196608
      _ExtentX        =   7514
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
      Columns         =   0
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
      BorderDropShadowWidth=   1
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   500
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxPAdjustments.frx":08CA
   End
   Begin EditLib.fpText fptxtOpt2 
      Height          =   372
      Left            =   5880
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   6072
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin VB.Timer MsgAlertTimer 
      Interval        =   50
      Left            =   10920
      Top             =   240
   End
   Begin EditLib.fpLongInteger fpLongAcctNum 
      Height          =   372
      Left            =   1932
      TabIndex        =   0
      Top             =   1320
      Width           =   1800
      _Version        =   196608
      _ExtentX        =   3175
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin EditLib.fpText fptxtState 
      Height          =   372
      Left            =   1440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4512
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1085
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpText fptxtName 
      Height          =   372
      Left            =   1440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4092
      _Version        =   196608
      _ExtentX        =   7223
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpText fptxtAddress 
      Height          =   372
      Left            =   1440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3660
      Width           =   4092
      _Version        =   196608
      _ExtentX        =   7223
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpText fptxtCity 
      Height          =   372
      Left            =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4080
      Width           =   4092
      _Version        =   196608
      _ExtentX        =   7223
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpMask fptxtZip 
      Height          =   372
      Left            =   4080
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   4512
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "#####-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtDate 
      Height          =   408
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1860
      _Version        =   196608
      _ExtentX        =   3281
      _ExtentY        =   714
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
      Text            =   "01/31/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpLongInteger fpLngIntBill 
      Height          =   372
      Left            =   1920
      TabIndex        =   19
      Top             =   2280
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "0"
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
   Begin EditLib.fpText fptxtInterest 
      Height          =   372
      Left            =   5880
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4740
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "INTEREST"
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
   Begin EditLib.fpText fptxtPers 
      Height          =   372
      Left            =   5880
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "PERSONAL"
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
   Begin EditLib.fpText fptxtMachTools 
      Height          =   372
      Left            =   5880
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2952
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "MACHINE TOOLS"
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
   Begin EditLib.fpText fptxtMerchCap 
      Height          =   372
      Left            =   5880
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3408
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "MERCHANT CAP"
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
   Begin EditLib.fpText fptxtFarmEquip 
      Height          =   372
      Left            =   5880
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3852
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "FARM EQUIPMENT"
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
   Begin EditLib.fpText fptxtMobHomes 
      Height          =   372
      Left            =   5880
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4308
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "MOBILE HOMES"
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
   Begin EditLib.fpText fptxtPenalty 
      Height          =   372
      Left            =   5880
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5196
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "PENALTY"
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
   Begin EditLib.fpCurrency fpCurrPersOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrPersAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   5
      Top             =   2520
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrMTOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2952
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrMTAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   6
      Top             =   2952
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrMCOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3408
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrMCAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   7
      Top             =   3408
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrFEOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3852
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrFEAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   8
      Top             =   3852
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrMHOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4308
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrMHAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   9
      Top             =   4308
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrIntOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4740
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrIntAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   10
      Top             =   4740
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrPenOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5196
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrPenAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   11
      Tag             =   "1"
      Top             =   5196
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrTotOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrTotAdj 
      Height          =   372
      Left            =   9720
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpText fptxtNote 
      Height          =   348
      Left            =   1440
      TabIndex        =   4
      Top             =   7116
      Width           =   3780
      _Version        =   196608
      _ExtentX        =   6667
      _ExtentY        =   609
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
      AutoCase        =   1
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
      CharValidationText=   ""
      MaxLength       =   20
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
   Begin EditLib.fpCurrency fpCurrPrepayBal 
      Height          =   372
      Left            =   6000
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1608
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrPrepayAdjBal 
      Height          =   372
      Left            =   8040
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1608
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrPrepayAdjAmt 
      Height          =   372
      Left            =   9720
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1608
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
      Height          =   372
      Left            =   3840
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBills 
      Height          =   372
      Left            =   3840
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":0DA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   2040
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6360
      Width           =   2172
      _Version        =   131072
      _ExtentX        =   3831
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":0F7F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   6906
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":1161
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   1290
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":133D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHistory 
      Height          =   495
      Left            =   5025
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":1519
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMessage 
      Height          =   492
      Left            =   3150
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1584
      _Version        =   131072
      _ExtentX        =   2794
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":16F7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReset 
      Height          =   492
      Left            =   8790
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Press to reset values to zero while maintaining the customer number and adjustment type."
      Top             =   8040
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
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
      ButtonDesigner  =   "frmVATaxPAdjustments.frx":18D5
   End
   Begin EditLib.fpText fptxtOpt1 
      Height          =   372
      Left            =   5880
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpText fptxtOpt3 
      Height          =   372
      Left            =   5880
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6528
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
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
   Begin EditLib.fpCurrency fpCurrOpt1Owed 
      Height          =   372
      Left            =   8040
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrOpt1Adj 
      Height          =   372
      Left            =   9720
      TabIndex        =   12
      Top             =   5640
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrOpt2Owed 
      Height          =   372
      Left            =   8040
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6072
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrOpt2Adj 
      Height          =   372
      Left            =   9720
      TabIndex        =   13
      Top             =   6072
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin EditLib.fpCurrency fpCurrOpt3Owed 
      Height          =   372
      Left            =   8040
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6528
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
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
   Begin EditLib.fpCurrency fpCurrOpt3Adj 
      Height          =   372
      Left            =   9720
      TabIndex        =   14
      Tag             =   "1"
      Top             =   6528
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cust Acct #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   240
      TabIndex        =   67
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   600
      Left            =   2304
      Top             =   240
      Width           =   7008
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Tax Billing Adjustments"
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
      Left            =   3402
      TabIndex        =   66
      Top             =   360
      Width           =   4848
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3000
      TabIndex        =   65
      Top             =   4608
      Width           =   852
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   360
      TabIndex        =   64
      Top             =   4608
      Width           =   852
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   360
      TabIndex        =   63
      Top             =   4188
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   240
      TabIndex        =   62
      Top             =   3768
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   360
      TabIndex        =   61
      Top             =   3360
      Width           =   852
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1080
      TabIndex        =   60
      Top             =   1896
      Width           =   612
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   360
      TabIndex        =   59
      Top             =   2370
      Width           =   1332
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7992
      X2              =   7992
      Y1              =   1200
      Y2              =   7680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   9720
      X2              =   11280
      Y1              =   7068
      Y2              =   7068
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   8040
      X2              =   9600
      Y1              =   7068
      Y2              =   7068
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6072
      TabIndex        =   58
      Top             =   7272
      Width           =   1692
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   6492
      Left            =   5760
      Top             =   1200
      Width           =   5652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1572
      Left            =   240
      Top             =   1200
      Width           =   5532
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9672
      X2              =   9672
      Y1              =   1200
      Y2              =   7680
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6360
      TabIndex        =   57
      Top             =   2196
      Width           =   1092
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   8280
      TabIndex        =   56
      Top             =   2196
      Width           =   1092
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Amt Adj"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   9960
      TabIndex        =   55
      Top             =   2196
      Width           =   1092
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   5760
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Type:"
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
      Height          =   336
      Left            =   480
      TabIndex        =   54
      Top             =   5520
      Width           =   2148
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
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
      Height          =   336
      Left            =   600
      TabIndex        =   53
      Top             =   7140
      Width           =   708
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Adjustments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   52
      Top             =   5160
      Width           =   1812
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Customer Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   51
      Top             =   2760
      Width           =   2052
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6960
      TabIndex        =   50
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Pre Adj Bal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   8112
      TabIndex        =   49
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Prepay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5760
      TabIndex        =   48
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Pre Adj Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   9744
      TabIndex        =   47
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5760
      X2              =   11400
      Y1              =   2196
      Y2              =   2196
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4932
      Left            =   240
      Top             =   2760
      Width           =   5532
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   756
      Left            =   2304
      Top             =   120
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxPAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim TempAcctNum As Long
  Dim TempBillNum As Long
  Dim TempBillRec As Double
  Dim BtnFnt As Double
  Public BillRec As Double
  Dim Release As Boolean
  Dim SaveOK As Boolean
  Dim ThisGBal As Double
  Dim ExitOK As Boolean
  Dim Invalid As Boolean
  Public RealPin As String
  Public PersPin As String
  Dim PersPaid As Double
  Dim MTPaid As Double
  Dim MCPaid As Double
  Dim FEPaid As Double
  Dim MHPaid As Double
  Dim PIntPaid As Double
  Dim PenPaid As Double
  Dim Rev1Paid As Double
  Dim Rev2Paid As Double
  Dim Rev3Paid As Double
  Dim TotPaid As Double
  Dim PersBal As Double
  Dim MTBal As Double
  Dim MCBal As Double
  Dim FEBal As Double
  Dim MHBal As Double
  Dim PIntBal As Double
  Dim PenBal As Double
  Dim Rev1Bal As Double
  Dim Rev2Bal As Double
  Dim Rev3Bal As Double
  Public ThisBillBal As Double
  Dim PersVal As Double
  Dim MTVal As Double
  Dim MCVal As Double
  Dim FEVal As Double
  Dim MHVal As Double
  Dim PIntVal As Double
  Dim PenVal As Double
  Dim Rev1Val As Double
  Dim Rev2Val As Double
  Dim Rev3Val As Double
  Dim FirstLoad As Boolean
  Public ThisBillNum As Long
  Dim NewBalThisBill As Double
  Public ThisBillType$
  Dim PayOrder() As Integer
  
Private Sub cmdBills_Click()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  If CLng(fpLongAcctNum.Value) <= 0 Then
    Call TaxMsg(900, "Please supply a valid customer account number before accessing bills.")
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  
  If Check4ValidCustNum(CLng(fpLongAcctNum.Value)) = False Then
     Call TaxMsg(800, "The customer account number entered cannot be found. Please enter a valid customer account number.")
     Close
     Exit Sub
  Else
    GCustNum = CLng(fpLongAcctNum.Value)
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, CLng(fpLongAcctNum.Value), TaxCustRec
  Close TCHandle
  If TaxCustRec.LastTrans = 0 Then
    Call TaxMsg(900, "This customer has no transactions saved.")
    Close
    Exit Sub
  End If
  
  If GCustNum = 0 Then
    Exit Sub
  End If
  
  BillRec = 0
  If GCustNum = 0 Then
    Close
    Exit Sub
  End If
  
  frmVATaxAdjustBillList.Show vbModal
  
  If BillRec > 0 Then
    TempBillRec = BillRec
    Call LoadMeBill
  ElseIf BillRec < 0 Then
    BillRec = TempBillRec
    Exit Sub
  Else
    Call TaxMsg(900, "ERROR: There was a problem loading the bill data. Please try again.")
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "cmdBills_Click", Erl)
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

End Sub

Private Sub cmdExit_Click()
  If Check4ValidCustNum(CLng(fpLongAcctNum.Value)) = True Then
    If fpCurrTotAdj.Value > 0 Then
      If TaxMsgWOpts(900, "You have made changes that are not saved. Press F10 to save these changes. Otherwise, press ESC to continue without saving.", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
        GoTo ExitNow
      Else
        Unload frmVATaxMsgWOpts
        Call cmdPost_Click
      End If
    End If
  End If
ExitNow:
  KillFile "C:\CPWork\txpadjust.dat"
  TempAcctNum = 0
  GCustNum = 0
  ExitOK = True
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  frmVATaxMsgGeneral.Label1.Caption = "1-Billing Downward Adjustment will decrease a customer's balance.  This is used to reduce an incorrect billing amount."
  frmVATaxMsgGeneral.Label2.Caption = "2-Billing Upward Adjustment will increase a customer's balance.  This is used to increase an incorrect billing amount."
  frmVATaxMsgGeneral.Label3.Caption = "3-Payment Adjustment will increase a customer's balance.  This is used if payments are entered by mistake or with an incorrect amount."
  frmVATaxMsgGeneral.Label4.Caption = "4-Release will decrease a customer's balance. This is used to reduce the amount billed because of an official governmental body directive."
  If CDbl(fpCurrPrepayBal.Value) > 0 Then
    frmVATaxMsgGeneral.Label5.Caption = "5-Prepay Adjust Down adjusts prepay balance when nothing is owed."
  End If
  frmVATaxMsgGeneral.Show vbModal
End Sub

Private Sub cmdHistory_Click()
  If GCustNum = 0 Then
    Exit Sub
  End If
  frmVATaxCustInfoTHist.Show
  DoEvents
  Me.Hide
End Sub

Private Sub cmdLookup_Click()
  frmVATaxCustLookup.Show vbModal
  DoEvents
End Sub

Private Sub cmdMessage_Click()
   If GCustNum > 0 Then
    frmVATaxMessage.Show vbModal
  End If

End Sub

Private Sub cmdPost_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer, y As Integer
  Dim SaveBillRec As Long '1/25/07
  Dim NextRec As Long '1/25/07
  
  On Error GoTo ERRORSTUFF
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
'  ReDim PayOrder(1 To 10) As Integer commented out 4/8/2010
'  PayOrder(1) = TaxMasterRec.PersPayOrder
'  PayOrder(2) = TaxMasterRec.MTPayOrder
'  PayOrder(3) = TaxMasterRec.MCPayOrder
'  PayOrder(4) = TaxMasterRec.FEPayOrder
'  PayOrder(5) = TaxMasterRec.MHPayOrder
'  PayOrder(6) = TaxMasterRec.PIntPayOrder
'  PayOrder(7) = TaxMasterRec.PPenPayOrder
'  PayOrder(8) = TaxMasterRec.POpt1PayOrder
'  PayOrder(9) = TaxMasterRec.POpt2PayOrder
'  PayOrder(10) = TaxMasterRec.POpt3PayOrder
  
  If GCustNum = 0 Then
    Call TaxMsg(900, "ERROR: No customer record number has been assigned.")
    Exit Sub
  End If
  
  If BillCnt = 0 And fpcboAdjType.Text <> "5-Prepay Adjust Down" Then 'added fpcboAdjType.Text <> "5-Prepay Adjust Down" on 11/8/07
    Call TaxMsg(900, "ERROR: This customer has no personal tax bills to adjust.")
    fpLngIntBill.SetFocus
    Exit Sub
  End If
    
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    If fpCurrPrepayAdjAmt = 0 Then
      Call TaxMsg(800, "There amount for 'Prepay Adjustment Amount' is zero. Nothing to post.")
      Close
      If fpCurrPrepayAdjAmt.Enabled = True Then
        fpCurrPrepayAdjAmt.SetFocus
      End If
      Exit Sub 'added 11/8/07
    End If
  End If
  
  If fpcboAdjType.Text <> "5-Prepay Adjust Down" Then
    If fpLngIntBill.Value = 0 Then
      Call TaxMsg(900, "Please make a selection from the bill list.")
      fpLngIntBill.SetFocus
      Exit Sub
    End If
  End If
  
  If fpcboAdjType.Text <> "5-Prepay Adjust Down" Then
    If CDbl(fpCurrTotAdj.Value) = 0 Then
      If fpcboAdjType.Text = "3-Adjustment for Payment" And CDbl(fpCurrPrepayBal.Value) <> CDbl(fpCurrPrepayAdjBal.Value) Then
         GoTo OK2Go
      Else
        Call TaxMsg(900, "The total adjustment is zero. No save required.")
        Exit Sub
      End If
    End If
  End If
  
OK2Go:
  If fpcboAdjType.Text = "1-Billing Downward Adjustment" Then
    If TaxMsgWOpts(900, "Are you ready to post this Billing Downward Adjustment transaction? Press F10 to Post. Otherwise, press ESC to abort.", "F10 POST", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  ElseIf fpcboAdjType.Text = "2-Billing Upward Adjustment" Then
    If TaxMsgWOpts(900, "Are you ready to post this Billing Upward Adjustment transaction? Press F10 to Post. Otherwise, press ESC to abort.", "F10 POST", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If BillRec > 0 And Check4PaidAmts(BillRec) = False Then
      Close
      Exit Sub
    End If
    If TaxMsgWOpts(900, "Are you ready to post this Adjustment for Payment transaction? Press F10 to Post. Otherwise, press ESC to abort.", "F10 POST", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  ElseIf fpcboAdjType.Text = "4-Release" Then
    If TaxMsgWOpts(900, "Are you ready to post this Release transaction? Press F10 to Post. Otherwise, press ESC to abort.", "F10 POST", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  ElseIf fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    If TaxMsgWOpts(900, "Are you ready to post this Prepay Adjustment Down? Press F10 to Post. Otherwise, press ESC to abort.", "F10 POST", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      fpLongAcctNum.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  Else
    Call TaxMsg(900, "Please supply an adjustment transaction type.")
    fpcboAdjType.SetFocus
    Exit Sub
  End If
  
  If fpcboAdjType.Text <> "2-Billing Upward Adjustment" And fpcboAdjType.Text <> "5-Prepay Adjust Down" Then
    If CDbl(fpCurrPersAdj.Value) > CDbl(fpCurrPersOwed.Value) Then
      Call TaxMsg(700, "The 'PERSONAL' adjustment amount is greater than the 'PERSONAL' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrPersAdj.SetFocus
      Exit Sub
    End If
   
    If CDbl(fpCurrMTAdj.Value) > CDbl(fpCurrMTOwed.Value) Then
      Call TaxMsg(700, "The 'MACHINE TOOLS' adjustment amount is greater than the 'MACHINE TOOLS' amount owed. Please revise the adjustment aamount so that it is less than or equal to the amount owed.")
      fpCurrMTAdj.SetFocus
      Exit Sub
    End If
  
    If CDbl(fpCurrMCAdj.Value) > CDbl(fpCurrMCOwed.Value) Then
      Call TaxMsg(700, "The 'MERCHANT CAP' adjustment amount is greater than the 'MERCHANT CAP' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrMCAdj.SetFocus
      Exit Sub
    End If
  
    If CDbl(fpCurrFEAdj.Value) > CDbl(fpCurrFEOwed.Value) Then
      Call TaxMsg(700, "The 'FARM EQUIPMENT' adjustment amount is greater than the 'FARM EQUIPMENT' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrFEAdj.SetFocus
      Exit Sub
    End If
  
    If CDbl(fpCurrMHAdj.Value) > CDbl(fpCurrMHOwed.Value) Then
      Call TaxMsg(700, "The 'MOBILE HOMES' adjustment amount is greater than the 'MOBILE HOMES' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrMHAdj.SetFocus
      Exit Sub
    End If
  
    If fpCurrOpt1Adj.Enabled = True Then
      If CDbl(fpCurrOpt1Adj.Value) > CDbl(fpCurrOpt1Owed.Value) Then
        Call TaxMsg(700, "The " + QPTrim$(fpTxtOpt1.Text) + " adjustment amount is greater than the " + QPTrim$(fpTxtOpt1.Text) + " amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
        fpCurrOpt1Adj.SetFocus
        Exit Sub
      End If
    End If
  
    If fpCurrOpt2Adj.Enabled = True Then
      If CDbl(fpCurrOpt2Adj.Value) > CDbl(fpCurrOpt2Owed.Value) Then
        Call TaxMsg(700, "The " + QPTrim$(fpTxtOpt2.Text) + " adjustment amount is greater than the " + QPTrim$(fpTxtOpt2.Text) + " amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
        fpCurrOpt2Adj.SetFocus
        Exit Sub
      End If
    End If
  
    If fpCurrOpt3Adj.Enabled = True Then
      If CDbl(fpCurrOpt3Adj.Value) > CDbl(fpCurrOpt3Owed.Value) Then
        Call TaxMsg(700, "The " + QPTrim$(fpTxtOpt3.Text) + " adjustment amount is greater than the " + QPTrim$(fpTxtOpt3.Text) + " amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
        fpCurrOpt3Adj.SetFocus
        Exit Sub
      End If
    End If
  End If
'----------------------------------------------------------------------
  Dim TaxAdjTrans As TaxTransactionType
  Dim TAHandle As Integer
  Dim NumOfTARecs As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim BillNum$, TotalAdj#
  Dim NextTransRec&
  Dim CreditAmt As Double
  Dim CreditBalance As Double
  Dim ThisAmt As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCustRec
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxTransFile TAHandle, NumOfTARecs
  
  Select Case QPTrim$(fpcboAdjType.Text)
    Case "1-Billing Downward Adjustment"
      GoSub AdjustBillDown
    Case "2-Billing Upward Adjustment"
      GoSub AdjustBillUp
    Case "3-Adjustment for Payment"
      GoSub AdjustPayDown
    Case "4-Release"
      GoSub Release
    Case "5-Prepay Adjust Down"
      GoSub PrepayAdjustDown
    Case Else
      Call TaxMsg(900, "ERROR: Adjustment type could not be determined.")
      Close
      fpcboAdjType.SetFocus
      Exit Sub
  End Select
  
  TaxAdjTrans.TaxYear = TaxTrans.TaxYear
  TaxAdjTrans.RealPin = 0
  TaxAdjTrans.PersPin = PersPin
  TaxAdjTrans.CustPin = TaxCustRec.PIN
  TaxAdjTrans.OperNum = OperNum
  TaxAdjTrans.BillType = "P"
  Put #TTHandle, BillRec, TaxTrans
PrePayOnly:
  NextTransRec& = (LOF(TTHandle) / Len(TaxTrans)) + 1

  TaxCustRec.LastTrans = NextTransRec&

  Put #TAHandle, NextTransRec&, TaxAdjTrans
  Put #TCHandle, GCustNum, TaxCustRec

  Close

'----------------------------------------------------------------------
  
  SaveOK = True
  Call Savemsg(900, "The adjustment transaction has been posted successfully.")
  Call MainLog("TaxAdj saved thru-TX," + QPTrim$(fpcboAdjType.Text) + ",Cust-" + Str(GCustNum) + ",for-" + Str(TotalAdj#) + ",on-" + fptxtDate.Text)
  
  If TaxMsgWOpts(900, "Do you wish to print an adjustment report? Press F10 to print the report. Otherwise, press ESC to skip the report.", "F10 Print", "ESC Don't Print") = "abort" Then
    Unload frmVATaxMsgWOpts
  Else
    Unload frmVATaxMsgWOpts
    DoEvents
    frmVATaxReportOpt.Show vbModal
    If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
      Unload frmVATaxReportOpt
      Call PrintGraphics
    ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
      frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this report."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Unload frmVATaxReportOpt
      Call PrintText
    End If
  End If
  Call Clearscreen
  
  Exit Sub
  
Release:
  'note: no overpayment amount occurs because releases only happen
  'when bill is still outstanding
  '7/12/06 changed revenues for release from charge to paid
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  TaxAdjTrans.TranType = 3              'Release
  TaxAdjTrans.Revenue.Principle1Pd = CDbl(fpCurrPersAdj.Value)
  TaxAdjTrans.Revenue.Principle2Pd = CDbl(fpCurrMTAdj.Value)
  TaxAdjTrans.Revenue.Principle3Pd = CDbl(fpCurrMCAdj.Value)
  TaxAdjTrans.Revenue.Principle4Pd = CDbl(fpCurrFEAdj.Value)
  TaxAdjTrans.Revenue.Principle5Pd = CDbl(fpCurrMHAdj.Value)
  TaxAdjTrans.Revenue.InterestPd = CDbl(fpCurrIntAdj.Value)
  TaxAdjTrans.Revenue.PenaltyPd = CDbl(fpCurrPenAdj.Value)
  TaxAdjTrans.Revenue.CollectionPd = 0
  TaxAdjTrans.Revenue.LateListPd = 0
  TaxAdjTrans.Revenue.RevOpt1Pd = CDbl(fpCurrOpt1Adj.Value)
  TaxAdjTrans.Revenue.RevOpt2Pd = CDbl(fpCurrOpt2Adj.Value)
  TaxAdjTrans.Revenue.RevOpt3Pd = CDbl(fpCurrOpt3Adj.Value)
  TaxAdjTrans.Amount = TotalAdj#
  TaxAdjTrans.CustomerRec = GCustNum
  TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
  TaxAdjTrans.BelongTo = BillRec
  SaveBillRec = BillRec '1/25/07
  NextRec = TaxCustRec.LastTrans '1/25/07
  Do While NextRec > 0 '1/25/07
    Get TTHandle, NextRec, TaxTrans 'with adjust pay down we must remove
    'the discount and we know that since discounts are only allowed when
    'a payment is made in full then only one type 2 transaction need be found
    If TaxTrans.BelongTo = SaveBillRec And TaxTrans.TranType = 2 Then
      TaxTrans.DiscAmt = 0
      Put TTHandle, NextRec, TaxTrans
      Exit Do
    End If
    NextRec = TaxTrans.LastTrans
  Loop
  Get #TTHandle, BillRec, TaxTrans
  TaxTrans.DiscAmt = 0 '1/25/07
  BillNum$ = ParseBillNum(TaxTrans.Description)
  If Len(QPTrim$(fptxtNote.Text)) = 0 Then
    TaxAdjTrans.Description = "Tax Release to Bill #" + BillNum$
  Else
    TaxAdjTrans.Description = QPTrim$(fptxtNote.Text) + ": Bill #" + BillNum$
  End If
  TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + CDbl(fpCurrPersAdj.Value))
  TaxTrans.Revenue.Principle2Pd = OldRound#(TaxTrans.Revenue.Principle2Pd + CDbl(fpCurrMTAdj.Value))
  TaxTrans.Revenue.Principle3Pd = OldRound#(TaxTrans.Revenue.Principle3Pd + CDbl(fpCurrMCAdj.Value))
  TaxTrans.Revenue.Principle4Pd = OldRound#(TaxTrans.Revenue.Principle4Pd + CDbl(fpCurrFEAdj.Value))
  TaxTrans.Revenue.Principle5Pd = OldRound#(TaxTrans.Revenue.Principle5Pd + CDbl(fpCurrMHAdj.Value))
  TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd + CDbl(fpCurrIntAdj.Value))
  TaxTrans.Revenue.PenaltyPd = OldRound#(TaxTrans.Revenue.PenaltyPd + CDbl(fpCurrPenAdj.Value))
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + CDbl(fpCurrOpt1Adj.Value))
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + CDbl(fpCurrOpt2Adj.Value))
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + CDbl(fpCurrOpt3Adj.Value))
  
  Return
  
AdjustPayDown:
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  CreditAmt = CDbl(fpCurrPrepayBal.Value)
  CreditBalance = CDbl(fpCurrPrepayBal.Value)
  If CreditAmt <= 0 Then
    TaxAdjTrans.TranType = 7
    TaxAdjTrans.Revenue.PrePaidBal = CDbl(fpCurrPrepayBal.Value)
    TaxAdjTrans.Revenue.PrePaidAmt = 0
    TaxAdjTrans.Revenue.PrePaidUsed = 0
    TaxAdjTrans.Revenue.Principle1Pd = CDbl(fpCurrPersAdj.Value)
    PersVal = CDbl(fpCurrPersAdj.Value)
    
    TaxAdjTrans.Revenue.Principle2Pd = CDbl(fpCurrMTAdj.Value)
    MTVal = CDbl(fpCurrMTAdj.Value)
    TaxAdjTrans.Revenue.Principle3Pd = CDbl(fpCurrMCAdj.Value)
    MCVal = CDbl(fpCurrMCAdj.Value)
    TaxAdjTrans.Revenue.Principle4Pd = CDbl(fpCurrFEAdj.Value)
    FEVal = CDbl(fpCurrFEAdj.Value)
    TaxAdjTrans.Revenue.Principle5Pd = CDbl(fpCurrMHAdj.Value)
    MHVal = CDbl(fpCurrMHAdj.Value)
    
    TaxAdjTrans.Revenue.InterestPd = CDbl(fpCurrIntAdj.Value)
    PIntVal = CDbl(fpCurrIntAdj.Value)
    TaxAdjTrans.Revenue.PenaltyPd = CDbl(fpCurrPenAdj.Value)
    PenVal = CDbl(fpCurrPenAdj.Value)
    
    TaxAdjTrans.Revenue.CollectionPd = 0
    TaxAdjTrans.Revenue.LateListPd = 0
    TaxAdjTrans.Revenue.RevOpt1Pd = CDbl(fpCurrOpt1Adj.Value)
    Rev1Val = CDbl(fpCurrOpt1Adj.Value)
    TaxAdjTrans.Revenue.RevOpt2Pd = CDbl(fpCurrOpt2Adj.Value)
    Rev2Val = CDbl(fpCurrOpt2Adj.Value)
    TaxAdjTrans.Revenue.RevOpt3Pd = CDbl(fpCurrOpt3Adj.Value)
    Rev3Val = CDbl(fpCurrOpt3Adj.Value)
    TaxAdjTrans.Amount = TotalAdj#
    TaxAdjTrans.CustomerRec = GCustNum
    TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
    TaxAdjTrans.BelongTo = BillRec
    Get #TTHandle, BillRec, TaxTrans
    BillNum$ = ParseBillNum(TaxTrans.Description)
    If Len(QPTrim$(fptxtNote.Text)) = 0 Then
      TaxAdjTrans.Description = "Tax Adj Pay Down #" + BillNum$
    Else
      TaxAdjTrans.Description = QPTrim$(fptxtNote.Text) + ": Bill #" + BillNum$
    End If
    SaveBillRec = BillRec '1/25/07
    NextRec = TaxCustRec.LastTrans '1/25/07
    Do While NextRec > 0 '1/25/07
      Get TTHandle, NextRec, TaxTrans 'with adjust pay down we must remove
      'the discount and we know that since discounts are only allowed when
      'a payment is made in full then only one type 2 transaction need be found
      If TaxTrans.BelongTo = SaveBillRec And TaxTrans.TranType = 2 Then
        TaxTrans.DiscAmt = 0
        Put TTHandle, NextRec, TaxTrans
        Exit Do
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    Get #TTHandle, BillRec, TaxTrans
    TaxTrans.DiscAmt = 0 '1/25/07
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd - CDbl(fpCurrPersAdj.Value))
    TaxTrans.Revenue.Principle2Pd = OldRound#(TaxTrans.Revenue.Principle2Pd - CDbl(fpCurrMTAdj.Value))
    TaxTrans.Revenue.Principle3Pd = OldRound#(TaxTrans.Revenue.Principle3Pd - CDbl(fpCurrMCAdj.Value))
    TaxTrans.Revenue.Principle4Pd = OldRound#(TaxTrans.Revenue.Principle4Pd - CDbl(fpCurrFEAdj.Value))
    TaxTrans.Revenue.Principle5Pd = OldRound#(TaxTrans.Revenue.Principle5Pd - CDbl(fpCurrMHAdj.Value))
    
    TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd - CDbl(fpCurrIntAdj.Value))
    TaxTrans.Revenue.PenaltyPd = OldRound#(TaxTrans.Revenue.PenaltyPd - CDbl(fpCurrPenAdj.Value))
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd - CDbl(fpCurrOpt1Adj.Value))
    TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd - CDbl(fpCurrOpt2Adj.Value))
    TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd - CDbl(fpCurrOpt3Adj.Value))
    NewBalThisBill = 0
    NewBalThisBill = OldRound(TaxTrans.Revenue.Collection + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Interest + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc))
  
  Else
    TaxAdjTrans.TranType = 10 'adjust pay down with a reduction in credit balance
    For x = 10 To 1 Step -1
      If x = PayOrder(10) Then
        If CDbl(fpCurrOpt3Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrOpt3Adj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.RevOpt3Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrOpt3Adj.Value))
            Rev3Val = 0
          Else
            TaxAdjTrans.Revenue.RevOpt3Pd = OldRound(CDbl(fpCurrOpt3Adj.Value) - CreditAmt)
            Rev3Val = OldRound(CDbl(fpCurrOpt3Adj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(9) Then
        If CDbl(fpCurrOpt2Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrOpt2Adj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.RevOpt2Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrOpt2Adj.Value))
            Rev2Val = 0
          Else
            TaxAdjTrans.Revenue.RevOpt2Pd = OldRound(CDbl(fpCurrOpt2Adj.Value) - CreditAmt)
            Rev2Val = OldRound(CDbl(fpCurrOpt2Adj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(8) Then
        If CDbl(fpCurrOpt1Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrOpt1Adj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.RevOpt1Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrOpt1Adj.Value))
            Rev1Val = 0
          Else
            TaxAdjTrans.Revenue.RevOpt1Pd = OldRound(CDbl(fpCurrOpt1Adj.Value) - CreditAmt)
            Rev1Val = OldRound(CDbl(fpCurrOpt1Adj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(1) Then
        If CDbl(fpCurrPersAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt >= CDbl(fpCurrPersAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle1Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPersAdj.Value))
            PersVal = 0
          Else
            TaxAdjTrans.Revenue.Principle1Pd = OldRound(CDbl(fpCurrPersAdj.Value) - CreditAmt)
            PersVal = OldRound(CDbl(fpCurrPersAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(2) Then
        If CDbl(fpCurrMTAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrMTAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle2Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrMTAdj.Value))
            MTVal = 0
          Else
            TaxAdjTrans.Revenue.Principle2Pd = OldRound(CDbl(fpCurrMTAdj.Value) - CreditAmt)
            MTVal = OldRound(CDbl(fpCurrMTAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(3) Then
        If CDbl(fpCurrMCAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrMCAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle3Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrMCAdj.Value))
            MCVal = 0
          Else
            TaxAdjTrans.Revenue.Principle3Pd = OldRound(CDbl(fpCurrMCAdj.Value) - CreditAmt)
            MCVal = OldRound(CDbl(fpCurrMCAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(4) Then
        If CDbl(fpCurrFEAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrFEAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle4Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrFEAdj.Value))
            FEVal = 0
          Else
            TaxAdjTrans.Revenue.Principle4Pd = OldRound(CDbl(fpCurrFEAdj.Value) - CreditAmt)
            FEVal = OldRound(CDbl(fpCurrFEAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(5) Then
        If CDbl(fpCurrMHAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrMHAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle5Pd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrMHAdj.Value))
            MHVal = 0
          Else
            TaxAdjTrans.Revenue.Principle5Pd = OldRound(CDbl(fpCurrMHAdj.Value) - CreditAmt)
            MHVal = OldRound(CDbl(fpCurrMHAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(6) Then
        If CDbl(fpCurrIntAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrIntAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.InterestPd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrIntAdj.Value))
            PIntVal = 0
          Else
            TaxAdjTrans.Revenue.InterestPd = OldRound(CDbl(fpCurrIntAdj.Value) - CreditAmt)
            PIntVal = OldRound(CDbl(fpCurrIntAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(7) Then
        If CDbl(fpCurrPenAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrPenAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.PenaltyPd = 0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPenAdj.Value))
            PenVal = 0
          Else
            TaxAdjTrans.Revenue.PenaltyPd = OldRound(CDbl(fpCurrPenAdj.Value) - CreditAmt)
            PenVal = OldRound(CDbl(fpCurrPenAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      End If
    Next x
    
    TaxAdjTrans.Revenue.PrePaidAmt = 0
    
    If CreditAmt = 0 Then
      TaxAdjTrans.Revenue.PrePaidUsed = CreditBalance
    Else
      TaxAdjTrans.Revenue.PrePaidUsed = OldRound(CreditBalance - CreditAmt)
    End If
    
    TaxAdjTrans.Revenue.PrePaidBal = OldRound(CreditBalance - TaxAdjTrans.Revenue.PrePaidUsed)
    TaxAdjTrans.Amount = TotalAdj#
    TaxAdjTrans.Description = "Tax Adj Pay Dn With Credit Applied #" + BillNum$
    TaxAdjTrans.CustomerRec = GCustNum
    TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
    TaxAdjTrans.BelongTo = BillRec
 
    Get #TTHandle, BillRec, TaxTrans
    BillNum$ = ParseBillNum(TaxTrans.Description)
    For x = 1 To 10
      If x = PayOrder(2) Then
        ThisAmt = OldRound(CDbl(fpCurrMTOwed.Value) - MTVal)
        TaxTrans.Revenue.Principle2Pd = ThisAmt
      ElseIf x = PayOrder(1) Then
        ThisAmt = OldRound(CDbl(fpCurrPersOwed.Value) - PersVal)
        TaxTrans.Revenue.Principle1Pd = ThisAmt
      ElseIf x = PayOrder(3) Then
        ThisAmt = OldRound(CDbl(fpCurrMCOwed.Value) - MCVal)
        TaxTrans.Revenue.Principle3Pd = ThisAmt
      ElseIf x = PayOrder(4) Then
        ThisAmt = OldRound(CDbl(fpCurrFEOwed.Value) - FEVal)
        TaxTrans.Revenue.Principle4Pd = ThisAmt
      ElseIf x = PayOrder(5) Then
        ThisAmt = OldRound(CDbl(fpCurrMHOwed.Value) - MHVal)
        TaxTrans.Revenue.Principle5Pd = ThisAmt
      ElseIf x = PayOrder(6) Then
        ThisAmt = OldRound(CDbl(fpCurrIntOwed.Value) - PIntVal)
        TaxTrans.Revenue.InterestPd = ThisAmt
      ElseIf x = PayOrder(7) Then
        ThisAmt = OldRound(CDbl(fpCurrPenOwed.Value) - PenVal)
        TaxTrans.Revenue.PenaltyPd = ThisAmt
      ElseIf x = PayOrder(8) Then
        ThisAmt = OldRound(CDbl(fpCurrOpt1Owed.Value) - Rev1Val)
        TaxTrans.Revenue.RevOpt1Pd = ThisAmt
      ElseIf x = PayOrder(9) Then
        ThisAmt = OldRound(CDbl(fpCurrOpt2Owed.Value) - Rev2Val)
        TaxTrans.Revenue.RevOpt2Pd = ThisAmt
      ElseIf x = PayOrder(10) Then
        ThisAmt = OldRound(CDbl(fpCurrOpt3Owed.Value) - Rev3Val)
        TaxTrans.Revenue.RevOpt3Pd = ThisAmt
      End If
    Next x
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.LateListPd = 0
    NewBalThisBill = 0
    NewBalThisBill = OldRound(TaxTrans.Revenue.Collection + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Interest + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc))
  End If
  
  Return

AdjustBillDown:
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  TaxAdjTrans.TranType = 13 'adjust bill down not affecting credit balance
  TaxAdjTrans.Revenue.PrePaidAmt = 0
  TaxAdjTrans.Revenue.PrePaidBal = CDbl(fpCurrPrepayBal.Value)
  TaxAdjTrans.Revenue.PrePaidUsed = 0
  TaxAdjTrans.Revenue.Principle1 = CDbl(fpCurrPersAdj.Value)
  TaxAdjTrans.Revenue.Principle2 = CDbl(fpCurrMTAdj.Value)
  TaxAdjTrans.Revenue.Principle3 = CDbl(fpCurrMCAdj.Value)
  TaxAdjTrans.Revenue.Principle4 = CDbl(fpCurrFEAdj.Value)
  TaxAdjTrans.Revenue.Principle5 = CDbl(fpCurrMHAdj.Value)
  
  TaxAdjTrans.Revenue.Interest = CDbl(fpCurrIntAdj.Value)
  TaxAdjTrans.Revenue.Penalty = CDbl(fpCurrPenAdj.Value)
  
  TaxAdjTrans.Revenue.Collection = 0
  TaxAdjTrans.Revenue.LateList = 0
  TaxAdjTrans.Revenue.RevOpt1 = CDbl(fpCurrOpt1Adj.Value)
  TaxAdjTrans.Revenue.RevOpt2 = CDbl(fpCurrOpt2Adj.Value)
  TaxAdjTrans.Revenue.RevOpt3 = CDbl(fpCurrOpt3Adj.Value)
  TaxAdjTrans.Amount = TotalAdj#
  TaxAdjTrans.CustomerRec = GCustNum
  TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
  TaxAdjTrans.BelongTo = BillRec
    
  Get #TTHandle, BillRec, TaxTrans
  BillNum$ = ParseBillNum(TaxTrans.Description)
  If Len(QPTrim$(fptxtNote.Text)) = 0 Then
    TaxAdjTrans.Description = "Tax Adj Bill Down #" + BillNum$
  Else
    TaxAdjTrans.Description = QPTrim$(fptxtNote.Text) + ": Bill #" + BillNum$
  End If
  TaxTrans.Revenue.Principle1 = OldRound#(TaxTrans.Revenue.Principle1 - CDbl(fpCurrPersAdj.Value))
  TaxTrans.Revenue.Principle2 = OldRound#(TaxTrans.Revenue.Principle2 - CDbl(fpCurrMTAdj.Value))
  TaxTrans.Revenue.Principle3 = OldRound#(TaxTrans.Revenue.Principle3 - CDbl(fpCurrMCAdj.Value))
  TaxTrans.Revenue.Principle4 = OldRound#(TaxTrans.Revenue.Principle4 - CDbl(fpCurrFEAdj.Value))
  TaxTrans.Revenue.Principle5 = OldRound#(TaxTrans.Revenue.Principle5 - CDbl(fpCurrMHAdj.Value))
  
  TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest - CDbl(fpCurrIntAdj.Value))
  TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty - CDbl(fpCurrPenAdj.Value))
  
  TaxTrans.Revenue.Collection = 0
  TaxTrans.Revenue.LateList = 0
  TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 - CDbl(fpCurrOpt1Adj.Value))
  TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 - CDbl(fpCurrOpt2Adj.Value))
  TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 - CDbl(fpCurrOpt3Adj.Value))
  
  Return
  
AdjustBillUp:
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  CreditAmt = CDbl(fpCurrPrepayBal.Value)
  CreditBalance = CDbl(fpCurrPrepayBal.Value)
  If CreditAmt <= 0 Then
    TaxAdjTrans.TranType = 14 'adjust bill up with no affect on credit balance
    TaxAdjTrans.Revenue.Collection = 0
    TaxAdjTrans.Revenue.LateList = 0

    TaxAdjTrans.Revenue.Principle1 = CDbl(fpCurrPersAdj.Value)
    PersVal = CDbl(fpCurrPersAdj.Value)
    TaxAdjTrans.Revenue.Principle2 = CDbl(fpCurrMTAdj.Value)
    MTVal = CDbl(fpCurrMTAdj.Value)
    TaxAdjTrans.Revenue.Principle3 = CDbl(fpCurrMCAdj.Value)
    MCVal = CDbl(fpCurrMCAdj.Value)
    TaxAdjTrans.Revenue.Principle4 = CDbl(fpCurrFEAdj.Value)
    FEVal = CDbl(fpCurrFEAdj.Value)
    TaxAdjTrans.Revenue.Principle5 = CDbl(fpCurrMHAdj.Value)
    MHVal = CDbl(fpCurrMHAdj.Value)
    
    TaxAdjTrans.Revenue.Interest = CDbl(fpCurrIntAdj.Value)
    PIntVal = CDbl(fpCurrIntAdj.Value)
    TaxAdjTrans.Revenue.Penalty = CDbl(fpCurrPenAdj.Value)
    PenVal = CDbl(fpCurrPenAdj.Value)
    
    TaxAdjTrans.Revenue.RevOpt1 = CDbl(fpCurrOpt1Adj.Value)
    Rev1Val = CDbl(fpCurrOpt1Adj.Value)
    TaxAdjTrans.Revenue.RevOpt2 = CDbl(fpCurrOpt2Adj.Value)
    Rev2Val = CDbl(fpCurrOpt2Adj.Value)
    TaxAdjTrans.Revenue.RevOpt3 = CDbl(fpCurrOpt3Adj.Value)
    Rev3Val = CDbl(fpCurrOpt3Adj.Value)
    
    TaxAdjTrans.Amount = TotalAdj#
    TaxAdjTrans.CustomerRec = GCustNum
    TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
    TaxAdjTrans.BelongTo = BillRec
    SaveBillRec = BillRec '1/25/07
    NextRec = TaxCustRec.LastTrans '1/25/07
    Do While NextRec > 0 '1/25/07
      Get TTHandle, NextRec, TaxTrans 'with adjust pay down we must remove
      'the discount and we know that since discounts are only allowed when
      'a payment is made in full then only one type 2 transaction need be found
      If TaxTrans.BelongTo = SaveBillRec And TaxTrans.TranType = 2 Then
        TaxTrans.DiscAmt = 0
        Put TTHandle, NextRec, TaxTrans
        Exit Do
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    Get #TTHandle, BillRec, TaxTrans
    TaxTrans.DiscAmt = 0 '1/25/07
    BillNum$ = ParseBillNum(TaxTrans.Description)
    If Len(QPTrim$(fptxtNote.Text)) = 0 Then
      TaxAdjTrans.Description = "Tax Adj Bill Up #" + BillNum$
    Else
      TaxAdjTrans.Description = QPTrim$(fptxtNote.Text) + ": Bill #" + BillNum$
    End If
    TaxTrans.Revenue.Principle1 = OldRound#(TaxTrans.Revenue.Principle1 + CDbl(fpCurrPersAdj.Value))
    TaxTrans.Revenue.Principle2 = OldRound#(TaxTrans.Revenue.Principle2 + CDbl(fpCurrMTAdj.Value))
    TaxTrans.Revenue.Principle3 = OldRound#(TaxTrans.Revenue.Principle3 + CDbl(fpCurrMCAdj.Value))
    TaxTrans.Revenue.Principle4 = OldRound#(TaxTrans.Revenue.Principle4 + CDbl(fpCurrFEAdj.Value))
    TaxTrans.Revenue.Principle5 = OldRound#(TaxTrans.Revenue.Principle5 + CDbl(fpCurrMHAdj.Value))
    
    TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + CDbl(fpCurrIntAdj.Value))
    TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty + CDbl(fpCurrPenAdj.Value))
    
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 + CDbl(fpCurrOpt1Adj.Value))
    TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 + CDbl(fpCurrOpt2Adj.Value))
    TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 + CDbl(fpCurrOpt3Adj.Value))
    NewBalThisBill = 0
    NewBalThisBill = OldRound(TaxTrans.Revenue.Collection + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Interest + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc))
  Else
    TaxAdjTrans.TranType = 24 'adjust bill up with a reduction in credit balance
    For x = 1 To 10
      If x = PayOrder(1) Then
        If CDbl(fpCurrPersAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrPersAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle1Pd = CDbl(fpCurrPersAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPersAdj.Value))
          Else
            TaxAdjTrans.Revenue.Principle1Pd = CreditAmt 'OldRound(CDbl(fpCurrIntAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(2) Then
        If CDbl(fpCurrMTAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrMTAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle2Pd = CDbl(fpCurrMTAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrMTAdj.Value))
          Else
            TaxAdjTrans.Revenue.Principle2Pd = CreditAmt 'OldRound(CDbl(fpCurrAdvColAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(3) Then
        If CDbl(fpCurrMCAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrMCAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle3Pd = CDbl(fpCurrMCAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrMCAdj.Value))
          Else
            TaxAdjTrans.Revenue.Principle3Pd = CreditAmt 'OldRound(CDbl(fpCurrLateListAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(4) Then
        If CDbl(fpCurrFEAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrFEAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle4Pd = CDbl(fpCurrFEAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrFEAdj.Value))
          Else
            TaxAdjTrans.Revenue.Principle4Pd = CreditAmt 'OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(5) Then
        If CDbl(fpCurrMHAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrMHAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.Principle5Pd = CDbl(fpCurrMHAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrMHAdj.Value))
          Else
            TaxAdjTrans.Revenue.Principle5Pd = CreditAmt 'OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(6) Then
        If CDbl(fpCurrIntAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrIntAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.InterestPd = CDbl(fpCurrIntAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrIntAdj.Value))
          Else
            TaxAdjTrans.Revenue.InterestPd = CreditAmt 'OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(7) Then
        If CDbl(fpCurrPenAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrPenAdj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.PenaltyPd = CDbl(fpCurrPenAdj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPenAdj.Value))
          Else
            TaxAdjTrans.Revenue.PenaltyPd = CreditAmt 'OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(8) Then
        If CDbl(fpCurrOpt1Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrOpt1Adj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.RevOpt1Pd = CDbl(fpCurrOpt1Adj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrOpt1Adj.Value))
          Else
            TaxAdjTrans.Revenue.RevOpt1Pd = CreditAmt 'OldRound(CDbl(fpCurrOpt1Adj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(9) Then
        If CDbl(fpCurrOpt2Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrOpt2Adj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.RevOpt2Pd = CDbl(fpCurrOpt2Adj.Value) '0
           CreditAmt = OldRound(CreditAmt - CDbl(fpCurrOpt2Adj.Value))
          Else
            TaxAdjTrans.Revenue.RevOpt2Pd = CreditAmt 'OldRound(CDbl(fpCurrOpt2Adj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      ElseIf x = PayOrder(10) Then
        If CDbl(fpCurrOpt3Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
          If CreditAmt > CDbl(fpCurrOpt3Adj.Value) Then 'if the credit available can cover this adjustment
            TaxAdjTrans.Revenue.RevOpt3Pd = CDbl(fpCurrOpt3Adj.Value) '0
            CreditAmt = OldRound(CreditAmt - CDbl(fpCurrOpt3Adj.Value))
          Else
            TaxAdjTrans.Revenue.RevOpt3Pd = CreditAmt 'OldRound(CDbl(fpCurrOpt3Adj.Value) - CreditAmt)
            CreditAmt = 0
          End If
        End If
      End If
    Next x
    
    TaxAdjTrans.Revenue.PrePaidAmt = 0
    If CreditAmt = 0 Then
      TaxAdjTrans.Revenue.PrePaidUsed = CreditBalance
    Else
      TaxAdjTrans.Revenue.PrePaidUsed = OldRound(CreditBalance - CreditAmt)
    End If
    
    TaxAdjTrans.Revenue.PrePaidBal = OldRound(CreditBalance - TaxAdjTrans.Revenue.PrePaidUsed)
    TaxAdjTrans.Amount = OldRound(TotalAdj# - CDbl(fpCurrTotOwed.Value))
'    TaxAdjTrans.Description = "Tax Adj Bill Up With Credit Applied#" + BillNum$
    TaxAdjTrans.CustomerRec = GCustNum
    TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
    TaxAdjTrans.BelongTo = BillRec
    
    Get #TTHandle, BillRec, TaxTrans
    BillNum$ = ParseBillNum(TaxTrans.Description)
    TaxAdjTrans.Description = "Adj Bill Up/Credit Used #" + BillNum$
    
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.LateList = 0
    For x = 1 To 10
      If x = PayOrder(1) Then
        TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + TaxAdjTrans.Revenue.Principle1Pd)
        TaxTrans.Revenue.Principle1 = OldRound#(TaxTrans.Revenue.Principle1 + CDbl(fpCurrPersAdj.Value))
      ElseIf x = PayOrder(2) Then
        TaxTrans.Revenue.Principle2Pd = OldRound#(TaxTrans.Revenue.Principle2Pd + TaxAdjTrans.Revenue.Principle2Pd)
        TaxTrans.Revenue.Principle2 = OldRound#(TaxTrans.Revenue.Principle2 + CDbl(fpCurrMTAdj.Value))
      ElseIf x = PayOrder(3) Then
        TaxTrans.Revenue.Principle3Pd = OldRound#(TaxTrans.Revenue.Principle3Pd + TaxAdjTrans.Revenue.Principle3Pd)
        TaxTrans.Revenue.Principle3 = OldRound#(TaxTrans.Revenue.Principle3 + CDbl(fpCurrMCAdj.Value))
      ElseIf x = PayOrder(4) Then
        TaxTrans.Revenue.Principle4Pd = OldRound#(TaxTrans.Revenue.Principle4Pd + TaxAdjTrans.Revenue.Principle4Pd)
        TaxTrans.Revenue.Principle4 = OldRound#(TaxTrans.Revenue.Principle4 + CDbl(fpCurrFEAdj.Value))
      ElseIf x = PayOrder(5) Then
        TaxTrans.Revenue.Principle5Pd = OldRound#(TaxTrans.Revenue.Principle5Pd + TaxAdjTrans.Revenue.Principle5Pd)
        TaxTrans.Revenue.Principle5 = OldRound#(TaxTrans.Revenue.Principle5 + CDbl(fpCurrMHAdj.Value))
      ElseIf x = PayOrder(6) Then
        TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd + TaxAdjTrans.Revenue.InterestPd)
        TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + CDbl(fpCurrIntAdj.Value))
      ElseIf x = PayOrder(7) Then
        TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd + TaxAdjTrans.Revenue.PenaltyPd)
        TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty + CDbl(fpCurrPenAdj.Value))
      ElseIf x = PayOrder(8) Then
        TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd + TaxAdjTrans.Revenue.RevOpt1Pd)
        TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 + CDbl(fpCurrOpt1Adj.Value))
      ElseIf x = PayOrder(9) Then
        TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd + TaxAdjTrans.Revenue.RevOpt2Pd)
        TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 + CDbl(fpCurrOpt2Adj.Value))
      ElseIf x = PayOrder(10) Then
        TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd + TaxAdjTrans.Revenue.RevOpt3Pd)
        TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 + CDbl(fpCurrOpt3Adj.Value))
      End If
    Next x
    
    NewBalThisBill = 0
    NewBalThisBill = OldRound(TaxTrans.Revenue.Collection + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Interest + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
    NewBalThisBill = OldRound(NewBalThisBill + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd))
    NewBalThisBill = OldRound(NewBalThisBill - (TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc))
  End If
  
  Return
  
PrepayAdjustDown:
  TaxAdjTrans.TranType = 11
  TaxAdjTrans.TaxYear = TaxMasterRec.RTaxYear
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  TaxAdjTrans.Amount = fpCurrPrepayAdjAmt
  TaxAdjTrans.BelongTo = 0
  TaxAdjTrans.BillType = "P" 'added P (was just " ") 11/8/2007
  TaxAdjTrans.CustomerRec = GCustNum
  TaxAdjTrans.CustPin = TaxCustRec.PIN
  If Len(QPTrim$(fptxtNote.Text)) = 0 Then
    TaxAdjTrans.Description = "Tax Adj PrePay Down #"
  Else
    TaxAdjTrans.Description = QPTrim$(fptxtNote.Text) + " Adj Prepay Down"
  End If
  TaxAdjTrans.DiscAmt = 0
  TaxAdjTrans.DiscXDate = 0
  TaxAdjTrans.DMVBatch = 0
  TaxAdjTrans.DMVSubmitted = ""
  TaxAdjTrans.FromPrePay = 0
  TaxAdjTrans.InternalPin = 0
  TaxAdjTrans.LastTrans = TaxCustRec.LastTrans
  TaxAdjTrans.Revenue.pad = ""
  TaxAdjTrans.Revenue.Collection = 0
  TaxAdjTrans.Revenue.CollectionPd = 0
  TaxAdjTrans.Revenue.Future1 = 0
  TaxAdjTrans.Revenue.Future1Pd = 0
  TaxAdjTrans.Revenue.Future2 = 0
  TaxAdjTrans.Revenue.Future2Pd = 0
  TaxAdjTrans.Revenue.Interest = 0
  TaxAdjTrans.Revenue.InterestPd = 0
  TaxAdjTrans.Revenue.LateList = 0
  TaxAdjTrans.Revenue.LateListPd = 0
  TaxAdjTrans.Revenue.Penalty = 0
  TaxAdjTrans.Revenue.PenaltyPd = 0
  TaxAdjTrans.Revenue.PrePaidAmt = 0
  TaxAdjTrans.Revenue.PrePaidBal = OldRound(CDbl(fpCurrPrepayAdjBal.Value) - CDbl(fpCurrPrepayAdjAmt.Value))
  TaxAdjTrans.Revenue.PrePaidUsed = CDbl(fpCurrPrepayAdjAmt.Value)
  TaxAdjTrans.Revenue.Principle1 = 0
  TaxAdjTrans.Revenue.Principle1Pd = 0
  TaxAdjTrans.Revenue.Principle2 = 0
  TaxAdjTrans.Revenue.Principle2Pd = 0
  TaxAdjTrans.Revenue.Principle3 = 0
  TaxAdjTrans.Revenue.Principle3Pd = 0
  TaxAdjTrans.Revenue.Principle4 = 0
  TaxAdjTrans.Revenue.Principle4Pd = 0
  TaxAdjTrans.Revenue.Principle5 = 0
  TaxAdjTrans.Revenue.Principle5Pd = 0
  TaxAdjTrans.Revenue.RevOpt1 = 0
  TaxAdjTrans.Revenue.RevOpt1Pd = 0
  TaxAdjTrans.Revenue.RevOpt2 = 0
  TaxAdjTrans.Revenue.RevOpt2Pd = 0
  TaxAdjTrans.Revenue.RevOpt3 = 0
  TaxAdjTrans.Revenue.RevOpt3Pd = 0
  TaxCustRec.PrePayBal = CDbl(fpCurrPrepayAdjBal.Value)
  GoTo PrePayOnly
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "cmdPost_Click", Erl)
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

  
End Sub

Private Sub cmdReset_Click()
  Dim ThisCust As Long
  Dim ThisType$
  
  Call ClearUserInput
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdLookup_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%B"
      Call cmdBills_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPost_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%R"
      Call cmdReset_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%H"
      Call cmdHistory_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%M"
      Call cmdMessage_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Release = False
  ThisGBal = 0
  Invalid = False
  ExitOK = False
  FirstLoad = True
  Call LoadMe
  FirstLoad = False
  Me.HelpContextID = hlpTaxBilling
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "C:\CPWork\txpadjust.dat"
      GCustNum = 0
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPAdjustments.")
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

Private Sub LoadMe()
  Dim One As Integer
  Dim AHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisDesc As String
  
  On Error GoTo ERRORSTUFF
  
  ThisBillType = "R"
  fpCurrPrepayAdjAmt.ControlType = ControlTypeReadOnly
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  MsgAlertTimer.Enabled = False
  ThisDesc = QPTrim$(TaxMasterRec.POptRev1)
  If ThisDesc <> "" Then
    fpTxtOpt1.Text = QPTrim$(TaxMasterRec.POptRev1)
  Else
    fpTxtOpt1.Text = "NOT IN USE"
    fpCurrOpt1Adj.Enabled = False
  End If
  
  ThisDesc = QPTrim$(TaxMasterRec.POptRev2)
  If ThisDesc <> "" Then
    fpTxtOpt2.Text = QPTrim$(TaxMasterRec.POptRev2)
  Else
    fpTxtOpt2.Text = "NOT IN USE"
    fpCurrOpt2Adj.Enabled = False
  End If
  
  ThisDesc = QPTrim$(TaxMasterRec.POptRev3)
  If ThisDesc <> "" Then
    fpTxtOpt3.Text = QPTrim$(TaxMasterRec.POptRev3)
  Else
    fpTxtOpt3.Text = "NOT IN USE"
    fpCurrOpt3Adj.Enabled = False
  End If
  
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\txpadjust.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  fptxtDate.Text = Date
  fpcboAdjType.AddItem "1-Billing Downward Adjustment"
  fpcboAdjType.AddItem "2-Billing Upward Adjustment"
  fpcboAdjType.AddItem "3-Adjustment for Payment" 'Downward Adjustment"
  fpcboAdjType.AddItem "4-Release"
  fpcboAdjType.ListIndex = 0
  TempAcctNum = 0
  
  ReDim PayOrder(1 To 10) As Integer 'moved from cmdPost 4/8/2010
  PayOrder(1) = TaxMasterRec.PersPayOrder
  PayOrder(2) = TaxMasterRec.MTPayOrder
  PayOrder(3) = TaxMasterRec.MCPayOrder
  PayOrder(4) = TaxMasterRec.FEPayOrder
  PayOrder(5) = TaxMasterRec.MHPayOrder
  PayOrder(6) = TaxMasterRec.PIntPayOrder
  PayOrder(7) = TaxMasterRec.PPenPayOrder
  PayOrder(8) = TaxMasterRec.POpt1PayOrder
  PayOrder(9) = TaxMasterRec.POpt2PayOrder
  PayOrder(10) = TaxMasterRec.POpt3PayOrder
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "LoadMe", Erl)
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

End Sub

Public Sub LoadMeEdit()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim LoadThis$
  Dim OpNum$
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  OpNum = ""
  
  If Check4CustInPayBatch(fpLongAcctNum.Value, OpNum, BillType) = True Then
    If BillType = "R" Then GoTo GoAhead
    Call TaxMsg(700, "This customer is included in a personal payment file for operator #" + OpNum + " that has not been posted. Please either post this payment or delete this customer from the payment file.")
    Call Clearscreen
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  
GoAhead:
  MsgAlertTimer.Enabled = False
  If CLng(fpLongAcctNum.Text) <= 0 Then Exit Sub
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, fpLongAcctNum.Value, TaxCustRec
  Close TCHandle
  fptxtName.Text = QPTrim$(TaxCustRec.CustName)
  LoadThis = QPTrim$(TaxCustRec.Addr1)
  If LoadThis <> "" Then
    fptxtAddress.Text = LoadThis
  Else
    fptxtAddress.Text = QPTrim$(TaxCustRec.Addr2)
  End If
  fptxtCity.Text = QPTrim$(TaxCustRec.City)
  fptxtState.Text = QPTrim$(TaxCustRec.State)
  fptxtZip.Text = QPTrim$(TaxCustRec.Zip)
  GCustNum = CLng(fpLongAcctNum.Value)
  If GCustNum > 0 Then
    ThisGBal = GetCustPersBalance(GCustNum, -1)
  End If
  If CustHasMsg(GCustNum) Then
    MsgAlertTimer.Enabled = True
  Else
    MsgAlertTimer.Enabled = False
    cmdMessage.ForeColor = &H80000012
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "LoadMeEdit", Erl)
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

End Sub

Private Sub fpcboAdjType_Change()
  Dim AcctNum As Long
  Dim PrePayBal As Double
  Dim PrePayAdjBal As Double
  Dim ThisName$
  Dim ThisAdd1$
  Dim ThisCity$
  Dim ThisState$
  Dim ThisZip$
  
  If FirstLoad = True Then
    Exit Sub
  End If
  
  Call ClearUserInput
  Label10.Caption = "Balance"
  fpCurrPersOwed = PersBal
  fpCurrMCOwed = MCBal
  fpCurrMTOwed = MTBal
  fpCurrFEOwed = FEBal
  fpCurrMHOwed = MHBal
  fpCurrIntOwed = PIntBal
  fpCurrPenOwed = PenBal
  fpCurrOpt1Owed = Rev1Bal
  fpCurrOpt2Owed = Rev2Bal
  fpCurrOpt3Owed = Rev3Bal
  fpCurrTotOwed = ThisBillBal
  
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    fpCurrPrepayAdjAmt.ControlType = ControlTypeNormal
  Else
    fpCurrPrepayAdjAmt.ControlType = ControlTypeReadOnly
  End If
  
  If fpcboAdjType.Text <> "5-Prepay Adjust Down" And fpcboAdjType.Text <> "3-Adjustment for Payment" Then
    cmdBills.Enabled = True
    fpCurrPrepayAdjAmt = 0
    fpCurrPrepayAdjAmt.Enabled = False
    fpCurrPersAdj.Enabled = True
    fpCurrMTAdj.Enabled = True
    fpCurrMCAdj.Enabled = True
    fpCurrFEAdj.Enabled = True
    fpCurrMHAdj.Enabled = True
    fpCurrIntAdj.Enabled = True
    fpCurrPenAdj.Enabled = True
    If QPTrim$(fpTxtOpt1.Text) <> "NOT IN USE" Then
      fpCurrOpt1Adj.Enabled = True
    End If
    If QPTrim$(fpTxtOpt2.Text) <> "NOT IN USE" Then
      fpCurrOpt2Adj.Enabled = True
    End If
    If QPTrim$(fpTxtOpt3.Text) <> "NOT IN USE" Then
      fpCurrOpt3Adj.Enabled = True
    End If
  ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then
    Label10.Caption = "Paid"
    fpCurrPersOwed = PersPaid
    fpCurrMCOwed = MCPaid
    fpCurrMTOwed = MTPaid
    fpCurrFEOwed = FEPaid
    fpCurrMHOwed = MHPaid
    fpCurrIntOwed = PIntPaid
    fpCurrPenOwed = PenPaid
    fpCurrOpt1Owed = Rev1Paid
    fpCurrOpt2Owed = Rev2Paid
    fpCurrOpt3Owed = Rev3Paid
    fpCurrTotOwed = TotPaid
  Else
    cmdBills.Enabled = False
    fpLngIntBill = 0
    fpCurrPrepayAdjAmt.Enabled = True
    fpCurrPersAdj.Enabled = False
    fpCurrMTAdj.Enabled = False
    fpCurrMCAdj.Enabled = False
    fpCurrFEAdj.Enabled = False
    fpCurrMHAdj.Enabled = False
    fpCurrIntAdj.Enabled = False
    fpCurrPenAdj.Enabled = False
    fpCurrOpt1Adj.Enabled = False
    fpCurrOpt2Adj.Enabled = False
    fpCurrOpt3Adj.Enabled = False
  End If
End Sub

Private Sub fpcboAdjType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAdjType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAdjType.ListIndex = -1
  End If
  If fpcboAdjType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtNote.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcboAdjType_LostFocus()
  If fpcboAdjType = "5-Prepay Adjust Down" Then Exit Sub
  
  If fpLongAcctNum = 0 Then Exit Sub
  
  If fpLngIntBill = 0 Then
    If frmVATaxAdjustBillList.Visible = True Then Exit Sub
    Call TaxMsg(900, "Please select a bill first.")
    fpcboAdjType = "1-Billing Downward Adjustment"
    fpLngIntBill.SetFocus
    Exit Sub
  End If

End Sub

Private Sub fpCurrMCAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrMCAdj.Value) > CDbl(fpCurrMCOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrMCAdj = CDbl(fpCurrMCOwed.Value)
      fpCurrMCAdj.SetFocus
    End If
  End If
  Call FigureAdjCol(3)
End Sub

Private Sub fpCurrIntAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrIntAdj.Value) > CDbl(fpCurrIntOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrIntAdj = CDbl(fpCurrIntOwed.Value)
      fpCurrIntAdj.SetFocus
    End If
  End If
  Call FigureAdjCol(6)
End Sub

Private Sub fpCurrFEAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrFEAdj.Value) > CDbl(fpCurrFEOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrFEAdj = CDbl(fpCurrFEOwed.Value)
      fpCurrFEAdj.SetFocus
    End If
  End If
  Call FigureAdjCol(4)
End Sub

Private Sub fpCurrMHAdj_LostFocus() '1/25/07
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrMHAdj.Value) > CDbl(fpCurrMHOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrMHAdj = CDbl(fpCurrMHOwed.Value)
      fpCurrMHAdj.SetFocus
    End If
  End If
  Call FigureAdjCol(5)

End Sub

Private Sub fpCurrPenAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrPenAdj.Value) > CDbl(fpCurrPenOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrPenAdj = CDbl(fpCurrPenOwed.Value)
      fpCurrPenAdj.SetFocus
    End If
  End If
  
  Call FigureAdjCol(7)

End Sub

Private Sub fpCurrPrepayAdjAmt_LostFocus()
  fpCurrPrepayAdjBal = OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjAmt.Value))
  If fpCurrPrepayAdjBal < 0 Then
    Call TaxMsg(800, "The prepay adjustment cannot allow the prepay balance to be less than zero.")
    fpCurrPrepayAdjBal = fpCurrPrepayBal
    fpCurrPrepayAdjAmt = CDbl(fpCurrPrepayBal.Value)
    fpCurrPrepayAdjAmt.SetFocus
  End If

End Sub

Private Sub fpCurrPersAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrPersAdj.Value) > CDbl(fpCurrPersOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrPersAdj = CDbl(fpCurrPersOwed.Value)
      fpCurrPersAdj.SetFocus
    End If
  End If
  
  Call FigureAdjCol(1)
End Sub

Private Sub fpCurrOpt1Adj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpCurrOpt2Adj.Enabled = True Then
      fpCurrOpt2Adj.SetFocus
    ElseIf fpCurrOpt3Adj.Enabled = True Then
      fpCurrOpt3Adj.SetFocus
    Else
      fpLongAcctNum.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    fpCurrMHAdj.SetFocus
  End If
End Sub

Private Sub fpCurrOpt1Adj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrOpt1Adj.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrOpt1Adj = CDbl(fpCurrOpt1Owed.Value)
      fpCurrOpt1Adj.SetFocus
    End If
  End If
  Call FigureAdjCol(8)
End Sub

Private Sub fpCurrOpt2Adj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpCurrOpt3Adj.Enabled = True Then
      fpCurrOpt3Adj.SetFocus
    Else
      fptxtNote.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrOpt1Adj.Enabled = True Then
      fpCurrOpt1Adj.SetFocus
    Else
      fpCurrMHAdj.SetFocus
    End If
  End If

End Sub

Private Sub fpCurrOpt2Adj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrOpt2Adj.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrOpt2Adj = CDbl(fpCurrOpt2Owed.Value)
      fpCurrOpt2Adj.SetFocus
    End If
  End If
  Call FigureAdjCol(9)
End Sub

Private Sub fpCurrOpt3Adj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpLongAcctNum.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrOpt2Adj.Enabled = True Then
      fpCurrOpt2Adj.SetFocus
    ElseIf fpCurrOpt1Adj.Enabled = True Then
      fpCurrOpt1Adj.SetFocus
    Else
      fpCurrMHAdj.SetFocus
    End If
  End If
End Sub

Private Sub fpCurrOpt3Adj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrOpt3Adj.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrOpt3Adj = CDbl(fpCurrOpt3Owed.Value)
      fpCurrOpt3Adj.SetFocus
    End If
  End If
  Call FigureAdjCol(10)
End Sub

Private Sub fpCurrTotAdj_Change()
  Call UpdatePrepay
End Sub

Private Sub fpLngIntBill_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fpLongAcctNum.SetFocus
  ElseIf KeyCode = vbKeyDown Then
    fpcboAdjType.SetFocus
'    fpCurrPrincAdj.SetFocus
  End If
End Sub


Private Sub fpLongAcctNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtDate.SetFocus
'    fpLngIntBill.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrOpt3Adj.Enabled = True Then
      fpCurrOpt3Adj.SetFocus
    ElseIf fpCurrOpt2Adj.Enabled = True Then
      fpCurrOpt2Adj.SetFocus
    ElseIf fpCurrOpt1Adj.Enabled = True Then
      fpCurrOpt1Adj.SetFocus
    Else
      fpCurrMHAdj.SetFocus
    End If
  End If
End Sub

Private Sub fpLongAcctNum_LostFocus()
  Dim SaveAcctNum As Long
  Dim PrePay As Double
  
  On Error GoTo ERRORSTUFF
  
  If ExitOK = True Then Exit Sub
  If CLng(fpLongAcctNum.Value) = 0 Then Exit Sub
  SaveAcctNum = 0
  If TempAcctNum = CLng(fpLongAcctNum.Value) Then Exit Sub
  
  If Check4ValidCustNum(CLng(fpLongAcctNum.Value)) = True Then
    If fpCurrTotAdj.Value > 0 Then
      If TaxMsgWOpts(900, "You have made changes that are not saved. Press F10 to save these changes. Otherwise, press ESC to continue without saving.", "F10 Save", "ESC Don't Save") = "abort" Then
        Unload frmVATaxMsgWOpts
        SaveAcctNum = fpLongAcctNum.Value
        Call Clearscreen
        TempAcctNum = SaveAcctNum
        fpLongAcctNum.Value = SaveAcctNum
        Call LoadMeEdit
      Else
        Unload frmVATaxMsgWOpts
        Unload frmVATaxMsgWOpts
        SaveAcctNum = fpLongAcctNum.Value
        fpLongAcctNum = TempAcctNum
        SaveOK = False
        Call cmdPost_Click
        If SaveOK = True Then
          Call Clearscreen
          fpLongAcctNum.Value = SaveAcctNum
          TempAcctNum = fpLongAcctNum.Value
          Call LoadMeEdit
        End If
      End If
    Else 'adjustment is zero
      SaveAcctNum = fpLongAcctNum.Value
      Call Clearscreen
      TempAcctNum = SaveAcctNum
      fpLongAcctNum.Value = SaveAcctNum
      Call LoadMeEdit
    End If
    PrePay = GetCustPersBalance(CLng(fpLongAcctNum.Value), -1)
    If PrePay < 0 Then
      fpCurrPrepayBal.Text = Abs(PrePay)
      fpCurrPrepayAdjBal.Text = Abs(PrePay)
      fpCurrPrepayAdjAmt.Text = 0
      fpcboAdjType.AddItem "5-Prepay Adjust Down"
    Else
      fpCurrPrepayBal.Text = 0
      fpCurrPrepayAdjBal.Text = 0
      fpCurrPrepayAdjAmt.Text = 0
      If fpcboAdjType.ListCount = 5 Then
        fpcboAdjType.Clear
        fpcboAdjType.AddItem "1-Billing Downward Adjustment"
        fpcboAdjType.AddItem "2-Billing Upward Adjustment"
        fpcboAdjType.AddItem "3-Adjustment for Payment"
        fpcboAdjType.AddItem "4-Release"
      End If
    End If
  Else
    Call TaxMsg(900, "An invalid customer account number, " + fpLongAcctNum.Text + ", has been entered. Resetting account number to zero. Please try again.")
    fpLongAcctNum.Text = "0"
    Close
    fpCurrPrepayBal.Text = 0
    fpCurrPrepayAdjAmt.Text = 0
    fpCurrPrepayAdjBal.Text = 0
      If fpcboAdjType.ListCount = 5 Then
        fpcboAdjType.Clear
        fpcboAdjType.AddItem "1-Billing Downward Adjustment"
        fpcboAdjType.AddItem "2-Billing Upward Adjustment"
        fpcboAdjType.AddItem "3-Adjustment for Payment"
        fpcboAdjType.AddItem "4-Release"
      End If
    Exit Sub
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "fpLongAcctNum_LostFocus", Erl)
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


End Sub
Private Function Check4ValidCustNum(ThisCust As Long) As Boolean
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim Number$
  Dim Name$
  Dim Found As Boolean

  On Error GoTo ERRORSTUFF
  
  Check4ValidCustNum = True
  
  If fpLongAcctNum.Value = 0 Then
    Check4ValidCustNum = False
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no tax customers saved."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If ThisCust = TaxRec.Acct Then
      If TaxRec.Deleted <> 0 Then
        Check4ValidCustNum = False
      End If
      Exit For
    End If
  Next x

  Close CHandle

  If x > NumOfCRecs Then
    Call Clearscreen
    Check4ValidCustNum = False
  End If
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "Check4ValidCustNum", Erl)
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

Private Sub Clearscreen()
  fptxtNote.Text = ""
  Label10.Caption = "Balance"
  fpcboAdjType.Text = "1-Billing Downward Adjustment"
  TempAcctNum = 0
  fpLongAcctNum.Text = 0
  fptxtDate = Date
  fpLngIntBill.Text = 0
  fptxtName.Text = ""
  fptxtAddress.Text = ""
  fptxtCity.Text = ""
  fptxtState.Text = ""
  fptxtZip.Text = ""
  fpCurrPersOwed.Text = 0
  fpCurrPersAdj.Text = 0
  fpCurrMTOwed.Text = 0
  fpCurrMTAdj.Text = 0
  fpCurrMCOwed.Text = 0
  fpCurrMCAdj.Text = 0
  fpCurrFEOwed.Text = 0
  fpCurrFEAdj.Text = 0
  fpCurrMHOwed.Text = 0
  fpCurrMHAdj.Text = 0
  fpCurrIntOwed.Text = 0
  fpCurrIntAdj.Text = 0
  fpCurrPenOwed.Text = 0
  fpCurrPenAdj.Text = 0
  fpCurrOpt1Owed.Text = 0
  fpCurrOpt1Adj.Text = 0
  fpCurrOpt2Owed.Text = 0
  fpCurrOpt2Adj.Text = 0
  fpCurrOpt3Owed.Text = 0
  fpCurrOpt3Adj.Text = 0
  fpCurrTotOwed.Text = 0
  fpCurrTotAdj.Text = 0
  fpCurrPrepayBal.Text = 0
  fpCurrPrepayAdjBal.Text = 0
  fpCurrPrepayAdjAmt.Text = 0
  If fpcboAdjType.ListCount = 5 Then
    fpcboAdjType.Clear
    fpcboAdjType.AddItem "1-Billing Downward Adjustment"
    fpcboAdjType.AddItem "2-Billing Upward Adjustment"
    fpcboAdjType.AddItem "3-Adjustment for Payment"
    fpcboAdjType.AddItem "4-Release"
  End If
End Sub

Private Sub LoadTemps()
  TempAcctNum = fpLongAcctNum.Value
  TempBillNum = fpLngIntBill.Value
End Sub

Private Sub LoadMeBill()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim NextTrans As Long
  Dim BillNum As Long
  Dim TPers#, TInterest#, TPenalty#
  Dim TMT#, TMC#, TMH#
  Dim TFE#, TOpt1#, TOpt2#, TOpt3#
  Dim WhatsLeft#, ThisDif# 'added 8/9/2006
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxTransFile TRHandle, NumOfTRRecs
  Get TRHandle, BillRec, TaxRec
  Close TRHandle
  If TaxRec.DiscAmt > 0 Then Call ApplyDisc(TRHandle, TaxRec) 'added 1/25/07
  PersPaid = TaxRec.Revenue.Principle1Pd + TaxRec.PPTRADisc - TaxRec.PPTRARmvl '1/16/07 added PPTRADisc & PPTRARmvl
  MTPaid = TaxRec.Revenue.Principle2Pd
  MCPaid = TaxRec.Revenue.Principle3Pd
  FEPaid = TaxRec.Revenue.Principle4Pd
  MHPaid = TaxRec.Revenue.Principle5Pd
  PIntPaid = TaxRec.Revenue.InterestPd
  PenPaid = TaxRec.Revenue.PenaltyPd
  Rev1Paid = TaxRec.Revenue.RevOpt1Pd
  Rev2Paid = TaxRec.Revenue.RevOpt2Pd
  Rev3Paid = TaxRec.Revenue.RevOpt3Pd
  TotPaid = OldRound(PersPaid + MTPaid + MCPaid + FEPaid + MHPaid + PIntPaid + PenPaid + Rev1Paid + Rev2Paid + Rev3Paid)
  BillNum& = CLng(ParseBillNum(TaxRec.Description))
  If BillNum = CLng(fpLngIntBill.Value) Then
    TPers# = TaxRec.Revenue.Principle1
    TPers# = OldRound(TPers# - (TaxRec.Revenue.Principle1Pd + TaxRec.PPTRADisc - TaxRec.PPTRARmvl)) '1/16/07 added PPTRADisc and PPTRARmvl'
    PersBal = TPers#
    TMT# = TaxRec.Revenue.Principle2
    TMT# = OldRound(TMT# - TaxRec.Revenue.Principle2Pd)
    MTBal = TMT#
    TMC# = TaxRec.Revenue.Principle3
    TMC# = OldRound(TMC# - TaxRec.Revenue.Principle3Pd) '8/9/2006 changed to TMC from TMT
    MCBal = TMC#
    TFE# = TaxRec.Revenue.Principle4
    TFE# = OldRound(TFE# - TaxRec.Revenue.Principle4Pd)
    FEBal = TFE#
    TMH# = TaxRec.Revenue.Principle5
    TMH# = OldRound(TMH# - TaxRec.Revenue.Principle5Pd)
    MHBal = TMH#
    
    
    TInterest# = OldRound#(TaxRec.Revenue.Interest - TaxRec.Revenue.InterestPd)
    PIntBal = TInterest#
    TPenalty# = OldRound#(TaxRec.Revenue.Penalty - TaxRec.Revenue.PenaltyPd)
    PenBal = TPenalty#
    
    TOpt1# = OldRound(TaxRec.Revenue.RevOpt1 - TaxRec.Revenue.RevOpt1Pd)
    Rev1Bal = TOpt1#
    TOpt2# = OldRound(TaxRec.Revenue.RevOpt2 - TaxRec.Revenue.RevOpt2Pd)
    Rev2Bal = TOpt2#
    TOpt3# = OldRound(TaxRec.Revenue.RevOpt3 - TaxRec.Revenue.RevOpt3Pd)
    Rev3Bal = TOpt3#
    
    If TaxRec.DiscAmt > 0 Then GoSub ApplyDisc 'added 8/9/2006
    
    fpCurrPersOwed.Text = TPers#
    fpCurrPersAdj.Text = 0
    fpCurrMTOwed.Text = TMT#
    fpCurrMTAdj.Text = 0
    fpCurrMCOwed.Text = TMC#
    fpCurrMCAdj.Text = 0
    fpCurrFEOwed.Text = TFE#
    fpCurrFEAdj.Text = 0
    fpCurrMHOwed.Text = TMH#
    fpCurrMHAdj.Text = 0
    fpCurrIntOwed.Text = TInterest#
    fpCurrIntAdj.Text = 0
    fpCurrPenOwed.Text = TPenalty#
    fpCurrPenAdj.Text = 0
    fpCurrOpt1Owed.Text = TOpt1#
    fpCurrOpt1Adj.Text = 0
    fpCurrOpt2Owed.Text = TOpt2#
    fpCurrOpt2Adj.Text = 0
    fpCurrOpt3Owed.Text = TOpt3#
    fpCurrOpt3Adj.Text = 0
    fpCurrTotOwed.Text = OldRound(TPers# + TMT# + TMC# + TFE# + TMH# + TInterest# + TPenalty# + TOpt1# + TOpt2# + TOpt3#)
    fpCurrTotAdj.Text = 0
  Else
    Call TaxMsg(800, "The bill number entered could not be found in this customer's transaction records. Please enter another bill number or press F8 to bring up a complete list of all the bills for this customer.")
  End If
  
  Exit Sub
  
ApplyDisc: '8/9/2006
  WhatsLeft = TaxRec.DiscAmt
  ThisDif = OldRound(TaxRec.Revenue.Principle5 - TaxRec.Revenue.Principle5Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      MHBal# = OldRound(MHBal# - ThisDif)
      TMH# = MHBal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      MHBal# = OldRound(MHBal# - WhatsLeft#)
      TMH# = MHBal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.Principle4 - TaxRec.Revenue.Principle4Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      FEBal# = OldRound(FEBal# - ThisDif)
      TFE# = FEBal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      FEBal# = OldRound(FEBal# - WhatsLeft#)
      TFE# = FEBal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.Principle3 - TaxRec.Revenue.Principle3Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      MCBal# = OldRound(MCBal# - ThisDif)
      TMC# = MCBal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      MCBal# = OldRound(MCBal# - WhatsLeft#)
      TMC# = MCBal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.Principle2 - TaxRec.Revenue.Principle2Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      MTBal# = OldRound(MTBal# - ThisDif)
      TMT# = MTBal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      MTBal# = OldRound(MTBal# - WhatsLeft#)
      TMT# = MTBal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.Principle1 - TaxRec.Revenue.Principle1Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      PersBal# = OldRound(PersBal# - ThisDif)
      TPers# = PersBal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      PersBal# = OldRound(PersBal# - WhatsLeft#)
      TPers# = PersBal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.RevOpt1 - TaxRec.Revenue.RevOpt1Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      Rev1Bal# = OldRound(Rev1Bal# - ThisDif)
      TOpt1# = Rev1Bal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      Rev1Bal# = OldRound(Rev1Bal# - WhatsLeft#)
      TOpt1# = Rev1Bal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.RevOpt2 - TaxRec.Revenue.RevOpt2Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      Rev2Bal# = OldRound(Rev2Bal# - ThisDif)
      TOpt2# = Rev2Bal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      Rev2Bal# = OldRound(Rev2Bal# - WhatsLeft#)
      TOpt2# = Rev2Bal#
      WhatsLeft = 0
      Return
    End If
  End If
  ThisDif = OldRound(TaxRec.Revenue.RevOpt3 - TaxRec.Revenue.RevOpt3Pd)
  If ThisDif > 0 Then
    If WhatsLeft >= ThisDif Then
      Rev3Bal# = OldRound(Rev3Bal# - ThisDif)
      TOpt3# = Rev3Bal#
      WhatsLeft = OldRound(WhatsLeft - ThisDif)
    Else
      Rev3Bal# = OldRound(Rev3Bal# - WhatsLeft#)
      TOpt3# = Rev3Bal#
      WhatsLeft = 0
      Return
    End If
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "LoadMeBill", Erl)
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

End Sub

Private Sub FigureAdjCol(Rev As Integer)
  Dim TotAdj As Double
  Dim BalDif As Double
  
  TotAdj = 0
  TotAdj = OldRound(CDbl(fpCurrPersAdj.Value) + CDbl(fpCurrMTAdj.Value) + CDbl(fpCurrMCAdj.Value) + CDbl(fpCurrFEAdj.Value))
  TotAdj = OldRound(TotAdj + CDbl(fpCurrMHAdj.Value) + CDbl(fpCurrIntAdj.Value) + CDbl(fpCurrPenAdj.Value))
  TotAdj = OldRound(TotAdj + CDbl(fpCurrOpt1Adj.Value) + CDbl(fpCurrOpt2Adj.Value) + CDbl(fpCurrOpt3Adj.Value))
  fpCurrTotAdj.Text = TotAdj
  
  fpCurrPrepayAdjBal = OldRound(CDbl(fpCurrPrepayBal.Value) - TotAdj)
  If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then
    fpCurrPrepayAdjBal = 0
  End If
End Sub

Private Sub fptxtDate_LostFocus()
  Dim ThisDate As Integer
  Dim ThatDate As Integer
  ThisDate = Date2Num(Date$)
  ThatDate = Date2Num(fptxtDate.Text)
  If Abs(ThatDate - ThisDate) > 60 Then
    If TaxMsgWOpts(800, "The date entered is more than 60 days from today's date. Press F10 if you want to continue with this date. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtDate.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("ERROR: User warned that the date entered is more than 60 days from today's date and elected to continue with the date entered anyway.")
    End If
  End If
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisType$
  Dim CustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim LastTrans As Long
  Dim ThisPrinc As Double
  Dim ThisPrincPd As Double
  Dim FF$, Line$
  Dim ThisAmt As Double
  
  On Error GoTo ERRORSTUFF
  
  FF$ = Chr$(12)
  Line$ = String(80, "-")
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, CustRec
  Close
  LastTrans = CustRec.LastTrans
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, LastTrans, TaxTrans
  If fpcboAdjType <> "5-Prepay Adjust Down" Then
    Get TTHandle, TaxTrans.BelongTo, TaxTrans
  End If
  Close TTHandle
  
  ReportFile$ = "TXADJRPT.PRN"
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  Print #RptHandle, Tab(30); "Tax Adjustment Report"
  Print #RptHandle, Now
  If InStr(fpcboAdjType.Text, "5") Then
    Print #RptHandle, QPTrim$(TaxMasterRec.Name); Tab(50); "Transaction Amount: " + QPTrim$(Using$("$##,##0.00", CDbl(fpCurrPrepayAdjBal.Value)))
  Else
    Print #RptHandle, QPTrim$(TaxMasterRec.Name); Tab(50); "Transaction Amount: " + QPTrim$(Using$("$##,##0.00", CDbl(fpCurrTotAdj.Value)))
  End If
  Print #RptHandle, Line
  Print #RptHandle,
  Print #RptHandle, Tab(14); "Customer Name:  " + QPTrim$(fptxtName.Text)
  Print #RptHandle, Tab(18); "Account #:  " + QPTrim$(Using$("########0", GCustNum))
  Print #RptHandle, Tab(20); "Address:  " + QPTrim$(fptxtAddress.Text)
  If fpcboAdjType <> "5-Prepay Adjust Down" Then
    Print #RptHandle, Tab(16); "Bill Number:  " + QPTrim$(Using$("########0", ThisBillNum))
  Else
    Print #RptHandle, Tab(16); "Bill Number:  NA"
  End If
  Print #RptHandle,
  If InStr(fpcboAdjType.Text, "1") Or InStr(fpcboAdjType.Text, "4") Then
    If InStr(fpcboAdjType.Text, "4") Then
      Print #RptHandle, Tab(5); "Adjustment Type:  Release"
    Else
      Print #RptHandle, Tab(5); "Adjustment Type:  Billing Adjustment Down"
    End If
    Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
    Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
    Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
    Print #RptHandle, Tab(5); QPTrim$(fptxtPers.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPersAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrPersOwed.Value) - CDbl(fpCurrPersAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtMachTools.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMTAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrMTOwed.Value) - CDbl(fpCurrMTAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtMerchCap.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMCAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrMCOwed.Value) - CDbl(fpCurrMCAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtFarmEquip.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrFEAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrFEOwed.Value) - CDbl(fpCurrFEAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtMobHomes.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMHAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrMHOwed.Value) - CDbl(fpCurrMHAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtInterest.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrIntAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrIntOwed.Value) - CDbl(fpCurrIntAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPenAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrPenOwed.Value) - CDbl(fpCurrPenAdj.Value)))
    If QPTrim$(fpTxtOpt1.Text) <> "" Then
      Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt1.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt1Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrOpt1Owed.Value) - CDbl(fpCurrOpt1Adj.Value)))
    End If
    If QPTrim$(fpTxtOpt2.Text) <> "" Then
      Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt2.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt2Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrOpt2Owed.Value) - CDbl(fpCurrOpt2Adj.Value)))
    End If
    If QPTrim$(fpTxtOpt3.Text) <> "" Then
      Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt3.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt3Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrOpt3Owed.Value) - CDbl(fpCurrOpt3Adj.Value)))
    End If
  Else
    If InStr(fpcboAdjType.Text, "2") Then
      Print #RptHandle, Tab(5); "Adjustment Type:  Billing Adjustment Up"
      Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
      Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
      Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
      Print #RptHandle, Tab(5); QPTrim$(fptxtPers.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPersAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(PersVal + PersBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtMachTools.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMTAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrMTOwed.Value) + CDbl(fpCurrMTAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtMerchCap.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMCAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrMCOwed.Value) + CDbl(fpCurrMCAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtFarmEquip.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrFEAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrFEOwed.Value) + CDbl(fpCurrFEAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtMobHomes.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMHAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrMHOwed.Value) + CDbl(fpCurrMHAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtInterest.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrIntAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrIntOwed.Value) + CDbl(fpCurrIntAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPenAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrPenOwed.Value) + CDbl(fpCurrPenAdj.Value)))
      If QPTrim$(fpTxtOpt1.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt1.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt1Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrOpt1Owed.Value) + CDbl(fpCurrOpt1Adj.Value)))
      End If
      If QPTrim$(fpTxtOpt2.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt2.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt2Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrOpt2Owed.Value) + CDbl(fpCurrOpt2Adj.Value)))
      End If
      If QPTrim$(fpTxtOpt3.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt3.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt3Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrOpt3Owed.Value) + CDbl(fpCurrOpt3Adj.Value)))
      End If
    ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then '#3
      Print #RptHandle, Tab(5); "Adjustment Type:  Payment Adjustment"
      Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
      Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
      Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
      Print #RptHandle, Tab(5); QPTrim$(fptxtPers.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPersAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(PersVal + PersBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtMachTools.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMTAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(MTVal + MTBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtMerchCap.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMCAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(MCVal + MCBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtFarmEquip.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrFEAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(FEVal + FEBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtMobHomes.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrMHAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(MHVal + MHBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtInterest.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrIntAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(PIntVal + PIntBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPenAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(PenVal + PenBal))
      If QPTrim$(fpTxtOpt1.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt1.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt1Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(Rev1Val + Rev1Bal))
      End If
      If QPTrim$(fpTxtOpt2.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt2.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt2Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(Rev2Val + Rev2Bal))
      End If
      If QPTrim$(fpTxtOpt3.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt3.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrOpt3Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(Rev3Val + Rev3Bal))
      End If
    ElseIf fpcboAdjType.Text = "5-Prepay Adjust Down" Then '
      Print #RptHandle, Tab(5); "Adjustment Type:  Prepay Adjustment Down"
      Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
      Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
      Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
      Print #RptHandle, Tab(5); QPTrim$(fptxtPers.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtMachTools.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtMerchCap.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtFarmEquip.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtMobHomes.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtInterest.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      If QPTrim$(fpTxtOpt1.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt1.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      End If
      If QPTrim$(fpTxtOpt2.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt2.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      End If
      If QPTrim$(fpTxtOpt3.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fpTxtOpt3.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      End If
    End If
  End If
  Print #RptHandle, Tab(5); "----Account Balance Information----"
  Print #RptHandle,
  If InStr(fpcboAdjType.Text, "3") Then
    Print #RptHandle, Tab(5); "Balance Excluding This Bill: "; Tab(39); Using$("$###,##0.00", OldRound(GetCustPersBalance(GCustNum, -1) - NewBalThisBill))
    Print #RptHandle, Tab(5); "Previous Balance This Bill: "; Tab(39); Using("$###,##0.00", ThisBillBal#)
    If CDbl(fpCurrPrepayBal.Value) > 0 Then
      Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
      Print #RptHandle, Tab(5); "Prepaid Used: "; Tab(39); Using("$###,##0.00", Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))))
    Else
      Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
    End If
    Print #RptHandle, Tab(5); "Balance For This Bill: "; Tab(39); Using("$###,##0.00", NewBalThisBill)
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustPersBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  ElseIf InStr(fpcboAdjType.Text, "1") Or InStr(fpcboAdjType.Text, "4") Then
    Print #RptHandle, Tab(5); "Balance Excluding This Bill: "; Tab(39); Using$("$###,##0.00", OldRound(ThisGBal - CDbl(fpCurrTotOwed.Value)))
    Print #RptHandle, Tab(5); "Previous Balance This Bill: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotOwed.Value))
    Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
    Print #RptHandle, Tab(5); "Balance For This Bill: "; Tab(39); Using("$###,##0.00", OldRound(CDbl(fpCurrTotOwed.Value) - CDbl(fpCurrTotAdj.Value)))
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustPersBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  ElseIf InStr(fpcboAdjType.Text, "2") Then
    Print #RptHandle, Tab(5); "Balance Excluding This Bill: "; Tab(39); Using$("$###,##0.00", OldRound(GetCustPersBalance(GCustNum, -1) - NewBalThisBill))
    Print #RptHandle, Tab(5); "Previous Balance This Bill: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotOwed.Value))
    Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
    If CDbl(fpCurrPrepayBal.Value) > 0 Then
      Print #RptHandle, Tab(5); "Prepaid Used: "; Tab(39); Using("$###,##0.00", Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))))
    End If
    Print #RptHandle, Tab(5); "Balance For This Bill: "; Tab(39); Using("$###,##0.00", NewBalThisBill)
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustPersBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  ElseIf InStr(fpcboAdjType.Text, "5") Then
    Print #RptHandle, Tab(5); "Previous Balance/Prepaid Balance: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrPrepayBal.Value))
    Print #RptHandle, Tab(5); "Prepaid Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrPrepayAdjAmt.Value))
    Print #RptHandle, Tab(5); "Prepaid Balance: "; Tab(39); Using("$###,##0.00", OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjAmt.Value)))
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustPersBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  End If
  Close
  
  ViewPrint ReportFile$, "Tax Customer Export"
  Kill ReportFile$
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "PrintText", Erl)
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


End Sub

Private Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdMessage.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdMessage.ForeColor = &H80000012
      cmdMessage.FontSize = BtnFnt
    Case 2
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 0.7
    Case 3
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 1.4
    Case 4
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.1
    Case 5
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.8
    Case 6
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 3.5
    Case 7
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 4.2
    Case 8
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 4.9
    Case 9
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If

End Sub

Private Function Check4PaidAmts(BillRec As Double) As Boolean
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim PersPaid As Double
  Dim MTPaid As Double
  Dim MCPaid As Double
  Dim FEPaid As Double
  Dim OptRev1Paid As Double
  Dim OptRev2Paid As Double
  Dim OptRev3Paid As Double
  
  On Error GoTo ERRORSTUFF
  
  Check4PaidAmts = True
  
  If BillRec = 0 Then Exit Function
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, BillRec, TaxTrans
  Close TTHandle
  If TaxTrans.DiscAmt > 0 Then Call ApplyDisc(TTHandle, TaxTrans)  'added 1/25/07
  
  PersPaid = OldRound(TaxTrans.Revenue.Principle1Pd) ' + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
  If CDbl(fpCurrPersAdj.Value) > PersPaid Then '11/21/2005 error here
    Call TaxMsg(800, "The total amount paid for PERSONAL is " + QPTrim$(Using$("$###,##0.00", PersPaid)) + ". Adjustments for payments cannot exceed the total amount paid.")
    fpCurrPersAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  MTPaid = TaxTrans.Revenue.Principle2Pd
  If CDbl(fpCurrMTAdj.Value) > MTPaid Then
    Call TaxMsg(800, "The total amount already paid for MACHINE TOOLS is " + QPTrim$(Using$("$###,##0.00", MTPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrMTAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  MCPaid = TaxTrans.Revenue.Principle3Pd
  If CDbl(fpCurrMCAdj.Value) > MCPaid Then
    Call TaxMsg(800, "The total amount already paid for MERCHANT CAP is " + QPTrim$(Using$("$###,##0.00", MCPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrMCAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  FEPaid = TaxTrans.Revenue.Principle4Pd
  If CDbl(fpCurrFEAdj.Value) > FEPaid Then
    Call TaxMsg(800, "The total amount already paid for FARM EQUIPMENT is " + QPTrim$(Using$("$###,##0.00", FEPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrFEAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  MHPaid = TaxTrans.Revenue.Principle5Pd
  If CDbl(fpCurrMHAdj.Value) > MHPaid Then
    Call TaxMsg(800, "The total amount already paid for MOBILE HOMES is " + QPTrim$(Using$("$###,##0.00", MHPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrMHAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  PIntPaid = TaxTrans.Revenue.InterestPd
  If CDbl(fpCurrIntAdj.Value) > PIntPaid Then '10/6/09 changed from MHPaid in error
    Call TaxMsg(800, "The total amount already paid for INTEREST is " + QPTrim$(Using$("$###,##0.00", PIntPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrIntAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  PenPaid = TaxTrans.Revenue.PenaltyPd
  If CDbl(fpCurrPenAdj.Value) > PenPaid Then
    Call TaxMsg(800, "The total amount already paid for PENALTY is " + QPTrim$(Using$("$###,##0.00", PenPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrPenAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  If fpCurrOpt1Adj.Enabled = True Then
    OptRev1Paid = TaxTrans.Revenue.RevOpt1Pd
    If CDbl(fpCurrOpt1Adj.Value) > OptRev1Paid Then
      Call TaxMsg(800, "The total amount already paid for " + QPTrim$(fpTxtOpt1.Text) + " is " + QPTrim$(Using$("$###,##0.00", OptRev1Paid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
      fpCurrOpt1Adj.SetFocus
      Close
      Check4PaidAmts = False
      Exit Function
    End If
  End If
    
  If fpCurrOpt2Adj.Enabled = True Then
    OptRev2Paid = TaxTrans.Revenue.RevOpt2Pd
    If CDbl(fpCurrOpt2Adj.Value) > OptRev2Paid Then
      Call TaxMsg(800, "The total amount already paid for " + QPTrim$(fpTxtOpt2.Text) + " is " + QPTrim$(Using$("$###,##0.00", OptRev2Paid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
      fpCurrOpt2Adj.SetFocus
      Close
      Check4PaidAmts = False
      Exit Function
    End If
  End If
    
  If fpCurrOpt3Adj.Enabled = True Then
    OptRev3Paid = TaxTrans.Revenue.RevOpt3Pd
    If CDbl(fpCurrOpt3Adj.Value) > OptRev3Paid Then
      Call TaxMsg(800, "The total amount already paid for " + QPTrim$(fpTxtOpt3.Text) + " is " + QPTrim$(Using$("$###,##0.00", OptRev3Paid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
      fpCurrOpt3Adj.SetFocus
      Close
      Check4PaidAmts = False
      Exit Function
    End If
  End If

  On Error GoTo ERRORSTUFF
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "Check4PaidAmts", Erl)
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

Private Sub UpdatePrepay()
  Dim x As Integer
  
  If CDbl(fpCurrPrepayAdjBal.Value) <= 0 Then Exit Sub
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    fpCurrPrepayAdjBal = OldRound(CDbl(fpCurrPrepayBal.Value) + CDbl(fpCurrPrepayAdjAmt.Value))
    If fpCurrPrepayAdjBal < 0 Then
      Call TaxMsg(800, "The prepay adjustment cannot cause the prepay balance to be less than zero.")
      fpCurrPrepayAdjBal = fpCurrPrepayBal
      fpCurrPrepayAdjAmt = fpCurrPrepayBal
    End If
    Exit Sub
  End If
  If fpcboAdjType.Text = "2-Billing Upward Adjustment" Then
    For x = 1 To 10
      If x = PayOrder(1) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPenAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPenAdj)
        End If
      ElseIf x = PayOrder(2) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrMTAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrMTAdj)
        End If
      ElseIf x = PayOrder(3) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrMCAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrMCAdj)
        End If
      ElseIf x = PayOrder(4) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrFEAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrFEAdj)
        End If
      ElseIf x = PayOrder(5) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrMHAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrMHAdj)
        End If
      ElseIf x = PayOrder(6) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrIntAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrIntAdj)
        End If
      ElseIf x = PayOrder(7) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPenAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPenAdj)
        End If
      ElseIf x = PayOrder(8) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrOpt1Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrOpt1Adj)
        End If
      ElseIf x = PayOrder(9) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrOpt2Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrOpt2Adj)
        End If
      ElseIf x = PayOrder(10) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrOpt3Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
          GoTo DoneHere
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrOpt3Adj)
        End If
      End If
    Next x
DoneHere:
  ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then
    For x = 1 To 10
      If x = PayOrder(1) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPersAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPersAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(2) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrMTAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrMTAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(3) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrMCAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrMCAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(4) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrFEAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrFEAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(5) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrMHAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrMHAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(6) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrIntAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrIntAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(7) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPenAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPenAdj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(8) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrOpt1Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrOpt1Adj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(9) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrOpt2Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrOpt2Adj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      ElseIf x = PayOrder(10) Then
        If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrOpt3Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
          fpCurrPrepayAdjBal = 0
        Else
          fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrOpt3Adj)
          If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
        End If
      End If
    Next x
  ElseIf fpcboAdjType.Text = "4-Release" Then
    fpCurrPrepayAdjBal = fpCurrTotAdj
  End If
 
End Sub

Private Sub ClearUserInput()
  fpCurrPrepayAdjBal = CDbl(fpCurrPrepayBal.Value)
  fpCurrPrepayAdjAmt = 0
  fpCurrPersAdj = 0
  fpCurrMTAdj = 0
  fpCurrMCAdj = 0
  fpCurrFEAdj = 0
  fpCurrMHAdj = 0
  fpCurrIntAdj = 0
  fpCurrPenAdj = 0
  fpCurrOpt1Adj = 0
  fpCurrOpt2Adj = 0
  fpCurrOpt3Adj = 0
  fpCurrTotAdj = 0
End Sub
Private Sub PrintGraphics()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisType$
  Dim CustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim dlm$
  Dim LastTrans As Long
  Dim ThisBal As Double
  Dim ThisAmt As Double
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, CustRec
  Close
  LastTrans = CustRec.LastTrans
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, LastTrans, TaxTrans
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then GoTo PrePayOnly
  Get TTHandle, TaxTrans.BelongTo, TaxTrans
PrePayOnly:
  Close TTHandle
  ThisBal = GetCustPersBalance(GCustNum, -1)
  
  ReportFile$ = "TAXRPTS\TXPADJRPT.RPT"
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  '                              0
  Print #RptHandle, QPTrim$(TaxMasterRec.Name); dlm;
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    '                       1                         2
    Print #RptHandle, CDbl(fpCurrPrepayAdjAmt.Value); dlm; QPTrim$(fptxtName.Text); dlm;
  Else
    '                       1                         2
    Print #RptHandle, CDbl(fpCurrTotAdj.Value); dlm; QPTrim$(fptxtName.Text); dlm;
  End If
  '                    3                       4
  Print #RptHandle, GCustNum; dlm; QPTrim$(fptxtAddress.Text); dlm;
  '                             5                               6
  Print #RptHandle, QPTrim$(fptxtPers.Text); dlm; QPTrim$(fptxtInterest.Text); dlm;
  '                               7                                8
  Print #RptHandle, QPTrim$(fptxtMachTools.Text); dlm; QPTrim$(fptxtMerchCap.Text); dlm;
  '
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    If QPTrim$(fpTxtOpt1.Text) = "NOT IN USE" Then
      '                  9
      Print #RptHandle, ""; dlm;
    Else
      '                             9
      Print #RptHandle, QPTrim$(fpTxtOpt1.Text); dlm;
    End If
    If QPTrim$(fpTxtOpt2.Text) = "NOT IN USE" Then
      '                 10
      Print #RptHandle, ""; dlm;
    Else
      '                               10
      Print #RptHandle, QPTrim$(fpTxtOpt2.Text); dlm;
    End If
    If QPTrim$(fpTxtOpt3.Text) = "NOT IN USE" Then
      '                 11
      Print #RptHandle, ""; dlm;
    Else
      '                              11
      Print #RptHandle, QPTrim$(fpTxtOpt3.Text); dlm;
    End If
  Else
    If fpCurrOpt1Adj.Enabled = True Then
      '                                 9
      Print #RptHandle, QPTrim$(fpTxtOpt1.Text); dlm;
    Else
      '                 9
      Print #RptHandle, ""; dlm;
    End If
    If fpCurrOpt2Adj.Enabled = True Then
      '                           10
      Print #RptHandle, QPTrim$(fpTxtOpt2.Text); dlm;
    Else
      '                 10
      Print #RptHandle, ""; dlm;
    End If
    If fpCurrOpt3Adj.Enabled = True Then
      '                               11
      Print #RptHandle, QPTrim$(fpTxtOpt3.Text); dlm;
    Else
      '                 11
      Print #RptHandle, ""; dlm;
    End If
  End If
  If InStr(fpcboAdjType.Text, "1") Or InStr(fpcboAdjType.Text, "4") Then
    '                              12                             13
    Print #RptHandle, CDbl(fpCurrPersAdj.Value); dlm; CDbl(fpCurrIntAdj.Value); dlm;
    '                              14                                 15
    Print #RptHandle, CDbl(fpCurrMTAdj.Value); dlm; CDbl(fpCurrMCAdj.Value); dlm;
    '                              16                                 17
    Print #RptHandle, CDbl(fpCurrOpt1Adj.Value); dlm; CDbl(fpCurrOpt2Adj.Value); dlm;
    '                              18                               19                            20                      21
    Print #RptHandle, CDbl(fpCurrOpt3Adj.Value); dlm; CDbl(fpCurrTotOwed.Value); dlm; CDbl(fpCurrTotAdj.Value); dlm; ThisBal; dlm;
    If InStr(fpcboAdjType.Text, "1") Then
      '                     22
      Print #RptHandle, "Billing Downward Adjustment"; dlm;
    Else
      '                     22
      Print #RptHandle, "Release"; dlm;
    End If
    '                                          23
    Print #RptHandle, OldRound(CDbl(fpCurrPersOwed.Value) - CDbl(fpCurrPersAdj.Value)); dlm;
    '                                          24
    Print #RptHandle, OldRound(CDbl(fpCurrIntOwed.Value) - CDbl(fpCurrIntAdj.Value)); dlm;
    '                                          25
    Print #RptHandle, OldRound(CDbl(fpCurrMTOwed.Value) - CDbl(fpCurrMTAdj.Value)); dlm;
    '                                          26
    Print #RptHandle, OldRound(CDbl(fpCurrMCOwed.Value) - CDbl(fpCurrMCAdj.Value)); dlm;
    '                                          27
    Print #RptHandle, OldRound(CDbl(fpCurrOpt1Owed.Value) - CDbl(fpCurrOpt1Adj.Value)); dlm;
    '                                          28
    Print #RptHandle, OldRound(CDbl(fpCurrOpt2Owed.Value) - CDbl(fpCurrOpt2Adj.Value)); dlm;
    '                                          29
    Print #RptHandle, OldRound(CDbl(fpCurrOpt3Owed.Value) - CDbl(fpCurrOpt3Adj.Value)); dlm;
    '                                          30
    Print #RptHandle, OldRound(ThisGBal - CDbl(fpCurrTotOwed.Value)); dlm;
    '                       31                                            32                           33
    Print #RptHandle, fpLngIntBill.Text; dlm; OldRound(CDbl(fpCurrTotOwed) - CDbl(fpCurrTotAdj)); dlm; 0; dlm;
    '                       34                        35                        36
    Print #RptHandle, fptxtFarmEquip.Text; dlm; fptxtMobHomes.Text; dlm; fptxtPenalty.Text; dlm;
    '                                             37
    Print #RptHandle, OldRound(CDbl(fpCurrFEOwed.Value) - CDbl(fpCurrFEAdj.Value)); dlm;
    '                                             38
    Print #RptHandle, OldRound(CDbl(fpCurrMHOwed.Value) - CDbl(fpCurrMHAdj.Value)); dlm;
    '                                             39
    Print #RptHandle, OldRound(CDbl(fpCurrPenOwed.Value) - CDbl(fpCurrPenAdj.Value)); dlm;
    '                          40                               41                          42                        43
    Print #RptHandle, CDbl(fpCurrFEAdj.Value); dlm; CDbl(fpCurrMHAdj.Value); dlm; CDbl(fpCurrPenAdj.Value); dlm; fptxtNote.Text
    
  Else
    If InStr(fpcboAdjType.Text, "2") Then
      '                              12                             13
      Print #RptHandle, CDbl(fpCurrPersAdj.Value); dlm; CDbl(fpCurrIntAdj.Value); dlm;
      '                              14                                 15
      Print #RptHandle, CDbl(fpCurrMTAdj.Value); dlm; CDbl(fpCurrMCAdj.Value); dlm;
      '                              16                                 17
      Print #RptHandle, CDbl(fpCurrOpt1Adj.Value); dlm; CDbl(fpCurrOpt2Adj.Value); dlm;
      '                              18                               19                            20                      21
      Print #RptHandle, CDbl(fpCurrOpt3Adj.Value); dlm; CDbl(fpCurrTotOwed.Value); dlm; CDbl(fpCurrTotAdj.Value); dlm; ThisBal; dlm;
      '                             22
      Print #RptHandle, "Billing Upward Adjustment"; dlm;
      '                                23
      Print #RptHandle, OldRound(PersBal + PersVal); dlm;
      '                                24
      Print #RptHandle, OldRound(PIntBal + PIntVal); dlm;
      '                                25
      Print #RptHandle, OldRound(MTBal + MTVal); dlm;
      '                                26
      Print #RptHandle, OldRound(MCBal + MCVal); dlm;
      '                                27
      Print #RptHandle, OldRound(Rev1Bal + Rev1Val); dlm;
      '                                28
      Print #RptHandle, OldRound(Rev2Bal + Rev2Val); dlm;
      '                                29
      Print #RptHandle, OldRound(Rev3Bal + Rev3Val); dlm;
      '                                        30
      Print #RptHandle, OldRound(GetCustPersBalance(GCustNum, -1) - NewBalThisBill); dlm;
      '                       31                      32                                              33
      Print #RptHandle, fpLngIntBill.Text; dlm; NewBalThisBill; dlm; Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))); dlm;
      '                       34                        35                        36
      Print #RptHandle, fptxtFarmEquip.Text; dlm; fptxtMobHomes.Text; dlm; fptxtPenalty.Text; dlm;
      '                                             37
      Print #RptHandle, OldRound(CDbl(fpCurrFEOwed.Value) + CDbl(fpCurrFEAdj.Value)); dlm;
      '                                             38
      Print #RptHandle, OldRound(CDbl(fpCurrMHOwed.Value) + CDbl(fpCurrMHAdj.Value)); dlm;
      '                                             39
      Print #RptHandle, OldRound(CDbl(fpCurrPenOwed.Value) + CDbl(fpCurrPenAdj.Value)); dlm;
      '                          40                               41                          42                         43
      Print #RptHandle, CDbl(fpCurrFEAdj.Value); dlm; CDbl(fpCurrMHAdj.Value); dlm; CDbl(fpCurrPenAdj.Value); dlm; fptxtNote.Text
    
    ElseIf InStr(fpcboAdjType.Text, "3") Then
      '                              12                             13
      Print #RptHandle, CDbl(fpCurrPersAdj.Value); dlm; CDbl(fpCurrIntAdj.Value); dlm;
      '                              14                                 15
      Print #RptHandle, CDbl(fpCurrMTAdj.Value); dlm; CDbl(fpCurrMCAdj.Value); dlm;
      '                              16                                 17
      Print #RptHandle, CDbl(fpCurrOpt1Adj.Value); dlm; CDbl(fpCurrOpt2Adj.Value); dlm;
      '                              18                         19                       20                   21
      Print #RptHandle, CDbl(fpCurrOpt3Adj.Value); dlm; ThisBillBal; dlm; CDbl(fpCurrTotAdj.Value); dlm; ThisBal; dlm;
      '                         22
      Print #RptHandle, "Payment Adjustment"; dlm;
      '                                23
      Print #RptHandle, OldRound(PersVal + PersBal); dlm;
      '                                24
      Print #RptHandle, OldRound(PIntVal + PIntBal); dlm;
      '                                25
      Print #RptHandle, OldRound(MTVal + MTBal); dlm;
      '                                26
      Print #RptHandle, OldRound(MCVal + MCBal); dlm;
      '                                27
      Print #RptHandle, OldRound(Rev1Val + Rev1Bal); dlm;
      '                                28
      Print #RptHandle, OldRound(Rev2Val + Rev2Bal); dlm;
      '                                29
      Print #RptHandle, OldRound(Rev3Val + Rev3Bal); dlm;
      '                                            30
      Print #RptHandle, OldRound(GetCustPersBalance(GCustNum, -1) - NewBalThisBill); dlm;
      '                       31                      32                                               33
      Print #RptHandle, fpLngIntBill.Text; dlm; NewBalThisBill; dlm; Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))); dlm;
      '                       34                        35                        36
      Print #RptHandle, fptxtFarmEquip.Text; dlm; fptxtMobHomes.Text; dlm; fptxtPenalty.Text; dlm;
'      '                                             37
'      Print #RptHandle, OldRound(CDbl(fpCurrFEOwed.Value) + CDbl(fpCurrFEAdj.Value)); dlm;
'      '                                             38
'      Print #RptHandle, OldRound(CDbl(fpCurrMHOwed.Value) + CDbl(fpCurrMHAdj.Value)); dlm;
'      '                                             39
'      Print #RptHandle, OldRound(CDbl(fpCurrPenOwed.Value) + CDbl(fpCurrPenAdj.Value)); dlm;
      '                       37
      Print #RptHandle, OldRound(FEBal + FEVal); dlm; 'changed from above on 5.24.07
      '                        38
      Print #RptHandle, OldRound(MHVal + MHBal); dlm;
      '                        39
      Print #RptHandle, OldRound(PenVal + PenBal); dlm;
      '                          40                               41                          42                        43
      Print #RptHandle, CDbl(fpCurrFEAdj.Value); dlm; CDbl(fpCurrMHAdj.Value); dlm; CDbl(fpCurrPenAdj.Value); dlm; fptxtNote.Text
    
    ElseIf fpcboAdjType.Text = "5-Prepay Adjust Down" Then
      ThisBal = OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))
      ThisAmt = CDbl(fpCurrPrepayAdjBal.Value)
      '                12      13
      Print #RptHandle, 0; dlm; 0; dlm;
      '                14      15
      Print #RptHandle, 0; dlm; 0; dlm;
      '                16      17
      Print #RptHandle, 0; dlm; 0; dlm;
      If CDbl(fpCurrPrepayAdjBal.Value) <> 0 Then fpCurrPrepayAdjBal = -CDbl(fpCurrPrepayAdjBal.Value)
      '                18                 19                                    20                                  21
      Print #RptHandle, 0; dlm; CDbl(fpCurrPrepayBal.Value); dlm; CDbl(fpCurrPrepayAdjAmt.Value); dlm; CDbl(fpCurrPrepayAdjBal.Value); dlm;
      '                          22
      Print #RptHandle, "Prepay Adjust Down"; dlm;
      '                                23
      Print #RptHandle, 0; dlm;
      '                24
      Print #RptHandle, 0; dlm;
      '                25
      Print #RptHandle, 0; dlm;
      '                26
      Print #RptHandle, 0; dlm;
      '                27
      Print #RptHandle, 0; dlm;
      '                28
      Print #RptHandle, 0; dlm;
      '                29
      Print #RptHandle, 0; dlm;
      '                   30           31        32      33
      Print #RptHandle, ThisAmt; dlm; "NA"; dlm; 0; dlm; 0; dlm;
      '                        34                       35                        36
      Print #RptHandle, fptxtFarmEquip.Text; dlm; fptxtMobHomes.Text; dlm; fptxtPenalty.Text; dlm;
      '                 37     38      39       40      41      42           43
      Print #RptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; fptxtNote.Text
    
    End If
  End If
  Close
  
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    arVATaxAdjOPOnlyReport.Show
  Else
    arVATaxPAdjRpt.Show
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPAdjustments", "PrintGraphics", Erl)
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


End Sub

Private Sub ApplyDisc(Handle As Integer, TaxRec As TaxTransactionType)
  Dim Disc1 As Double '1/25/2007
  Dim Disc2 As Double '1/25/2007
  Dim Disc3 As Double '1/25/2007
  Dim Disc4 As Double '1/25/2007
  Dim Disc5 As Double '1/25/2007
  Dim Disc6 As Double '1/25/2007
  Dim Disc7 As Double '1/25/2007
  Dim Disc8 As Double '1/25/2007
  Dim Disc9 As Double '9/19/07
  Dim Dif As Double '9/19/07
  Dim SaveAmt As Double '1/25/2007
  
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  Disc6 = 0
  Disc7 = 0
  Disc8 = 0
  Disc9 = 0
  If TaxRec.Amount = 0 Then Return
  If TaxRec.TranType = 1 Then
    SaveAmt = OldRound(TaxRec.Amount - TaxRec.DiscAmt)
  Else
    SaveAmt = TaxRec.Amount
    TaxRec.Amount = OldRound(TaxRec.Amount + TaxRec.DiscAmt)
  End If
  Disc1 = TaxRec.Revenue.Principle1Pd / SaveAmt
  Disc1 = Disc1 * TaxRec.DiscAmt
  Disc2 = TaxRec.Revenue.Principle2Pd / SaveAmt
  Disc2 = Disc2 * TaxRec.DiscAmt
  Disc3 = TaxRec.Revenue.Principle3Pd / SaveAmt
  Disc3 = Disc3 * TaxRec.DiscAmt
  Disc4 = TaxRec.Revenue.Principle4Pd / SaveAmt
  Disc4 = Disc4 * TaxRec.DiscAmt
  Disc5 = TaxRec.Revenue.Principle5Pd / SaveAmt
  Disc5 = Disc5 * TaxRec.DiscAmt
  Disc6 = TaxRec.Revenue.RevOpt1Pd / SaveAmt
  Disc6 = Disc6 * TaxRec.DiscAmt
  Disc7 = TaxRec.Revenue.RevOpt2Pd / SaveAmt
  Disc7 = Disc7 * TaxRec.DiscAmt
  Disc8 = TaxRec.Revenue.RevOpt3Pd / SaveAmt
  Disc8 = Disc8 * TaxRec.DiscAmt
  Disc9 = TaxRec.Revenue.LateListPd / SaveAmt
  Disc9 = Disc9 * TaxRec.DiscAmt
  
  TaxRec.Revenue.Principle1Pd = OldRound(TaxRec.Revenue.Principle1Pd + Disc1 + Disc9)
  TaxRec.Revenue.Principle2Pd = OldRound(TaxRec.Revenue.Principle2Pd + Disc2)
  TaxRec.Revenue.Principle3Pd = OldRound(TaxRec.Revenue.Principle3Pd + Disc3)
  TaxRec.Revenue.Principle4Pd = OldRound(TaxRec.Revenue.Principle4Pd + Disc4)
  TaxRec.Revenue.Principle5Pd = OldRound(TaxRec.Revenue.Principle5Pd + Disc5)
  TaxRec.Revenue.RevOpt1Pd = OldRound(TaxRec.Revenue.RevOpt1Pd + Disc6)
  TaxRec.Revenue.RevOpt2Pd = OldRound(TaxRec.Revenue.RevOpt2Pd + Disc7)
  TaxRec.Revenue.RevOpt3Pd = OldRound(TaxRec.Revenue.RevOpt3Pd + Disc8)
  
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5 + Disc6 + Disc7 + Disc8 + Disc9)
  If Dif <> 0 Then
    If Disc1 > 0 Or Disc9 > 0 Then
      TaxRec.Revenue.Principle1Pd = OldRound(TaxRec.Revenue.Principle1Pd + Dif)
    ElseIf Disc2 > 0 Then
      TaxRec.Revenue.Principle2Pd = OldRound(TaxRec.Revenue.Principle2Pd + Dif)
    ElseIf Disc3 > 0 Then
      TaxRec.Revenue.Principle3Pd = OldRound(TaxRec.Revenue.Principle3Pd + Dif)
    ElseIf Disc4 > 0 Then
      TaxRec.Revenue.Principle4Pd = OldRound(TaxRec.Revenue.Principle4Pd + Dif)
    ElseIf Disc5 > 0 Then
      TaxRec.Revenue.Principle5Pd = OldRound(TaxRec.Revenue.Principle5Pd + Dif)
    ElseIf Disc6 > 0 Then
      TaxRec.Revenue.RevOpt1Pd = OldRound(TaxRec.Revenue.RevOpt1Pd + Dif)
    ElseIf Disc7 > 0 Then
      TaxRec.Revenue.RevOpt2Pd = OldRound(TaxRec.Revenue.RevOpt2Pd + Dif)
    ElseIf Disc8 > 0 Then
      TaxRec.Revenue.RevOpt3Pd = OldRound(TaxRec.Revenue.RevOpt3Pd + Dif)
    End If
  End If
  
  
End Sub

