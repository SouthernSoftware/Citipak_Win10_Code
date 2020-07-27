VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxAdjustments 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Adjustments"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxAdjustments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAdjType 
      Height          =   375
      Left            =   960
      TabIndex        =   5
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
      ColDesigner     =   "frmVATaxAdjustments.frx":08CA
   End
   Begin VB.Timer MsgAlertTimer 
      Interval        =   50
      Left            =   360
      Top             =   840
   End
   Begin EditLib.fpLongInteger fpLongAcctNum 
      Height          =   375
      Left            =   5055
      TabIndex        =   0
      Top             =   1200
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
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4515
      Width           =   615
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
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4095
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
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3660
      Width           =   4095
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
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4080
      Width           =   4095
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
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   4515
      Width           =   1455
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
      Height          =   405
      Left            =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "The date you enter here will be the date that appears on the 'Payment Entry' screen. The date on that screen is not editable."
      Top             =   1680
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
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
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
   Begin EditLib.fpText fptxtRevOpt2 
      Height          =   375
      Left            =   5880
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6330
      Width           =   2055
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
   Begin EditLib.fpText fptxtRevTax 
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2055
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
      Text            =   "PRINCIPLE"
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
   Begin EditLib.fpText fptxtRevInt 
      Height          =   375
      Left            =   5880
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4345
      Width           =   2055
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
   Begin EditLib.fpText fptxtRecAdvCol 
      Height          =   375
      Left            =   5880
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4745
      Width           =   2055
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
      Text            =   "ADV/COLLECT"
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
   Begin EditLib.fpText fptxtRevLateList 
      Height          =   375
      Left            =   5880
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5144
      Width           =   2055
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
      Text            =   "LATE LISTING"
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
   Begin EditLib.fpText fptxtRevOpt1 
      Height          =   375
      Left            =   5880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5935
      Width           =   2055
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
   Begin EditLib.fpText fptxtRevOpt3 
      Height          =   375
      Left            =   5880
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6730
      Width           =   2055
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
   Begin EditLib.fpCurrency fpCurrPrincOwed 
      Height          =   375
      Left            =   8040
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrPrincAdj 
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrIntOwed 
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4345
      Width           =   1575
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
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   4345
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrAdvColOwed 
      Height          =   375
      Left            =   8040
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4745
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrAdvColAdj 
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   4745
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrLateListOwed 
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5144
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrLateListAdj 
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   5144
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrRevOpt1Owed 
      Height          =   375
      Left            =   8040
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5935
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrRevOpt1Adj 
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   5935
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrRevOpt2Owed 
      Height          =   375
      Left            =   8040
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrRevOpt2Adj 
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   6330
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrRevOpt3Owed 
      Height          =   375
      Left            =   8040
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6730
      Width           =   1575
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
   Begin EditLib.fpCurrency fpCurrRevOpt3Adj 
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Tag             =   "1"
      Top             =   6730
      Width           =   1575
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
      Height          =   375
      Left            =   8040
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1575
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
      Height          =   375
      Left            =   9720
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1575
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
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   7110
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
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1695
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
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1575
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
      Height          =   375
      Left            =   9720
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1575
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
      Left            =   7080
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   1200
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBills 
      Height          =   372
      Left            =   7080
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2160
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":0DA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   2040
      TabIndex        =   60
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":0F7F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   6900
      TabIndex        =   61
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":1161
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   1290
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":133D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHistory 
      Height          =   492
      Left            =   5016
      TabIndex        =   63
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":1519
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMessage 
      Height          =   495
      Left            =   3150
      TabIndex        =   64
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":16F7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReset 
      Height          =   495
      Left            =   8790
      TabIndex        =   65
      TabStop         =   0   'False
      ToolTipText     =   "Press to reset values to zero while maintaining the customer number and adjustment type."
      Top             =   8040
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
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
      ButtonDesigner  =   "frmVATaxAdjustments.frx":18D5
   End
   Begin EditLib.fpText fptxtPenalty 
      Height          =   372
      Left            =   5880
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5540
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
   Begin EditLib.fpCurrency fpCurrPenOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5540
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
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
      Top             =   5540
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4935
      Left            =   240
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5760
      X2              =   11400
      Y1              =   3640
      Y2              =   3640
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
      Left            =   9746
      TabIndex        =   57
      Top             =   2760
      Width           =   1572
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
      Height          =   255
      Left            =   5760
      TabIndex        =   55
      Top             =   2760
      Width           =   975
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
      Left            =   8110
      TabIndex        =   54
      Top             =   2760
      Width           =   1452
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
      Height          =   255
      Left            =   6960
      TabIndex        =   53
      Top             =   2760
      Width           =   975
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
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   2760
      Width           =   2055
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
      TabIndex        =   51
      Top             =   5160
      Width           =   1812
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
      Height          =   330
      Left            =   600
      TabIndex        =   50
      Top             =   7140
      Width           =   705
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
      TabIndex        =   49
      Top             =   5520
      Width           =   2148
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   5760
      Y1              =   5160
      Y2              =   5160
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
      Height          =   255
      Left            =   9960
      TabIndex        =   48
      Top             =   3640
      Width           =   1095
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
      Height          =   255
      Left            =   8280
      TabIndex        =   47
      Top             =   3640
      Width           =   1095
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
      Height          =   255
      Left            =   6360
      TabIndex        =   46
      Top             =   3640
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9675
      X2              =   9675
      Y1              =   2760
      Y2              =   7680
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   2640
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4935
      Left            =   5760
      Top             =   2760
      Width           =   5655
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
      Height          =   255
      Left            =   6075
      TabIndex        =   45
      Top             =   7275
      Width           =   1695
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   8040
      X2              =   9600
      Y1              =   7165
      Y2              =   7165
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   9720
      X2              =   11280
      Y1              =   7165
      Y2              =   7165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7995
      X2              =   7995
      Y1              =   2760
      Y2              =   7680
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
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   2280
      Width           =   1335
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
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   1770
      Width           =   615
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
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   3360
      Width           =   855
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
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   3765
      Width           =   855
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
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   4185
      Width           =   855
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
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   4605
      Width           =   855
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
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   4605
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Tax Billing Adjustments"
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
      Left            =   3648
      TabIndex        =   16
      Top             =   360
      Width           =   4344
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   600
      Left            =   2310
      Top             =   240
      Width           =   7005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cust Acct Number:"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   750
      Left            =   2310
      Top             =   120
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxAdjustments"
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
  Dim PrincPaid As Double
  Dim IntPaid As Double
  Dim AdvColPaid As Double
  Dim LateListPaid As Double
  Dim PenPaid As Double
  Dim Rev1Paid As Double
  Dim Rev2Paid As Double
  Dim Rev3Paid As Double
  Dim TotPaid As Double
  Dim PrincBal As Double
  Dim IntBal As Double
  Dim AdvColBal As Double
  Dim LateListBal As Double
  Dim PenBal As Double
  Dim Rev1Bal As Double
  Dim Rev2Bal As Double
  Dim Rev3Bal As Double
  Public ThisBillBal As Double
  Dim PrincVal As Double
  Dim IntVal As Double
  Dim AdvColVal As Double
  Dim LateListVal As Double
  Dim PenVal As Double
  Dim Rev1Val As Double
  Dim Rev2Val As Double
  Dim Rev3Val As Double
  Dim FirstLoad As Boolean
  Public ThisBillNum As Long
  Dim NewBalThisBill As Double
  Public ThisBillType$
  
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "cmdBills_Click", Erl)
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
  KillFile "C:\CPWork\txradjust.dat"
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
  Dim SaveBillRec As Long '1/25/07
  Dim NextRec As Long '1/25/07
  
  On Error GoTo ERRORSTUFF
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If GCustNum = 0 Then
    Call TaxMsg(900, "ERROR: No customer record number has been assigned.")
    Exit Sub
  End If
  
  If BillCnt = 0 Then
    If fpcboAdjType.Text <> "5-Prepay Adjust Down" Then
      Call TaxMsg(900, "ERROR: This customer has no real tax bills to adjust.")
      fpLngIntBill.SetFocus
      Exit Sub
    End If
  End If
    
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    If fpCurrPrepayAdjAmt = 0 Then
      Call TaxMsg(800, "There amount for 'Prepay Adjustment Amount' is zero. Nothing to post.")
      Close
      If fpCurrPrepayAdjAmt.Enabled = True Then
        fpCurrPrepayAdjAmt.SetFocus
      End If
      Exit Sub 'added 11/8/2007
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
    If CDbl(fpCurrPrincAdj.Value) > CDbl(fpCurrPrincOwed.Value) Then
      Call TaxMsg(700, "The 'PRINCIPLE' adjustment amount is greater than the 'PRINCIPLE' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrPrincAdj.SetFocus
      Exit Sub
    End If
   
    If CDbl(fpCurrIntAdj.Value) > CDbl(fpCurrIntOwed.Value) Then
      Call TaxMsg(700, "The 'INTEREST' adjustment amount is greater than the 'INTEREST' amount owed. Please revise the adjustment aamount so that it is less than or equal to the amount owed.")
      fpCurrIntAdj.SetFocus
      Exit Sub
    End If
  
    If CDbl(fpCurrAdvColAdj.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      Call TaxMsg(700, "The 'ADV/COLLECT' adjustment amount is greater than the 'ADV/COLLECT' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrAdvColAdj.SetFocus
      Exit Sub
    End If
  
    If CDbl(fpCurrLateListAdj.Value) > CDbl(fpCurrLateListOwed.Value) Then
      Call TaxMsg(700, "The 'LATE LISTING' adjustment amount is greater than the 'LATE LISTING' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrLateListAdj.SetFocus
      Exit Sub
    End If
  
    If CDbl(fpCurrPenAdj.Value) > CDbl(fpCurrPenOwed.Value) Then
      Call TaxMsg(700, "The 'PENALTY' adjustment amount is greater than the 'PENALTY' amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
      fpCurrPenAdj.SetFocus
      Exit Sub
    End If
  
    If fpCurrRevOpt1Adj.Enabled = True Then
      If CDbl(fpCurrRevOpt1Adj.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
        Call TaxMsg(700, "The " + QPTrim$(fptxtRevOpt1.Text) + " adjustment amount is greater than the " + QPTrim$(fptxtRevOpt1.Text) + " amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
        fpCurrRevOpt1Adj.SetFocus
        Exit Sub
      End If
    End If
  
    If fpCurrRevOpt2Adj.Enabled = True Then
      If CDbl(fpCurrRevOpt2Adj.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
        Call TaxMsg(700, "The " + QPTrim$(fptxtRevOpt2.Text) + " adjustment amount is greater than the " + QPTrim$(fptxtRevOpt2.Text) + " amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
        fpCurrRevOpt2Adj.SetFocus
        Exit Sub
      End If
    End If
  
    If fpCurrRevOpt3Adj.Enabled = True Then
      If CDbl(fpCurrRevOpt3Adj.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
        Call TaxMsg(700, "The " + QPTrim$(fptxtRevOpt3.Text) + " adjustment amount is greater than the " + QPTrim$(fptxtRevOpt3.Text) + " amount owed. Please revise the adjustment amount so that it is less than or equal to the amount owed.")
        fpCurrRevOpt3Adj.SetFocus
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
  TaxAdjTrans.RealPin = RealPin
  TaxAdjTrans.PersPin = PersPin
  TaxAdjTrans.CustPin = TaxCustRec.PIN
  TaxAdjTrans.OperNum = OperNum
  TaxAdjTrans.BillType = "R"
  Put #TTHandle, BillRec, TaxTrans
PrePayOnly:
  If QPTrim$(fpcboAdjType.Text) <> "5-Prepay Adjust Down" Then
    TaxAdjTrans.RealPin = RealPin 'added 1/29/08
  End If
'  TaxAdjTrans.RealPin = RealPin 'added 1/29/08
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
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  TaxAdjTrans.TranType = 3              'Release
  TaxAdjTrans.BillType = "R"
  '7/12/06 changed the revenue for Release from charged to paid (like in Dos)
  TaxAdjTrans.Revenue.Principle1Pd = CDbl(fpCurrPrincAdj.Value)
  TaxAdjTrans.Revenue.InterestPd = CDbl(fpCurrIntAdj.Value)
  TaxAdjTrans.Revenue.CollectionPd = CDbl(fpCurrAdvColAdj.Value)
  TaxAdjTrans.Revenue.LateListPd = CDbl(fpCurrLateListAdj.Value)
  TaxAdjTrans.Revenue.PenaltyPd = CDbl(fpCurrPenAdj.Value)
  TaxAdjTrans.Revenue.RevOpt1Pd = CDbl(fpCurrRevOpt1Adj.Value)
  TaxAdjTrans.Revenue.RevOpt2Pd = CDbl(fpCurrRevOpt2Adj.Value)
  TaxAdjTrans.Revenue.RevOpt3Pd = CDbl(fpCurrRevOpt3Adj.Value)
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
  TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + CDbl(fpCurrPrincAdj.Value))
  TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd + CDbl(fpCurrIntAdj.Value))
  TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd + CDbl(fpCurrAdvColAdj.Value))
  TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd + CDbl(fpCurrLateListAdj.Value))
  TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd + CDbl(fpCurrPenAdj.Value))
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + CDbl(fpCurrRevOpt1Adj.Value))
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + CDbl(fpCurrRevOpt2Adj.Value))
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + CDbl(fpCurrRevOpt3Adj.Value))
  
  Return
  
AdjustPayDown:
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  TaxAdjTrans.BillType = "R"
  CreditAmt = CDbl(fpCurrPrepayBal.Value)
  CreditBalance = CDbl(fpCurrPrepayBal.Value)
  If CreditAmt <= 0 Then
    TaxAdjTrans.TranType = 7
    TaxAdjTrans.Revenue.PrePaidBal = CDbl(fpCurrPrepayBal.Value)
    TaxAdjTrans.Revenue.PrePaidAmt = 0
    TaxAdjTrans.Revenue.PrePaidUsed = 0
    TaxAdjTrans.Revenue.Principle1Pd = CDbl(fpCurrPrincAdj.Value)
    PrincVal = CDbl(fpCurrPrincAdj.Value)
    TaxAdjTrans.Revenue.InterestPd = CDbl(fpCurrIntAdj.Value)
    IntVal = CDbl(fpCurrIntAdj.Value)
    TaxAdjTrans.Revenue.CollectionPd = CDbl(fpCurrAdvColAdj.Value)
    AdvColVal = CDbl(fpCurrAdvColAdj.Value)
    TaxAdjTrans.Revenue.LateListPd = CDbl(fpCurrLateListAdj.Value)
    LateListVal = CDbl(fpCurrLateListAdj.Value)
    TaxAdjTrans.Revenue.PenaltyPd = CDbl(fpCurrPenAdj.Value)
    PenVal = CDbl(fpCurrPenAdj.Value)
    TaxAdjTrans.Revenue.RevOpt1Pd = CDbl(fpCurrRevOpt1Adj.Value)
    Rev1Val = CDbl(fpCurrRevOpt1Adj.Value)
    TaxAdjTrans.Revenue.RevOpt2Pd = CDbl(fpCurrRevOpt2Adj.Value)
    Rev2Val = CDbl(fpCurrRevOpt2Adj.Value)
    TaxAdjTrans.Revenue.RevOpt3Pd = CDbl(fpCurrRevOpt3Adj.Value)
    Rev3Val = CDbl(fpCurrRevOpt3Adj.Value)
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
    BillNum$ = ParseBillNum(TaxTrans.Description)
    If Len(QPTrim$(fptxtNote.Text)) = 0 Then
      TaxAdjTrans.Description = "Tax Adj Pay Down #" + BillNum$
    Else
      TaxAdjTrans.Description = QPTrim$(fptxtNote.Text) + ": Bill #" + BillNum$
    End If
    TaxTrans.DiscAmt = 0 '1/25/07
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd - CDbl(fpCurrPrincAdj.Value))
    TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd - CDbl(fpCurrIntAdj.Value))
    TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd - CDbl(fpCurrAdvColAdj.Value))
    TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd - CDbl(fpCurrLateListAdj.Value))
    TaxTrans.Revenue.PenaltyPd = OldRound#(TaxTrans.Revenue.PenaltyPd - CDbl(fpCurrPenAdj.Value))
    TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd - CDbl(fpCurrRevOpt1Adj.Value))
    TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd - CDbl(fpCurrRevOpt2Adj.Value))
    TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd - CDbl(fpCurrRevOpt3Adj.Value))
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
    
    If CDbl(fpCurrRevOpt3Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrRevOpt3Adj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.RevOpt3Pd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrRevOpt3Adj.Value))
        Rev3Val = 0
      Else
        TaxAdjTrans.Revenue.RevOpt3Pd = OldRound(CDbl(fpCurrRevOpt3Adj.Value) - CreditAmt)
        Rev3Val = OldRound(CDbl(fpCurrRevOpt3Adj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrRevOpt2Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrRevOpt2Adj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.RevOpt2Pd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrRevOpt2Adj.Value))
        Rev2Val = 0
      Else
        TaxAdjTrans.Revenue.RevOpt2Pd = OldRound(CDbl(fpCurrRevOpt2Adj.Value) - CreditAmt)
        Rev2Val = OldRound(CDbl(fpCurrRevOpt2Adj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrRevOpt1Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrRevOpt1Adj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.RevOpt1Pd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrRevOpt1Adj.Value))
        Rev1Val = 0
      Else
        TaxAdjTrans.Revenue.RevOpt1Pd = OldRound(CDbl(fpCurrRevOpt1Adj.Value) - CreditAmt)
        Rev1Val = OldRound(CDbl(fpCurrRevOpt1Adj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrPrincAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt >= CDbl(fpCurrPrincAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.Principle1Pd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPrincAdj.Value))
        PrincVal = 0
      Else
        TaxAdjTrans.Revenue.Principle1Pd = OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
        PrincVal = OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
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
    
    If CDbl(fpCurrLateListAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrLateListAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.LateListPd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrLateListAdj.Value))
        LateListVal = 0
      Else
        TaxAdjTrans.Revenue.LateListPd = OldRound(CDbl(fpCurrLateListAdj.Value) - CreditAmt)
        LateListVal = OldRound(CDbl(fpCurrLateListAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrAdvColAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrAdvColAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.CollectionPd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrAdvColAdj.Value))
        AdvColVal = 0
      Else
        TaxAdjTrans.Revenue.CollectionPd = OldRound(CDbl(fpCurrAdvColAdj.Value) - CreditAmt)
        AdvColVal = OldRound(CDbl(fpCurrAdvColAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrIntAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrIntAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.InterestPd = 0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrIntAdj.Value))
        IntVal = 0
      Else
        TaxAdjTrans.Revenue.InterestPd = OldRound(CDbl(fpCurrIntAdj.Value) - CreditAmt)
        IntVal = OldRound(CDbl(fpCurrIntAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
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
    
    ThisAmt = OldRound(CDbl(fpCurrIntOwed.Value) - IntVal)
    TaxTrans.Revenue.InterestPd = ThisAmt
    
    ThisAmt = OldRound(CDbl(fpCurrPrincOwed.Value) - PrincVal)
    TaxTrans.Revenue.Principle1Pd = ThisAmt
    
    ThisAmt = OldRound(CDbl(fpCurrAdvColOwed.Value) - AdvColVal)
    TaxTrans.Revenue.CollectionPd = ThisAmt
    
    ThisAmt = OldRound(CDbl(fpCurrLateListOwed.Value) - LateListVal)
    TaxTrans.Revenue.LateListPd = ThisAmt
    
    ThisAmt = OldRound(CDbl(fpCurrPenOwed.Value) - PenVal)
    TaxTrans.Revenue.PenaltyPd = ThisAmt
    
    ThisAmt = OldRound(CDbl(fpCurrRevOpt1Owed.Value) - Rev1Val)
    TaxTrans.Revenue.RevOpt1Pd = ThisAmt
  
    ThisAmt = OldRound(CDbl(fpCurrRevOpt2Owed.Value) - Rev2Val)
    TaxTrans.Revenue.RevOpt2Pd = ThisAmt
  
    ThisAmt = OldRound(CDbl(fpCurrRevOpt3Owed.Value) - Rev3Val)
    TaxTrans.Revenue.RevOpt3Pd = ThisAmt
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
  TaxAdjTrans.Revenue.Principle1 = CDbl(fpCurrPrincAdj.Value)
  TaxAdjTrans.Revenue.Interest = CDbl(fpCurrIntAdj.Value)
  TaxAdjTrans.Revenue.Collection = CDbl(fpCurrAdvColAdj.Value)
  TaxAdjTrans.Revenue.LateList = CDbl(fpCurrLateListAdj.Value)
  TaxAdjTrans.Revenue.Penalty = CDbl(fpCurrPenAdj.Value)
  TaxAdjTrans.Revenue.RevOpt1 = CDbl(fpCurrRevOpt1Adj.Value)
  TaxAdjTrans.Revenue.RevOpt2 = CDbl(fpCurrRevOpt2Adj.Value)
  TaxAdjTrans.Revenue.RevOpt3 = CDbl(fpCurrRevOpt3Adj.Value)
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
  TaxTrans.Revenue.Principle1 = OldRound#(TaxTrans.Revenue.Principle1 - CDbl(fpCurrPrincAdj.Value))
  TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest - CDbl(fpCurrIntAdj.Value))
  TaxTrans.Revenue.Collection = OldRound#(TaxTrans.Revenue.Collection - CDbl(fpCurrAdvColAdj.Value))
  TaxTrans.Revenue.LateList = OldRound#(TaxTrans.Revenue.LateList - CDbl(fpCurrLateListAdj.Value))
  TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty - CDbl(fpCurrPenAdj.Value))
  TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 - CDbl(fpCurrRevOpt1Adj.Value))
  TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 - CDbl(fpCurrRevOpt2Adj.Value))
  TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 - CDbl(fpCurrRevOpt3Adj.Value))
  
  Return
  
AdjustBillUp:
  TotalAdj# = CDbl(fpCurrTotAdj.Value)
  TaxAdjTrans.TransDate = Date2Num(fptxtDate.Text)
  TaxAdjTrans.BillType = "R"
  CreditAmt = CDbl(fpCurrPrepayBal.Value)
  CreditBalance = CDbl(fpCurrPrepayBal.Value)
  If CreditAmt <= 0 Then
    TaxAdjTrans.TranType = 14 'adjust bill up with no affect on credit balance
    TaxAdjTrans.Revenue.Principle1 = CDbl(fpCurrPrincAdj.Value)
    PrincVal = CDbl(fpCurrPrincAdj.Value)
    TaxAdjTrans.Revenue.Interest = CDbl(fpCurrIntAdj.Value)
    IntVal = CDbl(fpCurrIntAdj.Value)
    TaxAdjTrans.Revenue.Collection = CDbl(fpCurrAdvColAdj.Value)
    AdvColVal = CDbl(fpCurrAdvColAdj.Value)
    TaxAdjTrans.Revenue.LateList = CDbl(fpCurrLateListAdj.Value)
    LateListVal = CDbl(fpCurrLateListAdj.Value)
    TaxAdjTrans.Revenue.Penalty = CDbl(fpCurrPenAdj.Value)
    PenVal = CDbl(fpCurrPenAdj.Value)
    TaxAdjTrans.Revenue.RevOpt1 = CDbl(fpCurrRevOpt1Adj.Value)
    Rev1Val = CDbl(fpCurrRevOpt1Adj.Value)
    TaxAdjTrans.Revenue.RevOpt2 = CDbl(fpCurrRevOpt2Adj.Value)
    Rev2Val = CDbl(fpCurrRevOpt2Adj.Value)
    TaxAdjTrans.Revenue.RevOpt3 = CDbl(fpCurrRevOpt3Adj.Value)
    Rev3Val = CDbl(fpCurrRevOpt3Adj.Value)
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
    TaxTrans.Revenue.Principle1 = OldRound#(TaxTrans.Revenue.Principle1 + CDbl(fpCurrPrincAdj.Value))
    TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + CDbl(fpCurrIntAdj.Value))
    TaxTrans.Revenue.Collection = OldRound#(TaxTrans.Revenue.Collection + CDbl(fpCurrAdvColAdj.Value))
    TaxTrans.Revenue.LateList = OldRound#(TaxTrans.Revenue.LateList + CDbl(fpCurrLateListAdj.Value))
    TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty + CDbl(fpCurrPenAdj.Value))
    TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 + CDbl(fpCurrRevOpt1Adj.Value))
    TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 + CDbl(fpCurrRevOpt2Adj.Value))
    TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 + CDbl(fpCurrRevOpt3Adj.Value))
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
    If CDbl(fpCurrIntAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrIntAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.InterestPd = CDbl(fpCurrIntAdj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrIntAdj.Value))
      Else
        TaxAdjTrans.Revenue.InterestPd = CreditAmt 'OldRound(CDbl(fpCurrIntAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrAdvColAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrAdvColAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.CollectionPd = CDbl(fpCurrAdvColAdj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrAdvColAdj.Value))
      Else
        TaxAdjTrans.Revenue.CollectionPd = CreditAmt 'OldRound(CDbl(fpCurrAdvColAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrLateListAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrLateListAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.LateListPd = CDbl(fpCurrLateListAdj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrLateListAdj.Value))
      Else
        TaxAdjTrans.Revenue.LateListPd = CreditAmt 'OldRound(CDbl(fpCurrLateListAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrPenAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrPenAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.PenaltyPd = CDbl(fpCurrPenAdj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPenAdj.Value))
      Else
        TaxAdjTrans.Revenue.PenaltyPd = CreditAmt 'OldRound(CDbl(fpCurrPenAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrPrincAdj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrPrincAdj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.Principle1Pd = CDbl(fpCurrPrincAdj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrPrincAdj.Value))
        PrincVal = 0
      Else
        TaxAdjTrans.Revenue.Principle1Pd = CreditAmt 'OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
        PrincVal = OldRound(CDbl(fpCurrPrincAdj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrRevOpt1Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrRevOpt1Adj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.RevOpt1Pd = CDbl(fpCurrRevOpt1Adj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrRevOpt1Adj.Value))
      Else
        TaxAdjTrans.Revenue.RevOpt1Pd = CreditAmt 'OldRound(CDbl(fpCurrRevOpt1Adj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrRevOpt2Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrRevOpt2Adj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.RevOpt2Pd = CDbl(fpCurrRevOpt2Adj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrRevOpt2Adj.Value))
      Else
        TaxAdjTrans.Revenue.RevOpt2Pd = CreditAmt 'OldRound(CDbl(fpCurrRevOpt2Adj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
    If CDbl(fpCurrRevOpt3Adj.Value) > 0 Then 'check to see if an adjustment amount (> 0) is entered
      If CreditAmt > CDbl(fpCurrRevOpt3Adj.Value) Then 'if the credit available can cover this adjustment
        TaxAdjTrans.Revenue.RevOpt3Pd = CDbl(fpCurrRevOpt3Adj.Value) '0
        CreditAmt = OldRound(CreditAmt - CDbl(fpCurrRevOpt3Adj.Value))
      Else
        TaxAdjTrans.Revenue.RevOpt3Pd = CreditAmt 'OldRound(CDbl(fpCurrRevOpt3Adj.Value) - CreditAmt)
        CreditAmt = 0
      End If
    End If
    
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
    TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd + TaxAdjTrans.Revenue.InterestPd)
    TaxTrans.Revenue.Interest = OldRound#(TaxTrans.Revenue.Interest + CDbl(fpCurrIntAdj.Value))
    
    TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd + TaxAdjTrans.Revenue.Principle1Pd)
    TaxTrans.Revenue.Principle1 = OldRound#(TaxTrans.Revenue.Principle1 + CDbl(fpCurrPrincAdj.Value))
    
    TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd + TaxAdjTrans.Revenue.CollectionPd)
    TaxTrans.Revenue.Collection = OldRound#(TaxTrans.Revenue.Collection + CDbl(fpCurrAdvColAdj.Value))
    
    TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd + TaxAdjTrans.Revenue.LateListPd)
    TaxTrans.Revenue.LateList = OldRound#(TaxTrans.Revenue.LateList + CDbl(fpCurrLateListAdj.Value))
    
    TaxTrans.Revenue.PenaltyPd = OldRound#(TaxTrans.Revenue.PenaltyPd + TaxAdjTrans.Revenue.PenaltyPd)
    TaxTrans.Revenue.Penalty = OldRound#(TaxTrans.Revenue.Penalty + CDbl(fpCurrPenAdj.Value))
    
    TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd + TaxAdjTrans.Revenue.RevOpt1Pd)
    TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 + CDbl(fpCurrRevOpt1Adj.Value))
  
    TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd + TaxAdjTrans.Revenue.RevOpt2Pd)
    TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 + CDbl(fpCurrRevOpt2Adj.Value))
  
    TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd + TaxAdjTrans.Revenue.RevOpt3Pd)
    TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 + CDbl(fpCurrRevOpt3Adj.Value))
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
  TaxAdjTrans.BillType = "R"
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "cmdPost_Click", Erl)
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
      KillFile "C:\CPWork\txradjust.dat"
      GCustNum = 0
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxAdjustments.")
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
  ThisDesc = QPTrim$(TaxMasterRec.OptRev1)
  If ThisDesc <> "" Then
    fptxtRevOpt1.Text = QPTrim$(TaxMasterRec.OptRev1)
  Else
    fptxtRevOpt1.Text = "NOT IN USE"
    fpCurrRevOpt1Adj.Enabled = False
  End If
  
  ThisDesc = QPTrim$(TaxMasterRec.OptRev2)
  If ThisDesc <> "" Then
    fptxtRevOpt2.Text = QPTrim$(TaxMasterRec.OptRev2)
  Else
    fptxtRevOpt2.Text = "NOT IN USE"
    fpCurrRevOpt2Adj.Enabled = False
  End If
  
  ThisDesc = QPTrim$(TaxMasterRec.OptRev3)
  If ThisDesc <> "" Then
    fptxtRevOpt3.Text = QPTrim$(TaxMasterRec.OptRev3)
  Else
    fptxtRevOpt3.Text = "NOT IN USE"
    fpCurrRevOpt3Adj.Enabled = False
  End If
  
  One = 1
  AHandle = FreeFile
  Open "C:\CPWork\txradjust.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  fptxtDate.Text = Date
  fpcboAdjType.AddItem "1-Billing Downward Adjustment"
  fpcboAdjType.AddItem "2-Billing Upward Adjustment"
  fpcboAdjType.AddItem "3-Adjustment for Payment" 'Downward Adjustment"
  fpcboAdjType.AddItem "4-Release"
  fpcboAdjType.ListIndex = 0
  TempAcctNum = 0
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "LoadMe", Erl)
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
    If BillType = "P" Then GoTo GoAhead
      Call TaxMsg(700, "This customer is included in a real payment file for operator #" + OpNum + " that has not been posted. Please either post this payment or delete this customer from the payment file.")
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
    ThisGBal = GetCustRealBalance(GCustNum, -1)
  End If
  If CustHasMsg(GCustNum) Then
    MsgAlertTimer.Enabled = True
  Else
    MsgAlertTimer.Enabled = False
    cmdMessage.ForeColor = &H80000012
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "LoadMeEdit", Erl)
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
  fpCurrPrincOwed = PrincBal
  fpCurrAdvColOwed = AdvColBal
  fpCurrIntOwed = IntBal
  fpCurrLateListOwed = LateListBal
  fpCurrPenOwed = PenBal
  fpCurrRevOpt1Owed = Rev1Bal
  fpCurrRevOpt2Owed = Rev2Bal
  fpCurrRevOpt3Owed = Rev3Bal
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
    fpCurrPrincAdj.Enabled = True
    fpCurrIntAdj.Enabled = True
    fpCurrAdvColAdj.Enabled = True
    fpCurrLateListAdj.Enabled = True
    fpCurrPenAdj.Enabled = True
    If QPTrim$(fptxtRevOpt1.Text) <> "NOT IN USE" Then
      fpCurrRevOpt1Adj.Enabled = True
    End If
    If QPTrim$(fptxtRevOpt2.Text) <> "NOT IN USE" Then
      fpCurrRevOpt2Adj.Enabled = True
    End If
    If QPTrim$(fptxtRevOpt3.Text) <> "NOT IN USE" Then
      fpCurrRevOpt3Adj.Enabled = True
    End If
  ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then
    Label10.Caption = "Paid"
    fpCurrPrincOwed = PrincPaid
    fpCurrAdvColOwed = AdvColPaid
    fpCurrIntOwed = IntPaid
    fpCurrLateListOwed = LateListPaid
    fpCurrPenOwed = PenPaid
    fpCurrRevOpt1Owed = Rev1Paid
    fpCurrRevOpt2Owed = Rev2Paid
    fpCurrRevOpt3Owed = Rev3Paid
    fpCurrTotOwed = TotPaid
  Else
    cmdBills.Enabled = False
    fpLngIntBill = 0
    fpCurrPrepayAdjAmt.Enabled = True
    fpCurrPrincAdj.Enabled = False
    fpCurrIntAdj.Enabled = False
    fpCurrAdvColAdj.Enabled = False
    fpCurrLateListAdj.Enabled = False
    fpCurrPenAdj.Enabled = False
    fpCurrRevOpt1Adj.Enabled = False
    fpCurrRevOpt2Adj.Enabled = False
    fpCurrRevOpt3Adj.Enabled = False
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

Private Sub fpCurrAdvColAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrAdvColAdj.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrAdvColAdj = CDbl(fpCurrAdvColOwed.Value)
      fpCurrAdvColAdj.SetFocus
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
  Call FigureAdjCol(2)
End Sub

Private Sub fpCurrLateListAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrLateListAdj.Value) > CDbl(fpCurrLateListOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrLateListAdj = CDbl(fpCurrLateListOwed.Value)
      fpCurrLateListAdj.SetFocus
    End If
  End If
  Call FigureAdjCol(4)
End Sub

Private Sub fpCurrPenAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrPenAdj.Value) > CDbl(fpCurrPenOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrPenAdj = CDbl(fpCurrPenOwed.Value)
      fpCurrPenAdj.SetFocus
    End If
  End If
  
  Call FigureAdjCol(1)

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

Private Sub fpCurrPrincAdj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrPrincAdj.Value) > CDbl(fpCurrPrincOwed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrPrincAdj = CDbl(fpCurrPrincOwed.Value)
      fpCurrPrincAdj.SetFocus
    End If
  End If
  
  Call FigureAdjCol(1)
End Sub

Private Sub fpCurrRevOpt1Adj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpCurrRevOpt2Adj.Enabled = True Then
      fpCurrRevOpt2Adj.SetFocus
    ElseIf fpCurrRevOpt3Adj.Enabled = True Then
      fpCurrRevOpt3Adj.SetFocus
    Else
      fpLongAcctNum.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    fpCurrLateListAdj.SetFocus
  End If
End Sub

Private Sub fpCurrRevOpt1Adj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrRevOpt1Adj.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrRevOpt1Adj = CDbl(fpCurrRevOpt1Owed.Value)
      fpCurrRevOpt1Adj.SetFocus
    End If
  End If
  Call FigureAdjCol(5)
End Sub

Private Sub fpCurrRevOpt2Adj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpCurrRevOpt3Adj.Enabled = True Then
      fpCurrRevOpt3Adj.SetFocus
    Else
      fptxtNote.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrRevOpt1Adj.Enabled = True Then
      fpCurrRevOpt1Adj.SetFocus
    Else
      fpCurrLateListAdj.SetFocus
    End If
  End If

End Sub

Private Sub fpCurrRevOpt2Adj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrRevOpt2Adj.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrRevOpt2Adj = CDbl(fpCurrRevOpt2Owed.Value)
      fpCurrRevOpt2Adj.SetFocus
    End If
  End If
  Call FigureAdjCol(6)
End Sub

Private Sub fpCurrRevOpt3Adj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpLongAcctNum.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrRevOpt2Adj.Enabled = True Then
      fpCurrRevOpt2Adj.SetFocus
    ElseIf fpCurrRevOpt1Adj.Enabled = True Then
      fpCurrRevOpt1Adj.SetFocus
    Else
      fpCurrLateListAdj.SetFocus
    End If
  End If
End Sub

Private Sub fpCurrRevOpt3Adj_LostFocus()
  If fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrRevOpt3Adj.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      Call TaxMsg(900, "Payment adjustment cannot exceed the amount paid.")
      fpCurrRevOpt3Adj = CDbl(fpCurrRevOpt3Owed.Value)
      fpCurrRevOpt3Adj.SetFocus
    End If
  End If
  Call FigureAdjCol(7)
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
    If fpCurrRevOpt3Adj.Enabled = True Then
      fpCurrRevOpt3Adj.SetFocus
    ElseIf fpCurrRevOpt2Adj.Enabled = True Then
      fpCurrRevOpt2Adj.SetFocus
    ElseIf fpCurrRevOpt1Adj.Enabled = True Then
      fpCurrRevOpt1Adj.SetFocus
    Else
      fpCurrLateListAdj.SetFocus
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
    PrePay = GetCustRealBalance(CLng(fpLongAcctNum.Value), -1)
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "fpLongAcctNum_LostFocus", Erl)
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "Check4ValidCustNum", Erl)
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
  fpCurrPrincOwed.Text = 0
  fpCurrPrincAdj.Text = 0
  fpCurrIntOwed.Text = 0
  fpCurrIntAdj.Text = 0
  fpCurrAdvColOwed.Text = 0
  fpCurrAdvColAdj.Text = 0
  fpCurrLateListOwed.Text = 0
  fpCurrLateListAdj.Text = 0
  fpCurrPenOwed.Text = 0
  fpCurrPenAdj.Text = 0
  fpCurrRevOpt1Owed.Text = 0
  fpCurrRevOpt1Adj.Text = 0
  fpCurrRevOpt2Owed.Text = 0
  fpCurrRevOpt2Adj.Text = 0
  fpCurrRevOpt3Owed.Text = 0
  fpCurrRevOpt3Adj.Text = 0
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
  Dim TPrinc#, TInterest#, TCollect#, TPen#
  Dim TLateList#, TOpt1#, TOpt2#, TOpt3#
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxTransFile TRHandle, NumOfTRRecs
  Get TRHandle, BillRec, TaxRec
  Close TRHandle
  If TaxRec.DiscAmt > 0 Then Call ApplyDisc(TRHandle, TaxRec) 'added 1/25/07
  PrincPaid = TaxRec.Revenue.Principle1Pd
  IntPaid = TaxRec.Revenue.InterestPd
  AdvColPaid = TaxRec.Revenue.CollectionPd
  LateListPaid = TaxRec.Revenue.LateListPd
  PenPaid = TaxRec.Revenue.PenaltyPd
  Rev1Paid = TaxRec.Revenue.RevOpt1Pd
  Rev2Paid = TaxRec.Revenue.RevOpt2Pd
  Rev3Paid = TaxRec.Revenue.RevOpt3Pd
  TotPaid = OldRound(PrincPaid + IntPaid + AdvColPaid + LateListPaid + Rev1Paid + Rev2Paid + Rev3Paid + PenPaid) 'added PenPaid on 5.24.07
  BillNum& = CLng(ParseBillNum(TaxRec.Description))
  If BillNum = CLng(fpLngIntBill.Value) Then
    TPrinc# = OldRound(TaxRec.Revenue.Principle1 + TaxRec.Revenue.Principle2 + TaxRec.Revenue.Principle3 + TaxRec.Revenue.Principle4 + TaxRec.Revenue.Principle5)
    TPrinc# = OldRound(TPrinc# - (TaxRec.Revenue.Principle1Pd + TaxRec.Revenue.Principle2Pd + TaxRec.Revenue.Principle3Pd + TaxRec.Revenue.Principle4Pd + TaxRec.Revenue.Principle5Pd)) 'took out + TaxRec.DiscAmt on 1/16/07
    PrincBal = TPrinc#
    TInterest# = OldRound#(TaxRec.Revenue.Interest - TaxRec.Revenue.InterestPd)
    IntBal = TInterest#
    TCollect# = OldRound#(TaxRec.Revenue.Collection - TaxRec.Revenue.CollectionPd)
    AdvColBal = TCollect#
    TLateList# = OldRound(TaxRec.Revenue.LateList - TaxRec.Revenue.LateListPd)
    LateListBal = TLateList#
    TPen# = OldRound(TaxRec.Revenue.Penalty - TaxRec.Revenue.PenaltyPd)
    PenBal# = TPen#
    TOpt1# = OldRound(TaxRec.Revenue.RevOpt1 - TaxRec.Revenue.RevOpt1Pd)
    Rev1Bal = TOpt1#
    TOpt2# = OldRound(TaxRec.Revenue.RevOpt2 - TaxRec.Revenue.RevOpt2Pd)
    Rev2Bal = TOpt2#
    TOpt3# = OldRound(TaxRec.Revenue.RevOpt3 - TaxRec.Revenue.RevOpt3Pd)
    Rev3Bal = TOpt3#
    fpCurrPrincOwed.Text = TPrinc#
    fpCurrPrincAdj.Text = 0
    fpCurrIntOwed.Text = TInterest#
    fpCurrIntAdj.Text = 0
    fpCurrAdvColOwed.Text = TCollect#
    fpCurrAdvColAdj.Text = 0
    fpCurrLateListOwed.Text = TLateList#
    fpCurrLateListAdj.Text = 0
    fpCurrPenOwed.Text = TPen#
    fpCurrPenAdj.Text = 0
    fpCurrRevOpt1Owed.Text = TOpt1#
    fpCurrRevOpt1Adj.Text = 0
    fpCurrRevOpt2Owed.Text = TOpt2#
    fpCurrRevOpt2Adj.Text = 0
    fpCurrRevOpt3Owed.Text = TOpt3#
    fpCurrRevOpt3Adj.Text = 0
    fpCurrTotOwed.Text = OldRound(TPrinc# + TInterest# + TCollect# + TLateList# + TPen# + TOpt1# + TOpt2# + TOpt3#)
    fpCurrTotAdj.Text = 0
  Else
    Call TaxMsg(800, "The bill number entered could not be found in this customer's transaction records. Please enter another bill number or press F8 to bring up a complete list of all the bills for this customer.")
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "LoadMeBill", Erl)
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
  TotAdj = OldRound(CDbl(fpCurrPrincAdj.Value) + CDbl(fpCurrIntAdj.Value) + CDbl(fpCurrAdvColAdj.Value))
  TotAdj = OldRound(TotAdj + CDbl(fpCurrLateListAdj.Value) + CDbl(fpCurrPenAdj.Value) + CDbl(fpCurrRevOpt1Adj.Value))
  TotAdj = OldRound(TotAdj + CDbl(fpCurrRevOpt2Adj.Value) + CDbl(fpCurrRevOpt3Adj.Value))
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
    Print #RptHandle, Tab(5); QPTrim$(fptxtRevTax.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPrincAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrPrincOwed.Value) - CDbl(fpCurrPrincAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtRevInt.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrIntAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrIntOwed.Value) - CDbl(fpCurrIntAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtRecAdvCol.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrAdvColAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrAdvColOwed.Value) - CDbl(fpCurrAdvColAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtRevLateList.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrLateListAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrLateListOwed.Value) - CDbl(fpCurrLateListAdj.Value)))
    Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using$("$##,##0.00", CDbl(fpCurrPenAdj.Value)); Tab(60); Using$("$##,##0.00", OldRound(CDbl(fpCurrPenOwed.Value) - CDbl(fpCurrPenAdj.Value)))
    If QPTrim$(fptxtRevOpt1.Text) <> "" Then
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt1.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt1Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrRevOpt1Owed.Value) - CDbl(fpCurrRevOpt1Adj.Value)))
    End If
    If QPTrim$(fptxtRevOpt2.Text) <> "" Then
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt2.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt2Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrRevOpt2Owed.Value) - CDbl(fpCurrRevOpt2Adj.Value)))
    End If
    If QPTrim$(fptxtRevOpt3.Text) <> "" Then
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt3.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt3Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrRevOpt3Owed.Value) - CDbl(fpCurrRevOpt3Adj.Value)))
    End If
  Else
    If InStr(fpcboAdjType.Text, "2") Then
      Print #RptHandle, Tab(5); "Adjustment Type:  Billing Adjustment Up"
      Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
      Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
      Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevTax.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPrincAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(PrincVal + PrincBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevInt.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrIntAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrIntOwed.Value) + CDbl(fpCurrIntAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtRecAdvCol.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrAdvColAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrAdvColOwed.Value) + CDbl(fpCurrAdvColAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevLateList.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrLateListAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrLateListOwed.Value) + CDbl(fpCurrLateListAdj.Value)))
      Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPenAdj.Value)); Tab(60); Using$("$##,##0.00", OldRound(CDbl(fpCurrPenOwed.Value) + CDbl(fpCurrPenAdj.Value)))
      If QPTrim$(fptxtRevOpt1.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt1.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt1Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrRevOpt1Owed.Value) + CDbl(fpCurrRevOpt1Adj.Value)))
      End If
      If QPTrim$(fptxtRevOpt2.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt2.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt2Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrRevOpt2Owed.Value) + CDbl(fpCurrRevOpt2Adj.Value)))
      End If
      If QPTrim$(fptxtRevOpt3.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt3.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt3Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(CDbl(fpCurrRevOpt3Owed.Value) + CDbl(fpCurrRevOpt3Adj.Value)))
      End If
    ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then '#3
      Print #RptHandle, Tab(5); "Adjustment Type:  Payment Adjustment"
      Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
      Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
      Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevTax.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrPrincAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(PrincVal + PrincBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevInt.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrIntAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(IntVal + IntBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtRecAdvCol.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrAdvColAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(AdvColVal + AdvColBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevLateList.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrLateListAdj.Value)); Tab(60); Using("$##,##0.00", OldRound(LateListVal + LateListBal))
      Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using$("$##,##0.00", CDbl(fpCurrPenAdj.Value)); Tab(60); Using$("$##,##0.00", OldRound(PenVal + PenBal))
      If QPTrim$(fptxtRevOpt1.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt1.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt1Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(Rev1Val + Rev1Bal))
      End If
      If QPTrim$(fptxtRevOpt2.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt2.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt2Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(Rev2Val + Rev2Bal))
      End If
      If QPTrim$(fptxtRevOpt3.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt3.Text); Tab(35); Using("$##,##0.00", CDbl(fpCurrRevOpt3Adj.Value)); Tab(60); Using("$##,##0.00", OldRound(Rev3Val + Rev3Bal))
      End If
    ElseIf fpcboAdjType.Text = "5-Prepay Adjust Down" Then '
      Print #RptHandle, Tab(5); "Adjustment Type:  Prepay Adjustment Down"
      Print #RptHandle, Tab(5); "Note: " + fptxtNote.Text
      Print #RptHandle, Tab(5); "Revenue Type"; Tab(35); "Adjustment Amount"; Tab(60); "New Balance"
      Print #RptHandle, Tab(5); "------------"; Tab(35); "-----------------"; Tab(60); "-----------"
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevTax.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevInt.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtRecAdvCol.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtRevLateList.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      Print #RptHandle, Tab(5); QPTrim$(fptxtPenalty.Text); Tab(35); Using$("$##,##0.00", 0); Tab(60); Using$("$##,##0.00", 0)
      If QPTrim$(fptxtRevOpt1.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt1.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      End If
      If QPTrim$(fptxtRevOpt2.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt2.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      End If
      If QPTrim$(fptxtRevOpt3.Text) <> "" Then
        Print #RptHandle, Tab(5); QPTrim$(fptxtRevOpt3.Text); Tab(35); Using("$##,##0.00", 0); Tab(60); Using("$##,##0.00", 0)
      End If
    End If
  End If
  Print #RptHandle, Tab(5); "----Account Balance Information----"
  Print #RptHandle,
  If InStr(fpcboAdjType.Text, "3") Then
    Print #RptHandle, Tab(5); "Balance Excluding This Bill: "; Tab(39); Using$("$###,##0.00", OldRound(GetCustRealBalance(GCustNum, -1) - NewBalThisBill))
    Print #RptHandle, Tab(5); "Previous Balance This Bill: "; Tab(39); Using("$###,##0.00", ThisBillBal#)
    If CDbl(fpCurrPrepayBal.Value) > 0 Then
      Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
      Print #RptHandle, Tab(5); "Prepaid Used: "; Tab(39); Using("$###,##0.00", Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))))
    Else
      Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
    End If
    Print #RptHandle, Tab(5); "Balance For This Bill: "; Tab(39); Using("$###,##0.00", NewBalThisBill)
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustRealBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  ElseIf InStr(fpcboAdjType.Text, "1") Or InStr(fpcboAdjType.Text, "4") Then
    Print #RptHandle, Tab(5); "Balance Excluding This Bill: "; Tab(39); Using$("$###,##0.00", OldRound(ThisGBal - CDbl(fpCurrTotOwed.Value)))
    Print #RptHandle, Tab(5); "Previous Balance This Bill: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotOwed.Value))
    Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
    Print #RptHandle, Tab(5); "Balance For This Bill: "; Tab(39); Using("$###,##0.00", OldRound(CDbl(fpCurrTotOwed.Value) - CDbl(fpCurrTotAdj.Value)))
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustRealBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  ElseIf InStr(fpcboAdjType.Text, "2") Then
    Print #RptHandle, Tab(5); "Balance Excluding This Bill: "; Tab(39); Using$("$###,##0.00", OldRound(GetCustRealBalance(GCustNum, -1) - NewBalThisBill))
    Print #RptHandle, Tab(5); "Previous Balance This Bill: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotOwed.Value))
    Print #RptHandle, Tab(5); "Current Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrTotAdj.Value))
    If CDbl(fpCurrPrepayBal.Value) > 0 Then
      Print #RptHandle, Tab(5); "Prepaid Used: "; Tab(39); Using("$###,##0.00", Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))))
    End If
    Print #RptHandle, Tab(5); "Balance For This Bill: "; Tab(39); Using("$###,##0.00", NewBalThisBill)
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustRealBalance(GCustNum, -1))
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(29); "Signature:__________________________________________"
    Print #RptHandle, FF$
  ElseIf InStr(fpcboAdjType.Text, "5") Then
    Print #RptHandle, Tab(5); "Previous Balance/Prepaid Balance: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrPrepayBal.Value))
    Print #RptHandle, Tab(5); "Prepaid Adjustment: "; Tab(39); Using("$###,##0.00", CDbl(fpCurrPrepayAdjAmt.Value))
    Print #RptHandle, Tab(5); "Prepaid Balance: "; Tab(39); Using("$###,##0.00", OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjAmt.Value)))
    Print #RptHandle, Tab(5); "Account Balance: "; Tab(39); Using("$###,##0.00", GetCustRealBalance(GCustNum, -1))
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "PrintText", Erl)
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
  Dim PrincPaid As Double
  Dim IntPaid As Double
  Dim AdvColPaid As Double
  Dim LateListPaid As Double
  Dim PenaltyPaid As Double
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
  
  PrincPaid = OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
  If CDbl(fpCurrPrincAdj.Value) > PrincPaid Then
    Call TaxMsg(800, "The total amount paid for PRINCIPLE is " + QPTrim$(Using$("$###,##0.00", PrincPaid)) + ". Adjustments for payments cannot exceed the total amount paid.")
    fpCurrPrincAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  IntPaid = TaxTrans.Revenue.InterestPd
  If CDbl(fpCurrIntAdj.Value) > IntPaid Then
    Call TaxMsg(800, "The total amount already paid for INTEREST is " + QPTrim$(Using$("$###,##0.00", IntPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrIntAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  AdvColPaid = TaxTrans.Revenue.CollectionPd
  If CDbl(fpCurrIntAdj.Value) > IntPaid Then
    Call TaxMsg(800, "The total amount already paid for ADV/COLLECT is " + QPTrim$(Using$("$###,##0.00", AdvColPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrAdvColAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  LateListPaid = TaxTrans.Revenue.LateListPd
  If CDbl(fpCurrLateListAdj.Value) > LateListPaid Then
    Call TaxMsg(800, "The total amount already paid for LATE LISTING is " + QPTrim$(Using$("$###,##0.00", LateListPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrLateListAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  PenaltyPaid = TaxTrans.Revenue.PenaltyPd
  If CDbl(fpCurrPenAdj.Value) > PenaltyPaid Then
    Call TaxMsg(800, "The total amount already paid for PENALTY is " + QPTrim$(Using$("$###,##0.00", PenaltyPaid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
    fpCurrPenAdj.SetFocus
    Close
    Check4PaidAmts = False
    Exit Function
  End If
  
  If fpCurrRevOpt1Adj.Enabled = True Then
    OptRev1Paid = TaxTrans.Revenue.RevOpt1Pd
    If CDbl(fpCurrRevOpt1Adj.Value) > OptRev1Paid Then
      Call TaxMsg(800, "The total amount already paid for " + QPTrim$(fptxtRevOpt1.Text) + " is " + QPTrim$(Using$("$###,##0.00", OptRev1Paid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
      fpCurrRevOpt1Adj.SetFocus
      Close
      Check4PaidAmts = False
      Exit Function
    End If
  End If
    
  If fpCurrRevOpt2Adj.Enabled = True Then
    OptRev2Paid = TaxTrans.Revenue.RevOpt2Pd
    If CDbl(fpCurrRevOpt2Adj.Value) > OptRev2Paid Then
      Call TaxMsg(800, "The total amount already paid for " + QPTrim$(fptxtRevOpt2.Text) + " is " + QPTrim$(Using$("$###,##0.00", OptRev2Paid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
      fpCurrRevOpt2Adj.SetFocus
      Close
      Check4PaidAmts = False
      Exit Function
    End If
  End If
    
  If fpCurrRevOpt3Adj.Enabled = True Then
    OptRev3Paid = TaxTrans.Revenue.RevOpt3Pd
    If CDbl(fpCurrRevOpt3Adj.Value) > OptRev3Paid Then
      Call TaxMsg(800, "The total amount already paid for " + QPTrim$(fptxtRevOpt3.Text) + " is " + QPTrim$(Using$("$###,##0.00", OptRev3Paid)) + ". Adjustments for payments cannot exceed the total amount already paid.")
      fpCurrRevOpt3Adj.SetFocus
      Close
      Check4PaidAmts = False
      Exit Function
    End If
  End If

  On Error GoTo ERRORSTUFF
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "Check4PaidAmts", Erl)
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
   
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrIntAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrIntAdj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrAdvColAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrAdvColAdj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrLateListAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrLateListAdj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPenAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPenAdj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPrincAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPrincAdj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrRevOpt1Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrRevOpt1Adj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrRevOpt2Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrRevOpt2Adj)
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrRevOpt3Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
      GoTo DoneHere
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrRevOpt3Adj)
    End If
DoneHere:
  ElseIf fpcboAdjType.Text = "3-Adjustment for Payment" Then
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrIntAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrIntAdj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrAdvColAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrAdvColAdj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrLateListAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrLateListAdj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPenAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPenAdj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrPrincAdj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrPrincAdj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrRevOpt1Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrRevOpt1Adj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrRevOpt2Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrRevOpt2Adj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
    
    If CDbl(fpCurrPrepayAdjBal.Value) > 0 And CDbl(fpCurrRevOpt3Adj.Value) >= CDbl(fpCurrPrepayAdjBal.Value) Then
      fpCurrPrepayAdjBal = 0
    Else
      fpCurrPrepayAdjBal = OldRound(fpCurrPrepayAdjBal - fpCurrRevOpt3Adj)
      If CDbl(fpCurrPrepayAdjBal.Value) < 0 Then fpCurrPrepayAdjBal = 0
    End If
  ElseIf fpcboAdjType.Text = "4-Release" Then
    fpCurrPrepayAdjBal = fpCurrTotAdj
  End If
 
End Sub

Private Sub ClearUserInput()
  fpCurrPrepayAdjBal = CDbl(fpCurrPrepayBal.Value)
  fpCurrPrepayAdjAmt = 0
  fpCurrPrincAdj = 0
  fpCurrIntAdj = 0
  fpCurrAdvColAdj = 0
  fpCurrLateListAdj = 0
  fpCurrPenAdj = 0
  fpCurrRevOpt1Adj = 0
  fpCurrRevOpt2Adj = 0
  fpCurrRevOpt3Adj = 0
  fpCurrTotAdj = 0
End Sub
'Private Function InPayBatchYN(CustRec As Long) As Boolean
'  Dim CitiPassFile As Integer, NumPassRecs As Integer
'  Dim CitiPass As CitiPassType
'  Dim x As Integer, y As Integer
'  Dim TaxPaymentRec As TaxPaymentRecType
'  Dim PHandle As Integer
'  Dim NumOfPRecs As Integer
'
'  If Len(Dir$("Citipass.dat")) Then
'    OpenCitiPassFile CitiPassFile, NumPassRecs
'    ReDim OPNums(1 To NumPassRecs) As Integer
'    ReDim OPNames(1 To NumPassRecs) As String
'    If Not CitiPassFile = -1 Then
'      For x = 1 To NumPassRecs
'        Get CitiPassFile, x, CitiPass
'        OPNums(x) = CitiPass.PassNum
'        OPNames(x) = QPTrim$(CitiPass.UserName)
'      Next x
'    End If
'  Else
'    Exit Function
'  End If
'  Close CitiPassFile
'
'  For x = 1 To NumPassRecs
'    If Exist("TAXCPR" + CStr(OPNums(x)) + ".DAT") Then
'      OpenTempPayFile PHandle, OPNums(x)
'      NumOfPRecs = LOF(PHandle) / Len(TaxPaymentRec)
'      For y = 1 To NumOfPRecs
'        Get PHandle, y, TaxPaymentRec
'        If TaxPaymentRec.CustAcct = CustRec Then
'          InPayBatchYN = True
'          Call TaxMsg(700, "This customer, " + QPTrim$(TaxPaymentRec.CustName) + ", is currently included in an unposted payment file for operator " + OPNames(x) + ". Please post this payment file before continuing with this adjustment.")
'          Close PHandle
'          Exit Function
'        End If
'      Next y
'      Close PHandle
'    End If
'  Next x
'
'  InPayBatchYN = False
'
'End Function
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
  ThisBal = GetCustRealBalance(GCustNum, -1)
  
  ReportFile$ = "TAXRPTS\TXADJRPT.RPT"
  
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
  Print #RptHandle, QPTrim$(fptxtRevTax.Text); dlm; QPTrim$(fptxtRevInt.Text); dlm;
  '                               7                                8
  Print #RptHandle, QPTrim$(fptxtRecAdvCol.Text); dlm; QPTrim$(fptxtRevLateList.Text); dlm;
  '
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    If QPTrim$(fptxtRevOpt1.Text) = "NOT IN USE" Then
      '                  9
      Print #RptHandle, ""; dlm;
    Else
      '                             9
      Print #RptHandle, QPTrim$(fptxtRevOpt1.Text); dlm;
    End If
    If QPTrim$(fptxtRevOpt2.Text) = "NOT IN USE" Then
      '                 10
      Print #RptHandle, ""; dlm;
    Else
      '                               10
      Print #RptHandle, QPTrim$(fptxtRevOpt2.Text); dlm;
    End If
    If QPTrim$(fptxtRevOpt3.Text) = "NOT IN USE" Then
      '                 11
      Print #RptHandle, ""; dlm;
    Else
      '                              11
      Print #RptHandle, QPTrim$(fptxtRevOpt3.Text); dlm;
    End If
  Else
    If fpCurrRevOpt1Adj.Enabled = True Then
      '                                 9
      Print #RptHandle, QPTrim$(fptxtRevOpt1.Text); dlm;
    Else
      '                 9
      Print #RptHandle, ""; dlm;
    End If
    If fpCurrRevOpt2Adj.Enabled = True Then
      '                           10
      Print #RptHandle, QPTrim$(fptxtRevOpt2.Text); dlm;
    Else
      '                 10
      Print #RptHandle, ""; dlm;
    End If
    If fpCurrRevOpt3Adj.Enabled = True Then
      '                               11
      Print #RptHandle, QPTrim$(fptxtRevOpt3.Text); dlm;
    Else
      '                 11
      Print #RptHandle, ""; dlm;
    End If
  End If
  If InStr(fpcboAdjType.Text, "1") Or InStr(fpcboAdjType.Text, "4") Then
    '                              12                             13
    Print #RptHandle, CDbl(fpCurrPrincAdj.Value); dlm; CDbl(fpCurrIntAdj.Value); dlm;
    '                              14                                 15
    Print #RptHandle, CDbl(fpCurrAdvColAdj.Value); dlm; CDbl(fpCurrLateListAdj.Value); dlm;
    '                              16                                 17
    Print #RptHandle, CDbl(fpCurrRevOpt1Adj.Value); dlm; CDbl(fpCurrRevOpt2Adj.Value); dlm;
    '                              18                               19                            20                      21
    Print #RptHandle, CDbl(fpCurrRevOpt3Adj.Value); dlm; CDbl(fpCurrTotOwed.Value); dlm; CDbl(fpCurrTotAdj.Value); dlm; ThisBal; dlm;
    If InStr(fpcboAdjType.Text, "1") Then
      '                     22
      Print #RptHandle, "Billing Downward Adjustment"; dlm;
    Else
      '                     22
      Print #RptHandle, "Release"; dlm;
    End If
    '                                          23
    Print #RptHandle, OldRound(CDbl(fpCurrPrincOwed.Value) - CDbl(fpCurrPrincAdj.Value)); dlm;
    '                                          24
    Print #RptHandle, OldRound(CDbl(fpCurrIntOwed.Value) - CDbl(fpCurrIntAdj.Value)); dlm;
    '                                          25
    Print #RptHandle, OldRound(CDbl(fpCurrAdvColOwed.Value) - CDbl(fpCurrAdvColAdj.Value)); dlm;
    '                                          26
    Print #RptHandle, OldRound(CDbl(fpCurrLateListOwed.Value) - CDbl(fpCurrLateListAdj.Value)); dlm;
    '                                          27
    Print #RptHandle, OldRound(CDbl(fpCurrRevOpt1Owed.Value) - CDbl(fpCurrRevOpt1Adj.Value)); dlm;
    '                                          28
    Print #RptHandle, OldRound(CDbl(fpCurrRevOpt2Owed.Value) - CDbl(fpCurrRevOpt2Adj.Value)); dlm;
    '                                          29
    Print #RptHandle, OldRound(CDbl(fpCurrRevOpt3Owed.Value) - CDbl(fpCurrRevOpt3Adj.Value)); dlm;
    '                                          30
    Print #RptHandle, OldRound(ThisGBal - CDbl(fpCurrTotOwed.Value)); dlm;
    '                       31                                            32                           33
    Print #RptHandle, fpLngIntBill.Text; dlm; OldRound(CDbl(fpCurrTotOwed) - CDbl(fpCurrTotAdj)); dlm; 0; dlm;
    '                       32                          33                                                34                                           35
    Print #RptHandle, fptxtPenalty.Text; dlm; CDbl(fpCurrPenAdj.Value); dlm; OldRound(CDbl(fpCurrPenOwed.Value) - CDbl(fpCurrPenAdj.Value)); dlm; fptxtNote.Text
  Else
    If InStr(fpcboAdjType.Text, "2") Then
      '                              12                             13
      Print #RptHandle, CDbl(fpCurrPrincAdj.Value); dlm; CDbl(fpCurrIntAdj.Value); dlm;
      '                              14                                 15
      Print #RptHandle, CDbl(fpCurrAdvColAdj.Value); dlm; CDbl(fpCurrLateListAdj.Value); dlm;
      '                              16                                 17
      Print #RptHandle, CDbl(fpCurrRevOpt1Adj.Value); dlm; CDbl(fpCurrRevOpt2Adj.Value); dlm;
      '                              18                               19                            20                      21
      Print #RptHandle, CDbl(fpCurrRevOpt3Adj.Value); dlm; CDbl(fpCurrTotOwed.Value); dlm; CDbl(fpCurrTotAdj.Value); dlm; ThisBal; dlm;
      '                             22
      Print #RptHandle, "Billing Upward Adjustment"; dlm;
      '                                23
      Print #RptHandle, OldRound(PrincBal + PrincVal); dlm;
      '                                24
      Print #RptHandle, OldRound(IntBal + IntVal); dlm;
      '                                25
      Print #RptHandle, OldRound(AdvColBal + AdvColVal); dlm;
      '                                26
      Print #RptHandle, OldRound(LateListBal + LateListVal); dlm;
      '                                27
      Print #RptHandle, OldRound(Rev1Bal + Rev1Val); dlm;
      '                                28
      Print #RptHandle, OldRound(Rev2Bal + Rev2Val); dlm;
      '                                29
      Print #RptHandle, OldRound(Rev3Bal + Rev3Val); dlm;
      '                                        30
      Print #RptHandle, OldRound(GetCustRealBalance(GCustNum, -1) - NewBalThisBill); dlm;
      '                       31                      32                                              33
      Print #RptHandle, fpLngIntBill.Text; dlm; NewBalThisBill; dlm; Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))); dlm;
      '                       32                          33                                                34                                          35
      Print #RptHandle, fptxtPenalty.Text; dlm; CDbl(fpCurrPenAdj.Value); dlm; OldRound(CDbl(fpCurrPenOwed.Value) - CDbl(fpCurrPenAdj.Value)); dlm; fptxtNote.Text
    ElseIf InStr(fpcboAdjType.Text, "3") Then
      '                              12                             13
      Print #RptHandle, CDbl(fpCurrPrincAdj.Value); dlm; CDbl(fpCurrIntAdj.Value); dlm;
      '                              14                                 15
      Print #RptHandle, CDbl(fpCurrAdvColAdj.Value); dlm; CDbl(fpCurrLateListAdj.Value); dlm;
      '                              16                                 17
      Print #RptHandle, CDbl(fpCurrRevOpt1Adj.Value); dlm; CDbl(fpCurrRevOpt2Adj.Value); dlm;
      '                              18                         19                       20                   21
      Print #RptHandle, CDbl(fpCurrRevOpt3Adj.Value); dlm; ThisBillBal; dlm; CDbl(fpCurrTotAdj.Value); dlm; ThisBal; dlm;
      '                         22
      Print #RptHandle, "Payment Adjustment"; dlm;
      '                                23
      Print #RptHandle, OldRound(PrincVal + PrincBal); dlm;
      '                                24
      Print #RptHandle, OldRound(IntVal + IntBal); dlm;
      '                                25
      Print #RptHandle, OldRound(AdvColVal + AdvColBal); dlm;
      '                                26
      Print #RptHandle, OldRound(LateListVal + LateListBal); dlm;
      '                                27
      Print #RptHandle, OldRound(Rev1Val + Rev1Bal); dlm;
      '                                28
      Print #RptHandle, OldRound(Rev2Val + Rev2Bal); dlm;
      '                                29
      Print #RptHandle, OldRound(Rev3Val + Rev3Bal); dlm;
      '                                            30
      Print #RptHandle, OldRound(GetCustRealBalance(GCustNum, -1) - NewBalThisBill); dlm;
      '                       31                      32                                               33
      Print #RptHandle, fpLngIntBill.Text; dlm; NewBalThisBill; dlm; Abs(OldRound(CDbl(fpCurrPrepayBal.Value) - CDbl(fpCurrPrepayAdjBal.Value))); dlm;
'      '                       32                          33                                                34                                           35
'      Print #RptHandle, fptxtPenalty.Text; dlm; CDbl(fpCurrPenAdj.Value); dlm; OldRound(CDbl(fpCurrPenOwed.Value) - CDbl(fpCurrPenAdj.Value)); dlm; fptxtNote.Text
      '                       32                          33                                 34                      35
      Print #RptHandle, fptxtPenalty.Text; dlm; CDbl(fpCurrPenAdj.Value); dlm; OldRound(PenBal + PenVal); dlm; fptxtNote.Text 'changed on 5.24.07
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
      '                   30           31             32
      Print #RptHandle, ThisAmt; dlm; "NA"; dlm; fptxtNote.Text
    
    End If
  End If
  Close
  
  If fpcboAdjType.Text = "5-Prepay Adjust Down" Then
    arVATaxAdjOPOnlyReport.Show
  Else
    arVATaxAdjRpt.Show
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxAdjustments", "PrintGraphics", Erl)
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
  Dim Disc6 As Double '1/25/2007
  Dim Disc7 As Double '1/25/2007
  Dim Disc8 As Double '1/25/2007
  Dim Disc9 As Double '9/19/07
  Dim Dif As Double '9/19/07
  Dim SaveAmt As Double '1/25/2007
  
  Disc1 = 0
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
  Disc6 = TaxRec.Revenue.RevOpt1Pd / SaveAmt
  Disc6 = Disc6 * TaxRec.DiscAmt
  Disc7 = TaxRec.Revenue.RevOpt2Pd / SaveAmt
  Disc7 = Disc7 * TaxRec.DiscAmt
  Disc8 = TaxRec.Revenue.RevOpt3Pd / SaveAmt
  Disc9 = TaxRec.Revenue.LateListPd / SaveAmt
  Disc9 = Disc9 * TaxRec.DiscAmt
  
  TaxRec.Revenue.Principle1Pd = OldRound(TaxRec.Revenue.Principle1Pd + Disc1 + Disc9)
  TaxRec.Revenue.RevOpt1Pd = OldRound(TaxRec.Revenue.RevOpt1Pd + Disc6)
  TaxRec.Revenue.RevOpt2Pd = OldRound(TaxRec.Revenue.RevOpt2Pd + Disc7)
  TaxRec.Revenue.RevOpt3Pd = OldRound(TaxRec.Revenue.RevOpt3Pd + Disc8)
'  Dif = OldRound(Disc1 + Disc6 + Disc7 + Disc8 + Disc9)'remmed out on 11/14/07
'  If Dif <> 0 Then
'    If Disc1 > 0 Or Disc9 > 0 Then
'      TaxRec.Revenue.Principle1Pd = OldRound(TaxRec.Revenue.Principle1Pd + Dif)
'    ElseIf Disc6 > 0 Then
'      TaxRec.Revenue.RevOpt1Pd = OldRound(TaxRec.Revenue.RevOpt1Pd + Dif)
'    ElseIf Disc7 > 0 Then
'      TaxRec.Revenue.RevOpt2Pd = OldRound(TaxRec.Revenue.RevOpt2Pd + Dif)
'    ElseIf Disc8 > 0 Then
'      TaxRec.Revenue.RevOpt3Pd = OldRound(TaxRec.Revenue.RevOpt3Pd + Dif)
'    End If
'  End If
  
End Sub

