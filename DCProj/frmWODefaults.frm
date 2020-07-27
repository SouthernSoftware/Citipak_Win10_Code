VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmWODefaults 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Default Work Orders"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmWODefaults.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboWOList 
      Height          =   384
      Left            =   5040
      TabIndex        =   0
      Top             =   1512
      Width           =   4404
      _Version        =   196608
      _ExtentX        =   7768
      _ExtentY        =   677
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
      ColDesigner     =   "frmWODefaults.frx":08CA
   End
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   0
      Left            =   2112
      TabIndex        =   2
      Top             =   3336
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   1
      Left            =   2112
      TabIndex        =   3
      Top             =   3612
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   2
      Left            =   2112
      TabIndex        =   4
      Top             =   3888
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   3
      Left            =   2112
      TabIndex        =   5
      Top             =   4164
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   4
      Left            =   2112
      TabIndex        =   6
      Top             =   4440
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   5
      Left            =   2112
      TabIndex        =   7
      Top             =   4716
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWORem 
      Height          =   276
      Index           =   0
      Left            =   2112
      TabIndex        =   8
      Top             =   5448
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWORem 
      Height          =   276
      Index           =   1
      Left            =   2112
      TabIndex        =   9
      Top             =   5720
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWORem 
      Height          =   276
      Index           =   2
      Left            =   2112
      TabIndex        =   10
      Top             =   5992
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWORem 
      Height          =   276
      Index           =   3
      Left            =   2112
      TabIndex        =   11
      Top             =   6264
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWORem 
      Height          =   276
      Index           =   4
      Left            =   2112
      TabIndex        =   12
      Top             =   6536
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fptxtWORem 
      Height          =   276
      Index           =   5
      Left            =   2112
      TabIndex        =   13
      Top             =   6804
      Width           =   8244
      _Version        =   196608
      _ExtentX        =   14541
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
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
      MaxLength       =   67
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
   Begin EditLib.fpText fpWOTitle 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3576
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2328
      Width           =   4764
      _Version        =   196608
      _ExtentX        =   8403
      _ExtentY        =   572
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPrint 
      Height          =   384
      Left            =   7548
      TabIndex        =   14
      Top             =   7752
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmWODefaults.frx":0D69
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   384
      Left            =   10320
      TabIndex        =   15
      Top             =   7752
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmWODefaults.frx":0F45
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   8934
      TabIndex        =   16
      Top             =   7752
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmWODefaults.frx":1121
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   17
      Top             =   8568
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   529
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
            TextSave        =   "2:52 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4/26/2005"
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
   Begin EditLib.fpLongInteger fpRecNo 
      Height          =   300
      Left            =   2088
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   672
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BorderStyle     =   1
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdDelete 
      Height          =   384
      Left            =   6168
      TabIndex        =   37
      Top             =   7752
      Width           =   1224
      _Version        =   131072
      _ExtentX        =   2159
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmWODefaults.frx":12FD
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "or Select Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1632
      TabIndex        =   36
      Top             =   1752
      Width           =   3204
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Work Order to Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1632
      TabIndex        =   35
      Top             =   1464
      Width           =   3204
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   5292
      Left            =   1344
      Top             =   2112
      Width           =   9540
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "6)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   34
      Top             =   4704
      Width           =   276
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   33
      Top             =   4428
      Width           =   276
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "4)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   32
      Top             =   4152
      Width           =   276
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   31
      Top             =   3876
      Width           =   276
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   30
      Top             =   3588
      Width           =   276
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   29
      Top             =   3312
      Width           =   276
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   1704
      X2              =   10464
      Y1              =   2784
      Y2              =   2784
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   2784
      X2              =   10392
      Y1              =   5232
      Y2              =   5232
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   4344
      X2              =   10392
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Labe54 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1752
      TabIndex        =   28
      Top             =   2976
      Width           =   2652
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1464
      TabIndex        =   27
      Top             =   2352
      Width           =   2028
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1776
      TabIndex        =   26
      Top             =   5088
      Width           =   1212
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "6)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   0
      Left            =   1848
      TabIndex        =   25
      Top             =   6792
      Width           =   276
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   24
      Top             =   6516
      Width           =   276
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "4)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   23
      Top             =   6240
      Width           =   276
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   22
      Top             =   5964
      Width           =   276
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   21
      Top             =   5676
      Width           =   276
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   1848
      TabIndex        =   20
      Top             =   5400
      Width           =   276
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order Setup"
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
      Index           =   1
      Left            =   3888
      TabIndex        =   19
      Top             =   648
      Width           =   4452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3228
      Top             =   480
      Width           =   5772
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3228
      Top             =   360
      Width           =   5772
   End
End
Attribute VB_Name = "frmWODefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim RecNo As Long, CntL As Long
Dim UBSetUpRec(1) As UBSetupRecType
Dim UBSetupLen As Integer
Dim BeenDone As Boolean
Dim BtnFnt As Double
Dim EditFlag As Boolean, AddingFlag As Boolean

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fpcboWOList.InsertRow = " " & Chr$(9) & "ADD NEW WORK ORDER DEFAULT"
  GetWOList fpcboWOList
  fpcboWOList.ListIndex = 0
  DoEvents
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
   ' DoEvents
    Temp_Class.ResizeControls Me
   ' DoEvents
   ' Me.Visible = True
   ' Me.AutoRedraw = False
   ' DoEvents
  End If
  DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via WODefaults by " + PWUser$
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
    Case vbKeyF4:
      KeyCode = 0
      fpCmdDelete_Click
    Case vbKeyF5:        'Print
      KeyCode = 0
      fpCmdPrint_Click
    Case vbKeyEscape:
      fpCmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      KeyCode = 0
      fpCmdSave_Click
    Case Else:
  End Select
End Sub

Private Sub fpCmdDelete_Click()
  If fpcboWOList.ListIndex <> 0 Then
    DeleteWO
  End If
    fpWOTitle = ""
    fptxtWOInf(0).Text = ""
    fptxtWOInf(1).Text = ""
    fptxtWOInf(2).Text = ""
    fptxtWOInf(3).Text = ""
    fptxtWOInf(4).Text = ""
    fptxtWOInf(5).Text = ""
    fptxtWORem(0).Text = ""
    fptxtWORem(1).Text = ""
    fptxtWORem(2).Text = ""
    fptxtWORem(3).Text = ""
    fptxtWORem(4).Text = ""
    fptxtWORem(5).Text = ""

End Sub

Private Sub fpCmdExit_Click()
  ExitWorkOrder
End Sub
Private Sub ExitWorkOrder()
  If Chk4Change = True Then
    If MsgBox("Would you like to abandon changes?", vbYesNo, "Abandon Changes?") = vbNo Then
      Exit Sub
    End If
  End If
    UBLog "OUT: Work Order Default."
    RecNo = 0
    BeenDone = False
    Load frmUBSetupMenu
    DoEvents
    frmUBSetupMenu.Show
    Unload frmWODefaults
End Sub

'Only load if existing record otherwise rec will be totrecs + 1
Private Sub LoadInfo2Form()
  Dim WorkOrderDefLen As Integer, NumWOs As Long
  Dim UBWrkOrdD As Integer

  ReDim WorkOrderDef(1) As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef(1))

  UBWrkOrdD = FreeFile
  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
  NumWOs = LOF(UBWrkOrdD) \ WorkOrderDefLen
  UBLog " IN: Work Order Setup."
  RecNo& = Val(fpRecNo)
  fpRecNo = 0
  If Not RecNo& > 0 Then
    RecNo& = NumWOs + 1
    EditFlag = False
    AddingFlag = True
    fpWOTitle = ""
    fptxtWOInf(0).Text = ""
    fptxtWOInf(1).Text = ""
    fptxtWOInf(2).Text = ""
    fptxtWOInf(3).Text = ""
    fptxtWOInf(4).Text = ""
    fptxtWOInf(5).Text = ""
    fptxtWORem(0).Text = ""
    fptxtWORem(1).Text = ""
    fptxtWORem(2).Text = ""
    fptxtWORem(3).Text = ""
    fptxtWORem(4).Text = ""
    fptxtWORem(5).Text = ""
  Else
    Get UBWrkOrdD, RecNo&, WorkOrderDef(1)
    EditFlag = True
    AddingFlag = False
    fpWOTitle = QPTrim(WorkOrderDef(1).WOType)
    fptxtWOInf(0).Text = QPTrim(WorkOrderDef(1).OrdersText.Text(1))
    fptxtWOInf(1).Text = QPTrim(WorkOrderDef(1).OrdersText.Text(2))
    fptxtWOInf(2).Text = QPTrim(WorkOrderDef(1).OrdersText.Text(3))
    fptxtWOInf(3).Text = QPTrim(WorkOrderDef(1).OrdersText.Text(4))
    fptxtWOInf(4).Text = QPTrim(WorkOrderDef(1).OrdersText.Text(5))
    fptxtWOInf(5).Text = QPTrim(WorkOrderDef(1).OrdersText.Text(6))
    fptxtWORem(0).Text = QPTrim(WorkOrderDef(1).RepliesText.Text(1))
    fptxtWORem(1).Text = QPTrim(WorkOrderDef(1).RepliesText.Text(2))
    fptxtWORem(2).Text = QPTrim(WorkOrderDef(1).RepliesText.Text(3))
    fptxtWORem(3).Text = QPTrim(WorkOrderDef(1).RepliesText.Text(4))
    fptxtWORem(4).Text = QPTrim(WorkOrderDef(1).RepliesText.Text(5))
    fptxtWORem(5).Text = QPTrim(WorkOrderDef(1).RepliesText.Text(6))
  End If
  Close
End Sub
Private Sub fpcboWOList_Click()
  fpcboWOList.col = 0
  fpRecNo = fpcboWOList.ColText
  LoadInfo2Form
End Sub
'
'Need to figure out what to do about recno if new one !!!!
'on save!!!!
Private Sub SaveWorkOrderDef()
  Dim WorkOrderDefLen As Integer
  Dim UBWrkOrdD As Integer, NumWOs As Long
  ReDim WorkOrderDef(1) As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef(1))

  UBWrkOrdD = FreeFile
  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
  NumWOs = LOF(UBWrkOrdD) \ WorkOrderDefLen
  WorkOrderDef(1).Deleted = False
  WorkOrderDef(1).WOType = QPTrim(fpWOTitle.Text)
  WorkOrderDef(1).OrdersText.Text(1) = QPTrim(fptxtWOInf(0).Text)
  WorkOrderDef(1).OrdersText.Text(2) = QPTrim(fptxtWOInf(1).Text)
  WorkOrderDef(1).OrdersText.Text(3) = QPTrim(fptxtWOInf(2).Text)
  WorkOrderDef(1).OrdersText.Text(4) = QPTrim(fptxtWOInf(3).Text)
  WorkOrderDef(1).OrdersText.Text(5) = QPTrim(fptxtWOInf(4).Text)
  WorkOrderDef(1).OrdersText.Text(6) = QPTrim(fptxtWOInf(5).Text)
  WorkOrderDef(1).RepliesText.Text(1) = QPTrim(fptxtWORem(0).Text)
  WorkOrderDef(1).RepliesText.Text(2) = QPTrim(fptxtWORem(1).Text)
  WorkOrderDef(1).RepliesText.Text(3) = QPTrim(fptxtWORem(2).Text)
  WorkOrderDef(1).RepliesText.Text(4) = QPTrim(fptxtWORem(3).Text)
  WorkOrderDef(1).RepliesText.Text(5) = QPTrim(fptxtWORem(4).Text)
  WorkOrderDef(1).RepliesText.Text(6) = QPTrim(fptxtWORem(5).Text)


  Put UBWrkOrdD, RecNo&, WorkOrderDef(1)

  Close
UBLog " Save: Work Order Default " + Str(RecNo&)
End Sub

Private Sub fpCmdPrint_Click()
'send sample to printer
  frmReportOpt.Show 1
  DeActivateControls Me
  If rptopt = 1 Then
  'do graphic report
    PrintWOSample True
  ElseIf rptopt = 2 Then
  'do text report
    PrintWOSample False
  End If
  ActivateControls Me

End Sub

Private Sub fpCmdSave_Click()
'do checking fields here then if ok do save
  If Len(fpWOTitle.Text) > 0 Then
    If RecNo > 0 Then
      SaveWorkOrderDef
      MsgBox "Work Order Default Saved.", vbOKOnly, "Completed"
      setscreen
    ' LoadInfo2Form
    End If
  Else
   MsgBox "Please fill out the Work Order Title before saving.", vbOKOnly, "Invalid Information"
  End If
End Sub
Private Sub setscreen()
  fpcboWOList.Clear
  fpcboWOList.InsertRow = " " & Chr$(9) & "ADD NEW WORK ORDER DEFAULT"
  GetWOList fpcboWOList
  fpcboWOList.ListIndex = 0
  fpWOTitle.SetFocus
End Sub

Private Sub fpcboWOList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboWOList.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboWOList.ListIndex = -1
    fpcboWOList.Action = ActionClearSearchBuffer
  End If
  If fpcboWOList.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpWOTitle.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtWOInf_ChangeMode(Index As Integer, EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtWOInf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Dim x As Integer
 x = Index
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    If Index < 5 Then
     For x = Index To 4
      If fptxtWOInf(x + 1).Enabled Then
        fptxtWOInf(x + 1).SetFocus
        Exit For
      End If
     Next
    End If
    If Index = 5 Then
       fptxtWORem(0).SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If Index > 0 Then
     For x = Index To 5
      If fptxtWOInf(x - 1).Enabled Then
        fptxtWOInf(x - 1).SetFocus
        Exit For
      End If
     Next
    End If
    If Index = 0 Then
      fpWOTitle.SetFocus
    End If
  End If
End Sub
Private Sub fptxtWORem_ChangeMode(Index As Integer, EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtWORem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Dim x As Integer
 x = Index
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    If Index < 5 Then
     For x = Index To 4
      If fptxtWORem(x + 1).Enabled Then
        fptxtWORem(x + 1).SetFocus
        Exit For
      End If
     Next
    End If
    If Index = 5 Then
       fpCmdSave.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If Index > 0 Then
     For x = Index To 5
      If fptxtWORem(x - 1).Enabled Then
        fptxtWORem(x - 1).SetFocus
        Exit For
      End If
     Next
    End If
    If Index = 0 Then
      fptxtWOInf(5).SetFocus
    End If
  End If
End Sub
Private Sub fpWOTitle_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpWOTitle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    fptxtWOInf(0).SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpcboWOList.SetFocus
  End If
End Sub
Private Function Chk4Change()
  Dim WorkOrderDefLen As Integer, NumWOs As Long
  Dim UBWrkOrdD As Integer

  ReDim WorkOrderDef(1) As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef(1))
  Chk4Change = False
  If fpcboWOList.ListIndex = 0 Then
    
    If Not fpWOTitle = "" Then Chk4Change = True
    If Not fptxtWOInf(0).Text = "" Then Chk4Change = True
    If Not fptxtWOInf(1).Text = "" Then Chk4Change = True
    If Not fptxtWOInf(2).Text = "" Then Chk4Change = True
    If Not fptxtWOInf(3).Text = "" Then Chk4Change = True
    If Not fptxtWOInf(4).Text = "" Then Chk4Change = True
    If Not fptxtWOInf(5).Text = "" Then Chk4Change = True
    If Not fptxtWORem(0).Text = "" Then Chk4Change = True
    If Not fptxtWORem(1).Text = "" Then Chk4Change = True
    If Not fptxtWORem(2).Text = "" Then Chk4Change = True
    If Not fptxtWORem(3).Text = "" Then Chk4Change = True
    If Not fptxtWORem(4).Text = "" Then Chk4Change = True
    If Not fptxtWORem(5).Text = "" Then Chk4Change = True
  Else
    UBWrkOrdD = FreeFile
    Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
    NumWOs = LOF(UBWrkOrdD) \ WorkOrderDefLen
    Get UBWrkOrdD, RecNo&, WorkOrderDef(1)
    If Not QPTrim(fpWOTitle) = QPTrim(WorkOrderDef(1).WOType) Then Chk4Change = True
    If Not QPTrim(fptxtWOInf(0).Text) = QPTrim(WorkOrderDef(1).OrdersText.Text(1)) Then Chk4Change = True
    If Not QPTrim(fptxtWOInf(1).Text) = QPTrim(WorkOrderDef(1).OrdersText.Text(2)) Then Chk4Change = True
    If Not QPTrim(fptxtWOInf(2).Text) = QPTrim(WorkOrderDef(1).OrdersText.Text(3)) Then Chk4Change = True
    If Not QPTrim(fptxtWOInf(3).Text) = QPTrim(WorkOrderDef(1).OrdersText.Text(4)) Then Chk4Change = True
    If Not QPTrim(fptxtWOInf(4).Text) = QPTrim(WorkOrderDef(1).OrdersText.Text(5)) Then Chk4Change = True
    If Not QPTrim(fptxtWOInf(5).Text) = QPTrim(WorkOrderDef(1).OrdersText.Text(6)) Then Chk4Change = True
    If Not QPTrim(fptxtWORem(0).Text) = QPTrim(WorkOrderDef(1).RepliesText.Text(1)) Then Chk4Change = True
    If Not QPTrim(fptxtWORem(1).Text) = QPTrim(WorkOrderDef(1).RepliesText.Text(2)) Then Chk4Change = True
    If Not QPTrim(fptxtWORem(2).Text) = QPTrim(WorkOrderDef(1).RepliesText.Text(3)) Then Chk4Change = True
    If Not QPTrim(fptxtWORem(3).Text) = QPTrim(WorkOrderDef(1).RepliesText.Text(4)) Then Chk4Change = True
    If Not QPTrim(fptxtWORem(4).Text) = QPTrim(WorkOrderDef(1).RepliesText.Text(5)) Then Chk4Change = True
    If Not QPTrim(fptxtWORem(5).Text) = QPTrim(WorkOrderDef(1).RepliesText.Text(6)) Then Chk4Change = True
  End If
  Close
End Function
Public Sub DeleteWO()
  Dim WorkOrderDefLen As Integer
  Dim UBWrkOrdD As Integer
  ReDim WorkOrderDef(1) As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef(1))

  UBWrkOrdD = FreeFile
  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
  WorkOrderDef(1).Deleted = True
  Put UBWrkOrdD, RecNo&, WorkOrderDef(1)

  Close
  sortwos
  UBLog " Delete: Work Order Default " + Str(RecNo&)
  MsgBox "Work Order Default Deleted.", vbOKOnly, "Deleted"
  setscreen
End Sub
Private Sub sortwos()
  Dim cnt As Long, NumWOs As Long, newcnt As Long
  Dim WorkOrderDefLen As Integer, WorkOrdertLen As Integer
  Dim UBWrkOrdD As Integer, UBWrkOrdT As Integer
  Dim WorkOrderT As WorkOrderDefType
  Dim WorkOrderDef As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef)
  KillFile UBPath$ + "UBWODef.old"
  SH_Rename (UBPath$ + "UBWODef.DAT"), (UBPath$ + "UBWODef.old")
  UBWrkOrdD = FreeFile
  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
  WorkOrdertLen = Len(WorkOrderT)
  UBWrkOrdT = FreeFile
  Open UBPath$ + "UBWODef.old" For Random Shared As UBWrkOrdT Len = WorkOrdertLen
  NumWOs = LOF(UBWrkOrdT) \ WorkOrdertLen
  newcnt = 0
  For cnt = 1 To NumWOs
    Get UBWrkOrdT, cnt, WorkOrderT
      If WorkOrderT.Deleted = False Then
        newcnt = newcnt + 1
        WorkOrderDef.Deleted = False
        WorkOrderDef.WOType = WorkOrderT.WOType
        WorkOrderDef.OrdersText.Text(1) = WorkOrderT.OrdersText.Text(1)
        WorkOrderDef.OrdersText.Text(2) = WorkOrderT.OrdersText.Text(2)
        WorkOrderDef.OrdersText.Text(3) = WorkOrderT.OrdersText.Text(3)
        WorkOrderDef.OrdersText.Text(4) = WorkOrderT.OrdersText.Text(4)
        WorkOrderDef.OrdersText.Text(5) = WorkOrderT.OrdersText.Text(5)
        WorkOrderDef.OrdersText.Text(6) = WorkOrderT.OrdersText.Text(6)
        WorkOrderDef.RepliesText.Text(1) = WorkOrderT.RepliesText.Text(1)
        WorkOrderDef.RepliesText.Text(2) = WorkOrderT.RepliesText.Text(2)
        WorkOrderDef.RepliesText.Text(3) = WorkOrderT.RepliesText.Text(3)
        WorkOrderDef.RepliesText.Text(4) = WorkOrderT.RepliesText.Text(4)
        WorkOrderDef.RepliesText.Text(5) = WorkOrderT.RepliesText.Text(5)
        WorkOrderDef.RepliesText.Text(6) = WorkOrderT.RepliesText.Text(6)
        Put UBWrkOrdD, newcnt&, WorkOrderDef
      End If
  Next
  Close
End Sub
Private Sub PrintWOSample(graphicflag As Boolean)
  Dim Dash As String, Handle As Integer, cnt As Long
  Dim ReportFile As String, RptHandle As Integer, ToPrint As String
  Dim Header As String, MtrCnt As Integer, IdxName As String
  Dim Rem1 As String, Rem2 As String, Rem3 As String, Rem4 As String
  Dim Rem5 As String, Rem6 As String
  Rem1$ = ""
  Rem2$ = ""
  Rem3$ = ""
  Rem4$ = ""
  Rem5$ = ""
  Rem6$ = ""
  If graphicflag = True Then
    Dash$ = String$(83, "_")
  Else
    Dash$ = String$(79, "_")
  End If
  ToPrint$ = ""
  FF$ = Chr$(12)

  ReportFile$ = UBPath$ + "WORKORDR.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  If Len(QPTrim(fptxtWORem(0).Text)) > 0 Then
    Rem1$ = QPTrim(fptxtWORem(0).Text)
  Else
    Rem1$ = Dash$
  End If
  If Len(QPTrim(fptxtWORem(1).Text)) > 0 Then
    Rem2$ = QPTrim(fptxtWORem(1).Text)
  Else
    Rem2$ = Dash$
  End If
  If Len(QPTrim(fptxtWORem(2).Text)) > 0 Then
    Rem3$ = QPTrim(fptxtWORem(2).Text)
  Else
    Rem3$ = Dash$
  End If
  If Len(QPTrim(fptxtWORem(3).Text)) > 0 Then
    Rem4$ = QPTrim(fptxtWORem(3).Text)
  Else
    Rem4$ = Dash$
  End If
  If Len(QPTrim(fptxtWORem(4).Text)) > 0 Then
    Rem5$ = QPTrim(fptxtWORem(4).Text)
  Else
    Rem5$ = Dash$
  End If

  If Len(QPTrim(fptxtWORem(5).Text)) > 0 Then
    Rem6$ = QPTrim(fptxtWORem(5).Text)
  Else
    Rem6$ = "BY: ______________________________   DATE: ____________________"
  End If
 
  
  If graphicflag = False Then
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, Tab(14); "W O R K   O R D E R   :   U T I L I T Y   D E P T ."
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "    Work Order#: "; Using("######", 999999); Tab(30); "Date Issued: "; Date$
    Print #RptHandle, "      Location#: "; "99-999999"; Tab(30); "Complete By: "; Date$
    Print #RptHandle, "       Account#: "; Str(99999); Tab(30); "  Completed: "; Date$
    Print #RptHandle, "  Customer Name: "; "John Smith"
    Print #RptHandle, "Service Address: "; "105 North Main Street"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Instruction or Description of Work Needed"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, fptxtWOInf(0).Text
    Print #RptHandle, fptxtWOInf(1).Text
    Print #RptHandle, fptxtWOInf(2).Text
    Print #RptHandle, fptxtWOInf(3).Text
    Print #RptHandle, fptxtWOInf(4).Text
    Print #RptHandle, fptxtWOInf(5).Text
    Print #RptHandle, " "
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Remarks Noted by Worker"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, Rem1$
    Print #RptHandle, " "
    Print #RptHandle, Rem2$
    Print #RptHandle, " "
    Print #RptHandle, Rem3$
    Print #RptHandle, " "
    Print #RptHandle, Rem4$
    Print #RptHandle, " "
    Print #RptHandle, Rem5$
    Print #RptHandle, " "
    Print #RptHandle, Rem6$
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "Meter Numbers:"

    For MtrCnt = 1 To 7
      'If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #RptHandle, "0123456733A"
      'End If
    Next
    Print #RptHandle, FF$;
  Else
    ToPrint$ = Date$ + "~"
    ToPrint$ = ToPrint$ + Using("######", 999999) + "~"
    ToPrint$ = ToPrint$ + "99-999999" + "~"
    ToPrint$ = ToPrint$ + "99999" + "~"
    ToPrint$ = ToPrint$ + "John Smith" + "~"
    ToPrint$ = ToPrint$ + "105 North Main Street" + "~"
    ToPrint$ = ToPrint$ + fptxtWOInf(0).Text + "~"
    ToPrint$ = ToPrint$ + fptxtWOInf(1).Text + "~"
    ToPrint$ = ToPrint$ + fptxtWOInf(2).Text + "~"
    ToPrint$ = ToPrint$ + fptxtWOInf(3).Text + "~"
    ToPrint$ = ToPrint$ + fptxtWOInf(4).Text + "~"
    ToPrint$ = ToPrint$ + fptxtWOInf(5).Text + "~"
    ToPrint$ = ToPrint$ + Rem1$ + "~"
    ToPrint$ = ToPrint$ + Rem2$ + "~"
    ToPrint$ = ToPrint$ + Rem3$ + "~"
    ToPrint$ = ToPrint$ + Rem4$ + "~"
    ToPrint$ = ToPrint$ + Rem5$ + "~"
    ToPrint$ = ToPrint$ + Rem6$

    For MtrCnt = 1 To 7
        ToPrint$ = ToPrint$ + "~" + "0123456789A"
    Next
    ToPrint$ = ToPrint$ + "~" + Date$ + "~"
    ToPrint$ = ToPrint$ + Date$

    Print #RptHandle, ToPrint$
    ToPrint$ = ""
  End If
  Close
  Header$ = "Customer Work Order"
  'PrintRptFile Header$, ReportFile$, LPTPort, RetCode, EntryPoint
  If graphicflag = False Then
    ViewPrint ReportFile$, Header$
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmWODefaults
    ARptWorkOrder.GetName ReportFile$
    ARptWorkOrder.startrpt
  End If

End Sub
