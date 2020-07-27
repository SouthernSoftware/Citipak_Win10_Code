VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmWorkOrderEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work Order Entry"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmWorkOrderEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboWOList 
      Height          =   360
      Left            =   4368
      TabIndex        =   0
      Top             =   1464
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      ColDesigner     =   "frmWorkOrderEntry.frx":08CA
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   816
      Top             =   432
   End
   Begin fpBtnAtlLibCtl.fpBtn fpLoadDefault 
      Height          =   372
      Left            =   9000
      TabIndex        =   1
      Top             =   1464
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":0CF9
   End
   Begin EditLib.fpText fptxtWOInf 
      Height          =   276
      Index           =   0
      Left            =   2112
      TabIndex        =   3
      Top             =   3144
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
      TabIndex        =   4
      Top             =   3420
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
      TabIndex        =   5
      Top             =   3696
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
      TabIndex        =   6
      Top             =   3972
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
      TabIndex        =   7
      Top             =   4248
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
      TabIndex        =   8
      Top             =   4524
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
      TabIndex        =   9
      Top             =   5088
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
      TabIndex        =   10
      Top             =   5352
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
      TabIndex        =   11
      Top             =   5616
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
      TabIndex        =   12
      Top             =   5892
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
      TabIndex        =   13
      Top             =   6168
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
      TabIndex        =   14
      Top             =   6444
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
   Begin EditLib.fpText fpStatus 
      Height          =   324
      Left            =   8784
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2136
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1080
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
   Begin EditLib.fpText fpCustName 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2520
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2496
      Width           =   4884
      _Version        =   196608
      _ExtentX        =   8615
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
      CaretInsert     =   0
      CaretOverWrite  =   3
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
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   35
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
   Begin EditLib.fpDateTime txtEntryDate 
      Height          =   324
      Left            =   8784
      TabIndex        =   2
      Top             =   2520
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
   Begin EditLib.fpDateTime txtCompletebyDate 
      Height          =   324
      Left            =   4164
      TabIndex        =   15
      Top             =   6888
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin EditLib.fpDateTime txtCompletedDate 
      Height          =   324
      Left            =   8424
      TabIndex        =   16
      Top             =   6888
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdMsg 
      Height          =   384
      Left            =   5268
      TabIndex        =   41
      Top             =   7920
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":0EDC
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   384
      Left            =   10248
      TabIndex        =   42
      Top             =   7920
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":10B6
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdConHist 
      Height          =   384
      Left            =   3756
      TabIndex        =   43
      Top             =   7920
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":1292
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdTranHist 
      Height          =   384
      Left            =   2232
      TabIndex        =   44
      Top             =   7920
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":146F
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdWorkHist 
      Height          =   384
      Left            =   720
      TabIndex        =   45
      Top             =   7920
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":164C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   8304
      TabIndex        =   17
      Top             =   7920
      Width           =   1680
      _Version        =   131072
      _ExtentX        =   2963
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":1829
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOwner 
      Height          =   384
      Left            =   6780
      TabIndex        =   46
      Top             =   7920
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
      ButtonDesigner  =   "frmWorkOrderEntry.frx":1A0B
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   47
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
            TextSave        =   "3:06 PM"
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
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   2280
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   648
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
      ControlType     =   0
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
   Begin VB.Label LabelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   576
      TabIndex        =   51
      Top             =   1152
      Width           =   3612
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Work Order Template:"
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
      Left            =   1008
      TabIndex        =   50
      Top             =   1560
      Width           =   3204
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   684
      Left            =   3228
      Top             =   432
      Width           =   5772
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Work Order Entry"
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
      TabIndex        =   49
      Top             =   600
      Width           =   4452
   End
   Begin VB.Label LabelAcctNo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2520
      TabIndex        =   40
      Top             =   2112
      Width           =   1140
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
      TabIndex        =   37
      Top             =   5040
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
      TabIndex        =   36
      Top             =   5316
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
      TabIndex        =   35
      Top             =   5604
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
      TabIndex        =   34
      Top             =   5880
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
      TabIndex        =   33
      Top             =   6156
      Width           =   276
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
      TabIndex        =   32
      Top             =   6432
      Width           =   276
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acct Status:"
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
      Height          =   324
      Left            =   7296
      TabIndex        =   31
      Top             =   2208
      Width           =   1428
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acct #:"
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
      Height          =   324
      Left            =   1368
      TabIndex        =   30
      Top             =   2184
      Width           =   1068
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete By Date:"
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
      Left            =   2172
      TabIndex        =   29
      Top             =   6960
      Width           =   1956
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
      TabIndex        =   28
      Top             =   4824
      Width           =   1212
   End
   Begin VB.Label Labe56 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   6612
      TabIndex        =   27
      Top             =   6960
      Width           =   1764
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date:"
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
      Left            =   7560
      TabIndex        =   26
      Top             =   2592
      Width           =   1284
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   1536
      TabIndex        =   25
      Top             =   2544
      Width           =   900
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
      TabIndex        =   24
      Top             =   2880
      Width           =   2652
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   4344
      X2              =   10392
      Y1              =   3024
      Y2              =   3024
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   2784
      X2              =   10392
      Y1              =   4968
      Y2              =   4968
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   1728
      X2              =   10488
      Y1              =   6792
      Y2              =   6792
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
      TabIndex        =   23
      Top             =   3120
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
      TabIndex        =   22
      Top             =   3396
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
      TabIndex        =   21
      Top             =   3684
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
      TabIndex        =   20
      Top             =   3960
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
      TabIndex        =   19
      Top             =   4236
      Width           =   276
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
      TabIndex        =   18
      Top             =   4512
      Width           =   276
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   5412
      Left            =   1344
      Top             =   1944
      Width           =   9540
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   804
      Left            =   3228
      Top             =   312
      Width           =   5772
   End
End
Attribute VB_Name = "frmWorkOrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim RecNo As Long, CntL As Long
Dim TransRec As Long, MsgRec As Long
Dim UBSetUpRec(1) As UBSetupRecType
Dim UBOwnerRec As UBOwnerRecType
Dim UBSetupLen As Integer
Dim OldBook As String, NBook As String
Dim FinalFlag As Boolean, UpDateOwner As Boolean
Dim BeenDone As Boolean, Behave As Integer
Dim BtnFnt As Double
Dim fromform As Form, toform As Form, codeopt As Integer
Dim EditFlag As Boolean, AddingFlag As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional Behavior As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  If Behavior <> 0 Then
    Behave = Behavior
  Else
    Behave = 0
  End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  DoEvents
End Sub
Private Sub Form_Activate()

  If Val(fpCustRecNo) > 0 And Not BeenDone Then
    RecNo& = Val(fpCustRecNo)
    BeenDone = True
    LoadCustInfo2Form
    DoEvents
  End If
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
        UBLog "Close via WorkOrderEntry by " + PWUser$
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
    Case vbKeyF2:
      KeyCode = 0
      If fpLoadDefault.Enabled Then
        Call fpLoadDefault_Click
      End If
    Case vbKeyF3:
      KeyCode = 0
      Call fpCmdWorkHist_Click
    Case vbKeyF4:
      KeyCode = 0
      Call fpCmdTranHist_Click
    Case vbKeyF6:
      KeyCode = 0
      Call fpCmdConHist_Click
    Case vbKeyF7:
      KeyCode = 0
      Call fpCmdMsg_Click
    Case vbKeyF8:
      KeyCode = 0
      Call fpCmdOwner_Click
    Case vbKeyEscape:
      fpCmdExit_Click
      KeyCode = 0
    Case vbKeyF10:  'save and print
      KeyCode = 0
      fpCmdSave_Click
    Case Else:
  End Select
End Sub

'Private Sub fpcboWOList_Click()
'  If Not BeenDone Then
'    getwodefault
'  End If
'End Sub

Private Sub fpcboWOList_GotFocus()
  BeenDone = False
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
      fpLoadDefault.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpLoadDefault_Click()
  If Not BeenDone Then
    If fpcboWOList.ListIndex <> 0 Then
      getwodefault
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
      txtEntryDate.SetFocus
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
       txtCompletebyDate.SetFocus
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


Private Sub txtEntryDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    fptxtWOInf(0).SetFocus
  End If
End Sub

Private Sub txtCompletebyDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    txtCompletedDate.SetFocus
  End If
End Sub
Private Sub txtCompletedDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    fpCmdSave.SetFocus
  End If
End Sub

Private Sub fpCmdExit_Click()
  If RecNo& > 0 Then
  If Chk4Change = True Then
    If MsgBox("Exit without saving work order?", vbYesNo, "Abandon Changes?") = vbNo Then
      Exit Sub
    End If
  End If
  End If
  ExitWorkOrderEntry
End Sub
Private Sub ExitWorkOrderEntry()

  UBLog "OUT: Work Order Entry."
  DoEvents
  RecNo = 0
  BeenDone = False
  TransRec = 0
  If Behave <> 55 Then
    If codeopt = 1 Then
      ActivateControls frmCustEditLookUP
    ElseIf codeopt = 2 Then
      ActivateControls frmDisplayList
    End If
  Else
    ActivateControls fromform
  End If
  Unload frmWorkOrderEntry
End Sub
Private Sub fpCmdConHist_Click()
  If RecNo > 0 Then
    frmRptConsumpHist.ShowCustConsHist (RecNo&)
  End If
End Sub
Private Sub fpCmdMsg_Click()
  If RecNo > 0 Then
    frmCustMsgEdit.CustRec = RecNo
    frmCustMsgEdit.Show vbModal
    DoEvents
    If CustHasMsg(RecNo) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      fpCmdMsg.ForeColor = &H80000012
      'fpCmdMsg.FontSize = BtnFnt
    End If
  End If

End Sub

Private Sub fpCmdOwner_Click()
  frmCustOwnerEdit.RecNo = RecNo
  frmCustOwnerEdit.Show vbModal
  DoEvents
  UpDateOwner = frmCustOwnerEdit.ActionFlag
  If UpDateOwner And RecNo > 0 Then  'an existing cust account
    'Call UBSaveOwnerInfo(RecNo)      'update owner info now. (user may not update cust)
    UpDateOwner = False
  End If                        'hey, Just forget about it.
  DoEvents
  'Call UNLoadOwnerForm
  'Unload frmCustOwnerEdit
End Sub

Private Sub fpCmdSave_Click()
'need to check stuff first then save
  If Chk4Blank = True Then
    MsgBox "You may not save a blank work order.", vbOKOnly, "Blank Fields"
    Exit Sub
  End If
'If Chk4Change = True Then
'  If MsgBox("Save Changes?", vbYesNo, "Save?") = vbNo Then
'    Exit Sub
'  End If
'End If
  SaveWorkOrderRec
  If MsgBox("Do you wish to print the work order now?", vbYesNo, "Print?") = vbYes Then
  frmReportOpt.Show 1
  DeActivateControls Me
  If rptopt = 1 Then
  'do graphic report
    PrintWorkOrders RecNo&, True
  ElseIf rptopt = 2 Then
  'do text report
    PrintWorkOrders RecNo&, False
  End If
  ActivateControls Me
  End If
'where to go from here??????????
  ExitWorkOrderEntry
End Sub

Private Sub fpCmdWorkHist_Click()
  If RecNo > 0 Then
    frmRptWrkOrdHist.ShowWrkOrdHistory (RecNo&)
  End If
End Sub
Private Sub fpCmdTranHist_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If TransRec > 0 Then
    'DeActivateControls Me
    DisplayCustTransList RecNo
    'ActivateControls Me
  Else
  MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
  End If
End Sub
Private Sub getwodefault()
  Dim WorkOrderDefLen As Integer, NumWOs As Long
  Dim UBWrkOrdD As Integer, Listrec As Long

  ReDim WorkOrderDef(1) As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef(1))

  If Chk4Change = True Then
  'stuff been entered so ask first
    If MsgBox("Do You Wish to Abandon Changes?", vbYesNo, "Abandon changes?") = vbNo Then
      Exit Sub
    End If
  End If
  'need to load info into fields from default file
  fpcboWOList.col = 0
  Listrec = fpcboWOList.ColText
  UBWrkOrdD = FreeFile
  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
  NumWOs = LOF(UBWrkOrdD) \ WorkOrderDefLen
  If Listrec& > 0 Then
    Get UBWrkOrdD, Listrec&, WorkOrderDef(1)
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

Private Sub LoadCustInfo2Form()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim UBCustF As Integer, LWTrans As Long, UBWrkOrd As Integer
  Dim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  Dim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  UBLog " IN: Work Order Entry."
  RecNo& = Val(fpCustRecNo)
  fpCustRecNo = 0

  UBCustF = FreeFile
  Open UBCustFile For Random Shared As UBCustF Len = UBCustRecLen
  Get UBCustF, RecNo&, UBCustRec(1)
  If UBCustRec(1).LastTrans > 0 Then
    TransRec = UBCustRec(1).LastTrans
  End If
  If CustHasMsg(RecNo) Then
    MsgAlertTimer.Enabled = True
    'MsgRec = tmpCustRec.MessageRec
  End If

  If UBCustRec(1).Status = "F" Then
    FinalFlag = True
  End If
  LabelAcctNo.Caption = RecNo&
  fpStatus.Text = " " + UBCustRec(1).Status
  fpCustName = QPTrim$(UBCustRec(1).CustName)
  LWTrans = UBCustRec(1).WOLastTrans
  If LWTrans > 0 Then
    EditFlag = True
    UBWrkOrd = FreeFile
    Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
    Get UBWrkOrd, LWTrans, WorkOrderRec(1)
    If WorkOrderRec(1).CompletedDate > 0 Then
      EditFlag = False
      AddingFlag = True
    Else
'      BCopy VARSEG(WorkOrderRec(1)), VarPtr(WorkOrderRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), WorkOrderRecLen, 0
'      UnPackBuffer 0, 0, Form$(), Fld()
      EditFlag = True
      AddingFlag = False
    End If
  Else
    EditFlag = False
    AddingFlag = True
  End If
  If AddingFlag Then
    LabelInfo.Caption = "New Work Order"
    txtEntryDate = Format(Now, "mm/dd/yyyy")
    Label17.Visible = True
    fpcboWOList.Visible = True
    fpcboWOList.InsertRow = " " & Chr$(9) & "ADD NEW WORK ORDER"
    GetWOList fpcboWOList
    fpcboWOList.ListIndex = 0
    fpLoadDefault.Enabled = True
  ElseIf EditFlag Then
    LabelInfo.Caption = "Edit Existing Work Order"
    txtEntryDate = Num2Date(WorkOrderRec(1).ENTRYDATE)
    fptxtWOInf(0).Text = QPTrim(WorkOrderRec(1).OrdersText.Text(1))
    fptxtWOInf(1).Text = QPTrim(WorkOrderRec(1).OrdersText.Text(2))
    fptxtWOInf(2).Text = QPTrim(WorkOrderRec(1).OrdersText.Text(3))
    fptxtWOInf(3).Text = QPTrim(WorkOrderRec(1).OrdersText.Text(4))
    fptxtWOInf(4).Text = QPTrim(WorkOrderRec(1).OrdersText.Text(5))
    fptxtWOInf(5).Text = QPTrim(WorkOrderRec(1).OrdersText.Text(6))
    fptxtWORem(0).Text = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
    fptxtWORem(1).Text = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
    fptxtWORem(2).Text = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
    fptxtWORem(3).Text = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
    fptxtWORem(4).Text = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
    fptxtWORem(5).Text = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
    If WorkOrderRec(1).CompleteByDate <> 0 Then
      txtCompletebyDate = Num2Date(WorkOrderRec(1).CompleteByDate)
    Else
      txtCompletebyDate = ""
    End If
    If WorkOrderRec(1).CompletedDate <> 0 Then
      txtCompletedDate = Num2Date(WorkOrderRec(1).CompletedDate)
    Else
      txtCompletedDate = ""
    End If
    Label17.Visible = False
    fpcboWOList.Visible = False
    fpLoadDefault.Enabled = False
  End If
  Close
Exit Sub

End Sub

Private Sub SaveWorkOrderRec()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer, whattrans As Long
  Dim UBCustF As Integer, LWTrans As Long, UBWrkOrd As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))


  UBCustF = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
  Get UBCustF, RecNo&, UBCustRec(1)

  UBWrkOrd = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
  'BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(WorkOrderRec(1)), VarPtr(WorkOrderRec(1)), WorkOrderRecLen, 0

  Select Case AddingFlag
  Case True
    If UBCustRec(1).WOLastTrans > 0 Then
      WorkOrderRec(1).PrevTransRec = UBCustRec(1).WOLastTrans
      whattrans = (LOF(UBWrkOrd) \ WorkOrderRecLen) + 1
      UBCustRec(1).WOLastTrans = whattrans
    Else
      whattrans = (LOF(UBWrkOrd) \ WorkOrderRecLen) + 1
      UBCustRec(1).WOLastTrans = whattrans
    End If
  Case False
    whattrans = UBCustRec(1).WOLastTrans
  End Select
  WorkOrderRec(1).ENTRYDATE = Date2Num(txtEntryDate)
  WorkOrderRec(1).CustRec = RecNo&
  WorkOrderRec(1).OrdersText.Text(1) = QPTrim(fptxtWOInf(0).Text)
  WorkOrderRec(1).OrdersText.Text(2) = QPTrim(fptxtWOInf(1).Text)
  WorkOrderRec(1).OrdersText.Text(3) = QPTrim(fptxtWOInf(2).Text)
  WorkOrderRec(1).OrdersText.Text(4) = QPTrim(fptxtWOInf(3).Text)
  WorkOrderRec(1).OrdersText.Text(5) = QPTrim(fptxtWOInf(4).Text)
  WorkOrderRec(1).OrdersText.Text(6) = QPTrim(fptxtWOInf(5).Text)
  WorkOrderRec(1).RepliesText.Text(1) = QPTrim(fptxtWORem(0).Text)
  WorkOrderRec(1).RepliesText.Text(2) = QPTrim(fptxtWORem(1).Text)
  WorkOrderRec(1).RepliesText.Text(3) = QPTrim(fptxtWORem(2).Text)
  WorkOrderRec(1).RepliesText.Text(4) = QPTrim(fptxtWORem(3).Text)
  WorkOrderRec(1).RepliesText.Text(5) = QPTrim(fptxtWORem(4).Text)
  WorkOrderRec(1).RepliesText.Text(6) = QPTrim(fptxtWORem(5).Text)
  WorkOrderRec(1).CompleteByDate = Date2Num(txtCompletebyDate)
  WorkOrderRec(1).CompletedDate = Date2Num(txtCompletedDate)


  Put UBWrkOrd, whattrans, WorkOrderRec(1)
  UBCustRec(1).WOLastTrans = whattrans
  Put UBCustF, RecNo&, UBCustRec(1)

  Close
  UBLog " Save: Work Order " + Str(whattrans)
  MsgBox "Work Order Saved.", vbOKOnly, "Save Complete"
End Sub
Private Function Chk4Blank()
  Dim cnt As Integer
  Chk4Blank = False
  cnt = 13
  If Len(txtEntryDate) = 0 Then cnt = cnt - 1
  'If fpcboWOList.ListIndex = 0 Then cnt = cnt + 1
  If fptxtWOInf(0).Text = "" Then cnt = cnt - 1
  If fptxtWOInf(1).Text = "" Then cnt = cnt - 1
  If fptxtWOInf(2).Text = "" Then cnt = cnt - 1
  If fptxtWOInf(3).Text = "" Then cnt = cnt - 1
  If fptxtWOInf(4).Text = "" Then cnt = cnt - 1
  If fptxtWOInf(5).Text = "" Then cnt = cnt - 1
  If fptxtWORem(0).Text = "" Then cnt = cnt - 1
  If fptxtWORem(1).Text = "" Then cnt = cnt - 1
  If fptxtWORem(2).Text = "" Then cnt = cnt - 1
  If fptxtWORem(3).Text = "" Then cnt = cnt - 1
  If fptxtWORem(4).Text = "" Then cnt = cnt - 1
  If fptxtWORem(5).Text = "" Then cnt = cnt - 1
  'If Len(txtCompletebyDate) = 0 Then cnt = cnt + 1
  'If Len(txtCompletedDate) = 0 Then cnt = cnt + 1
  If cnt < 2 Then Chk4Blank = True
End Function
Private Function Chk4Change()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim UBCustF As Integer, LWTrans As Long, UBWrkOrd As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))

  UBCustF = FreeFile
  Open UBCustFile For Random Shared As UBCustF Len = UBCustRecLen
  Get UBCustF, RecNo&, UBCustRec(1)
  
  LabelAcctNo.Caption = RecNo&
  fpStatus.Text = " " + UBCustRec(1).Status
  fpCustName = QPTrim$(UBCustRec(1).CustName)
  LWTrans = UBCustRec(1).WOLastTrans
  If AddingFlag Then
    If txtEntryDate <> Format(Now, "mm/dd/yyyy") Then Chk4Change = True
    'If fpcboWOList.ListIndex <> 0 Then Chk4Change = True
    If fptxtWOInf(0).Text <> "" Then Chk4Change = True
    If fptxtWOInf(1).Text <> "" Then Chk4Change = True
    If fptxtWOInf(2).Text <> "" Then Chk4Change = True
    If fptxtWOInf(3).Text <> "" Then Chk4Change = True
    If fptxtWOInf(4).Text <> "" Then Chk4Change = True
    If fptxtWOInf(5).Text <> "" Then Chk4Change = True
    If fptxtWORem(0).Text <> "" Then Chk4Change = True
    If fptxtWORem(1).Text <> "" Then Chk4Change = True
    If fptxtWORem(2).Text <> "" Then Chk4Change = True
    If fptxtWORem(3).Text <> "" Then Chk4Change = True
    If fptxtWORem(4).Text <> "" Then Chk4Change = True
    If fptxtWORem(5).Text <> "" Then Chk4Change = True
    If Len(txtCompletebyDate) <> 0 Then Chk4Change = True
    If Len(txtCompletedDate) <> 0 Then Chk4Change = True
  ElseIf EditFlag Then
    UBWrkOrd = FreeFile
    Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
    Get UBWrkOrd, LWTrans, WorkOrderRec(1)
    If txtEntryDate <> Num2Date(WorkOrderRec(1).ENTRYDATE) Then Chk4Change = True
    If fptxtWOInf(0).Text <> QPTrim(WorkOrderRec(1).OrdersText.Text(1)) Then Chk4Change = True
    If fptxtWOInf(1).Text <> QPTrim(WorkOrderRec(1).OrdersText.Text(2)) Then Chk4Change = True
    If fptxtWOInf(2).Text <> QPTrim(WorkOrderRec(1).OrdersText.Text(3)) Then Chk4Change = True
    If fptxtWOInf(3).Text <> QPTrim(WorkOrderRec(1).OrdersText.Text(4)) Then Chk4Change = True
    If fptxtWOInf(4).Text <> QPTrim(WorkOrderRec(1).OrdersText.Text(5)) Then Chk4Change = True
    If fptxtWOInf(5).Text <> QPTrim(WorkOrderRec(1).OrdersText.Text(6)) Then Chk4Change = True
    If fptxtWORem(0).Text <> QPTrim(WorkOrderRec(1).RepliesText.Text(1)) Then Chk4Change = True
    If fptxtWORem(1).Text <> QPTrim(WorkOrderRec(1).RepliesText.Text(2)) Then Chk4Change = True
    If fptxtWORem(2).Text <> QPTrim(WorkOrderRec(1).RepliesText.Text(3)) Then Chk4Change = True
    If fptxtWORem(3).Text <> QPTrim(WorkOrderRec(1).RepliesText.Text(4)) Then Chk4Change = True
    If fptxtWORem(4).Text <> QPTrim(WorkOrderRec(1).RepliesText.Text(5)) Then Chk4Change = True
    If fptxtWORem(5).Text <> QPTrim(WorkOrderRec(1).RepliesText.Text(6)) Then Chk4Change = True
    If txtCompletebyDate <> Num2Date(WorkOrderRec(1).CompleteByDate) Then Chk4Change = True
    If txtCompletedDate <> Num2Date(WorkOrderRec(1).CompletedDate) Then Chk4Change = True
  End If
  Close
End Function

Public Sub PrintWorkOrders(RecNo&, graphicflag As Boolean)
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim Dash As String, IdxNumOfRecs As Long, IdxName As String
  Dim ReportFile As String, RptHandle As Integer, Acct As Long
  Dim IdxRecLen As Integer, IdxFileSize As Long, cnt As Long
  Dim NumOfRecs As Long, Handle As Integer, UBCustF As Integer
  Dim UBWOFile As Integer, lcnt As Long, Book As Integer
  Dim BegRoute As Integer, EndRoute As Integer, ToPrint As String
  Dim Header As String, CopyCnt As Integer, MtrCnt As Integer
  Dim Rem1 As String, Rem2 As String, Rem3 As String, Rem4 As String
  Dim Rem5 As String, Rem6 As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
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
  FrmShowPctComp.Label1 = "Creating Work Order"
  FrmShowPctComp.Show , Me
  'Open Report File
  ReportFile$ = UBPath$ + "WORKORDR.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ' Location Order ********************************************************
  
  IdxName$ = UBPath$ + "UBCUSTBK.IDX"
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&(IdxName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  NumOfRecs = IdxNumOfRecs
  Handle = FreeFile
  Open IdxName$ For Random Shared As Handle Len = IdxRecLen
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IdxBuff(cnt&)
  Next
  Close Handle

  UBCustF = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen

  UBWOFile = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWOFile Len = WorkOrderRecLen
     FrmShowPctComp.ShowPctComp 1, 100
     If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        GoTo ExitHere
      End If

    'ShowProcessingScrn "Processing Work Order"
    'ShowPctComp 1, 1
    Get #UBCustF, RecNo&, UBCustRec(1)
    Get #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
    Acct& = RecNo&
    FrmShowPctComp.ShowPctComp 50, 100
    GoSub PrintThemOne
  

  'PRINT #RptHandle, FF$

  Close
  Erase UBCustRec, WorkOrderRec, IdxBuff
  FrmShowPctComp.ShowPctComp 100, 100
  Header$ = "Customer Work Orders "
  'PrintRptFile Header$, ReportFile$, LPTPort, RetCode, EntryPoint
  If graphicflag = False Then
    ViewPrint ReportFile$, Header$
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmWorkOrderEntry
    ARptWorkOrder.GetName ReportFile$
    ARptWorkOrder.startrpt
  End If
ExitHere:
  Exit Sub

PrintThemOne:
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(1))) > 0 Then
    Rem1$ = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
  Else
    Rem1$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(2))) > 0 Then
    Rem2$ = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
  Else
    Rem2$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(3))) > 0 Then
    Rem3$ = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
  Else
    Rem3$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(4))) > 0 Then
    Rem4$ = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
  Else
    Rem4$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(5))) > 0 Then
    Rem5$ = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
  Else
    Rem5$ = Dash$
  End If

  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(6))) > 0 Then
    Rem6$ = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
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
    Print #RptHandle, "    Work Order#: "; Using("######", UBCustRec(1).WOLastTrans); Tab(30); "Date Issued: "; Num2Date$(WorkOrderRec(1).ENTRYDATE)
    Print #RptHandle, "      Location#: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Tab(30); "Complete By: "; Num2Date$(WorkOrderRec(1).CompleteByDate)
    Print #RptHandle, "       Account#: "; Acct&; Tab(30); "  Completed: "; Num2Date$(WorkOrderRec(1).CompletedDate)
    Print #RptHandle, "  Customer Name: "; UBCustRec(1).CustName
    Print #RptHandle, "Service Address: "; UBCustRec(1).ServAddr
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Instruction or Description of Work Needed"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(1)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(2)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(3)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(4)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(5)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(6)
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
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #RptHandle, QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      End If
    Next
    Print #RptHandle, FF$;
  Else
    ToPrint$ = Num2Date$(WorkOrderRec(1).ENTRYDATE) + "~"
    ToPrint$ = ToPrint$ + Using("######", UBCustRec(1).WOLastTrans) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~"
    ToPrint$ = ToPrint$ + Str(Acct&) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).CustName + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).ServAddr + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(1) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(2) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(3) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(4) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(5) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(6) + "~"
    ToPrint$ = ToPrint$ + Rem1$ + "~"
    ToPrint$ = ToPrint$ + Rem2$ + "~"
    ToPrint$ = ToPrint$ + Rem3$ + "~"
    ToPrint$ = ToPrint$ + Rem4$ + "~"
    ToPrint$ = ToPrint$ + Rem5$ + "~"
    ToPrint$ = ToPrint$ + Rem6$

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      Else
        ToPrint$ = ToPrint$ + "~ "
      End If
    Next
    ToPrint$ = ToPrint$ + "~" + Num2Date$(WorkOrderRec(1).CompleteByDate) + "~"
    ToPrint$ = ToPrint$ + Num2Date$(WorkOrderRec(1).CompletedDate)

    Print #RptHandle, ToPrint$
    ToPrint$ = ""
  End If
  Return
End Sub
Private Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = fpCmdMsg.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      fpCmdMsg.ForeColor = &H80000012
      fpCmdMsg.FontSize = BtnFnt
    Case 2
      fpCmdMsg.ForeColor = &H80000011
      fpCmdMsg.FontSize = BtnFnt - 0.7
    Case 3
      fpCmdMsg.ForeColor = &H80000011
      fpCmdMsg.FontSize = BtnFnt - 1.4
    Case 4
      fpCmdMsg.ForeColor = &H80000010
      fpCmdMsg.FontSize = BtnFnt - 2.1
    Case 5
      fpCmdMsg.ForeColor = &H80000010
      fpCmdMsg.FontSize = BtnFnt - 2.8
    Case 6
      fpCmdMsg.ForeColor = &H8000000F
      fpCmdMsg.FontSize = BtnFnt - 3.5
    Case 7
      fpCmdMsg.ForeColor = &H8000000F
      fpCmdMsg.FontSize = BtnFnt - 4.2
    Case 8
      fpCmdMsg.ForeColor = &H8000000E
      fpCmdMsg.FontSize = BtnFnt - 4.9
    Case 9
      fpCmdMsg.ForeColor = &H8000000E
      fpCmdMsg.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
'  DoEvents
End Sub
