VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxBillGLSetUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax GL-Interface Account Setup"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxBillGLSetUp.frx":0000
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
      TabIndex        =   3
      Top             =   2892
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
      ColDesigner     =   "frmTaxBillGLSetUp.frx":08CA
   End
   Begin EditLib.fpText fptxtTPDebit 
      Height          =   375
      Left            =   6420
      TabIndex        =   12
      Top             =   2895
      Width           =   1935
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
      Left            =   4800
      TabIndex        =   0
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
      ButtonDesigner  =   "frmTaxBillGLSetUp.frx":0B1E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   630
      Left            =   7890
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7710
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
      ButtonDesigner  =   "frmTaxBillGLSetUp.frx":0CFD
   End
   Begin EditLib.fpText fptxtTPCredit 
      Height          =   375
      Left            =   8580
      TabIndex        =   13
      Top             =   2895
      Width           =   1935
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
      Height          =   375
      Left            =   6420
      TabIndex        =   14
      Top             =   3375
      Width           =   1935
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
      Height          =   375
      Left            =   8580
      TabIndex        =   15
      Top             =   3375
      Width           =   1935
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
      Height          =   375
      Left            =   6420
      TabIndex        =   16
      Top             =   3855
      Width           =   1935
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
      Height          =   375
      Left            =   8580
      TabIndex        =   17
      Top             =   3855
      Width           =   1935
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   495
      Left            =   7560
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1785
      _Version        =   131072
      _ExtentX        =   3149
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
      ButtonDesigner  =   "frmTaxBillGLSetUp.frx":0EDA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLTest 
      Height          =   630
      Left            =   1710
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7710
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
      ButtonDesigner  =   "frmTaxBillGLSetUp.frx":10B9
   End
   Begin EditLib.fpText fptxtLLDebit 
      Height          =   375
      Left            =   6420
      TabIndex        =   20
      Top             =   4350
      Width           =   1935
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
      Height          =   375
      Left            =   8580
      TabIndex        =   21
      Top             =   4350
      Width           =   1935
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
      Height          =   375
      Left            =   6420
      TabIndex        =   22
      Top             =   5280
      Width           =   1935
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
      Height          =   375
      Left            =   8580
      TabIndex        =   23
      Top             =   5280
      Width           =   1935
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
      Height          =   375
      Left            =   6420
      TabIndex        =   24
      Top             =   5760
      Width           =   1935
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
      Height          =   375
      Left            =   8580
      TabIndex        =   25
      Top             =   5760
      Width           =   1935
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
      Height          =   375
      Left            =   6420
      TabIndex        =   26
      Top             =   6240
      Width           =   1935
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
      Height          =   375
      Left            =   8580
      TabIndex        =   27
      Top             =   6240
      Width           =   1935
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
      Height          =   375
      Left            =   3420
      TabIndex        =   32
      Top             =   4440
      Width           =   1935
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
      Left            =   3420
      TabIndex        =   31
      Top             =   5400
      Width           =   2895
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
      Height          =   330
      Left            =   3060
      TabIndex        =   30
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11160
      X2              =   3040
      Y1              =   4920
      Y2              =   4920
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
      Left            =   3420
      TabIndex        =   29
      Top             =   5880
      Width           =   2895
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
      Left            =   3420
      TabIndex        =   28
      Top             =   6360
      Width           =   2895
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
      Left            =   4313
      TabIndex        =   11
      Top             =   1710
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   2280
      Y2              =   7420
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5160
      Left            =   1080
      Top             =   2280
      Width           =   10095
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
      Height          =   375
      Left            =   8820
      TabIndex        =   10
      Top             =   2535
      Width           =   1575
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
      Height          =   375
      Left            =   6660
      TabIndex        =   9
      Top             =   2535
      Width           =   1575
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
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   3975
      Width           =   1695
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
      Height          =   375
      Left            =   3420
      TabIndex        =   7
      Top             =   3495
      Width           =   1575
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
      Height          =   375
      Left            =   3420
      TabIndex        =   6
      Top             =   3015
      Width           =   1935
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
      Left            =   4313
      TabIndex        =   5
      Top             =   1215
      Width           =   3375
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
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   2535
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax GL Interface Account Setup"
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
      Left            =   3150
      TabIndex        =   2
      Top             =   510
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   360
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1493
      Top             =   270
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxBillGLSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Public GThisYear As String
  Dim TempTaxDBAcct As String
  Dim TempTaxCRAcct As String
  Dim TempIntDBAcct As String
  Dim TempIntCRAcct As String
  Dim TempAdvDBAcct As String
  Dim TempAdvCRAcct As String
  Dim TempLtLstDBAcct As String
  Dim TempLtLstCRAcct As String
  Dim TempOpt1DBAcct As String
  Dim TempOpt1CRAcct As String
  Dim TempOpt2DBAcct As String
  Dim TempOpt2CRAcct As String
  Dim TempOpt3DBAcct As String
  Dim TempOpt3CRAcct As String
  Dim TempYear As Integer
  Dim Exit2Bill As Boolean
  Dim Exit2Adv As Boolean
  Dim Exit2Int As Boolean
  Dim Exit2Man As Boolean

Private Sub cmdExit_Click()
  If Check4Changes = True Then
    Exit Sub
  End If
  Call LogSaves
'  If Exit2Bill = True Then
'    If Exist("revglbill.dat") Then KillFile "revglbill.dat"
'    frmTaxPrebilling.Show
'    DoEvents
'    Unload Me
'    Exit Sub
'  ElseIf Exit2Adv = True Then
'    If Exist("revgladv.dat") Then KillFile "revgladv.dat"
'    frmTaxCalcAdCol.Show
'    DoEvents
'    Unload Me
'    Exit Sub
'  ElseIf Exit2Int = True Then
'    If Exist("revglint.dat") Then KillFile "revglint.dat"
'    frmTaxCalcInterest.Show
'    DoEvents
'    Unload Me
'    Exit Sub
'  ElseIf Exit2Man = True Then
'    If Exist("revglman.dat") Then KillFile "revglman.dat"
'    frmTaxManualBillEntry.Show
'    DoEvents
'    Unload Me
'    Exit Sub
'  End If
  KillFile "taxbillGL.dat"
  'frmTaxBillSetUpMenu.Show
  DoEvents
  Unload frmTaxGLList
  Unload Me
End Sub

Private Sub cmdGLList_Click()
  frmTaxGLList.Show ' vbModal
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
     frmTaxMsg.Label1.Caption = "ERROR: No General Ledger file could be found. The General Ledger list cannot be loaded."
     frmTaxMsg.Label1.Top = 900
     frmTaxMsg.Show vbModal
     Close GLHandle
     Exit Sub
   End If
   
   If GLCnt < IdxCnt Then
     frmTaxMsg.Label1.Caption = "ERROR: The GL index count is greater than the GL file count."
     frmTaxMsg.Label1.Top = 900
     frmTaxMsg.Show vbModal
   End If
   
   ReDim GTestOK(1 To 6) As Boolean
   ReDim GTestNums(1 To 6) As String
   ReDim GTestDbCrt(1 To 6) As String
   ReDim GTestDesc(1 To 6) As String
   For x = 1 To 6
     GTestNums(x) = ""
     GTestDbCrt(x) = ""
     GTestDesc(x) = ""
   Next x
   
   For x = 1 To 6
     GTestOK(x) = False
     Select Case x
       Case 1
         GTestNums(x) = QPTrim$(fptxtTPDebit.Text)
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Tax Principle"
       Case 2
         GTestNums(x) = QPTrim$(fptxtTPCredit.Text)
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Tax Principle"
       Case 3
         GTestNums(x) = QPTrim$(fptxtIDebit.Text)
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Interest"
       Case 4
         GTestNums(x) = QPTrim$(fptxtICredit.Text)
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Interest"
       Case 5
         GTestNums(x) = QPTrim$(fptxtACDebit.Text)
         GTestDbCrt(x) = "Debit"
         GTestDesc(x) = "Adv/Collect"
       Case 6
         GTestNums(x) = QPTrim$(fptxtACCredit.Text)
         GTestDbCrt(x) = "Credit"
         GTestDesc(x) = "Adv/Collect"
     End Select
   Next x
   
   For x = 1 To IdxCnt
     If IdxRecs(x) <> 0 Then
       Get GLHandle, IdxRecs(x), GLRec
       If GLRec.Deleted Then GoTo SkipIt
       For y = 1 To 6
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
   
   For x = 1 To 6
     If GTestOK(x) = False And GTestNums(x) <> "" Then
       frmBadGLList.Show vbModal
       Exit For
     End If
   Next x
   
   If x > 6 Then
     frmTaxMsg.Label1.Caption = "All G/L numbers entries have been verified."
     frmTaxMsg.Label1.Top = 900
     frmTaxMsg.Show vbModal
   End If
     
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillGLSetUp", "cmdGLTest_Click", Erl)
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
   ' frmTaxBillSetUpMenu.Show
    DoEvents
    Unload Me
End Sub


Private Sub cmdSave_Click()
  Dim RevRec As TaxAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  
  OpenTaxGLInterBill RRHandle
  Get RRHandle, 1, RevRec
  For x = 1 To 51
    If RevRec.TaxAcct(x).TaxYear = CInt(GThisYear) Then
      RevRec.TaxAcct(x).TaxDBAcct = QPTrim$(fptxtTPDebit.Text)
      RevRec.TaxAcct(x).TaxCRAcct = QPTrim$(fptxtTPCredit.Text)
      RevRec.TaxAcct(x).IntDBAcct = QPTrim$(fptxtIDebit.Text)
      RevRec.TaxAcct(x).IntCRAcct = QPTrim$(fptxtICredit.Text)
      RevRec.TaxAcct(x).AdvDBAcct = QPTrim$(fptxtACDebit.Text)
      RevRec.TaxAcct(x).AdvCRAcct = QPTrim$(fptxtACCredit.Text)
      RevRec.TaxAcct(x).LtLstCRAcct = QPTrim$(fptxtLLCredit.Text)
      RevRec.TaxAcct(x).LtLstDBAcct = QPTrim$(fptxtLLDebit.Text)
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
    TempOpt1DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct)
    TempOpt1CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct)
    TempOpt2DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct)
    TempOpt2CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct)
    TempOpt3DBAcct = QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct)
    TempOpt3CRAcct = QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct)
    TempYear = RevRec.TaxAcct(x).TaxYear
  End If
  Unload frmTaxGLList
  
  Call Savemsg(900, "Your Bill Setup Data has been saved successfully.")
  If Exist("revglbill.dat") Then
    KillFile "revglbill.dat"
  ElseIf Exist("revgladv.dat") Then
    KillFile "revgladv.dat"
  ElseIf Exist("revglint.dat") Then
    KillFile "revglint.dat"
  End If
  
  If Exit2Bill = True Then
'    frmTaxPrebilling.Show
    DoEvents
    Unload Me
  ElseIf Exit2Adv = True Then
'    frmTaxCalcAdCol.Show
    DoEvents
    Unload Me
  ElseIf Exit2Int = True Then
'    frmTaxCalcInterest.Show
    DoEvents
    Unload Me
  ElseIf Exit2Man = True Then
'    frmTaxManualBillEntry.Show
    DoEvents
    Unload Me
  End If
    
    
    
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
      Unload frmTaxGLList
      KillFile "taxbillGL.dat"
      ClearInUse PWcnt
      TXLog ("CM.exe terminated via menu bar on frmTaxBillGLSetUp.")
      Call CMTerminate
      End
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
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
  TXLog ("CM-User opened frmTaxBillGLSetUp.")
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim RevRec As TaxAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  Dim One As Integer
  Dim AHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Exit2Bill = False
  Exit2Adv = False
  Exit2Int = False
  Exit2Man = False
  
  If Exist("revglbill.dat") Then
    Exit2Bill = True
  ElseIf Exist("revgladv.dat") Then
    Exit2Adv = True
  ElseIf Exist("revglint.dat") Then
    Exit2Int = True
  ElseIf Exist("revglman.dat") Then
    Exit2Man = True
  End If
  
  If QPTrim$(TaxMasterRec.OptRev1) <> "" Then
    Label10.Caption = QPTrim$(TaxMasterRec.OptRev1)
  Else
    Label10.Caption = "NO OPTION 1 SAVED"
  End If
  
  If QPTrim$(TaxMasterRec.OptRev2) <> "" Then
    Label12.Caption = QPTrim$(TaxMasterRec.OptRev2)
  Else
    Label12.Caption = "NO OPTION 2 SAVED"
  End If
  
  If QPTrim$(TaxMasterRec.OptRev3) <> "" Then
    Label13.Caption = QPTrim$(TaxMasterRec.OptRev3)
  Else
    Label13.Caption = "NO OPTION 3 SAVED"
  End If
  
  One = 1
  AHandle = FreeFile
  Open "taxbillGL.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  OpenTaxGLInterBill RRHandle
  If Exist(TxGLInterBill) Then
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

End Sub

Private Sub fpListYear_Click()
  Dim RevRec As TaxAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  
  If Exist("revglbill.dat") Then
    KillFile "revglbill.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  ElseIf Exist("revgladv.dat") Then
    KillFile "revgladv.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  ElseIf Exist("revglint.dat") Then
    KillFile "revglint.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  ElseIf Exist("revglman.dat") Then
    KillFile "revglman.dat"
    fpListYear.SearchText = CStr(GThisYear)
    fpListYear.Action = 0
    fpListYear.ListIndex = fpListYear.SearchIndex
    fpListYear.Row = fpListYear.ListIndex
    fpListYear.TopIndex = fpListYear.Row
  Else
    fpListYear.Row = fpListYear.ListIndex
    GThisYear = fpListYear.Text
  End If
  If QPTrim$(GThisYear) = "" Then
    Close
    Exit Sub
  End If
  lblYear.Caption = "For Year " + GThisYear
  
  OpenTaxGLInterBill RRHandle
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
    fptxtOR1Debit.Text = ""
    fptxtOR1Credit.Text = ""
    fptxtOR2Debit.Text = ""
    fptxtOR2Credit.Text = ""
    fptxtOR3Debit.Text = ""
    fptxtOR3Credit.Text = ""
  End If
End Sub

Private Sub fptxtACCredit_DblClick(Button As Integer)
  fptxtACCredit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtACDebit_DblClick(Button As Integer)
  fptxtACDebit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtICredit_DblClick(Button As Integer)
  fptxtICredit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtIDebit_DblClick(Button As Integer)
  fptxtIDebit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtLLCredit_DblClick(Button As Integer)
  fptxtLLCredit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtLLDebit_DblClick(Button As Integer)
  fptxtLLDebit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR1Credit_DblClick(Button As Integer)
  fptxtOR1Credit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR1Debit_DblClick(Button As Integer)
  fptxtOR1Debit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR2Credit_DblClick(Button As Integer)
  fptxtOR2Credit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR2Debit_DblClick(Button As Integer)
  fptxtOR2Debit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR3Credit_DblClick(Button As Integer)
  fptxtOR3Credit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOR3Debit_DblClick(Button As Integer)
  fptxtOR3Debit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtTPCredit_DblClick(Button As Integer)
  fptxtTPCredit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtTPDebit_DblClick(Button As Integer)
  fptxtTPDebit.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub LogSaves()
  Dim ThisDesc$
  
  If InStr(lblYear.Caption, CStr(TempYear)) = 0 Then Exit Sub
  
  If QPTrim$(TempTaxDBAcct) = "" Then TempTaxDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtTPDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempTaxDBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Tax Principle Debit was changed from " + QPTrim$(TempTaxDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempTaxCRAcct) = "" Then TempTaxCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtTPCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempTaxCRAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Tax Principle Credit was changed from " + QPTrim$(TempTaxCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempIntDBAcct) = "" Then TempIntDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtIDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempIntDBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Interest Debit was changed from " + QPTrim$(TempIntDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempIntCRAcct) = "" Then TempIntCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtICredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempIntCRAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Interest Credit was changed from " + QPTrim$(TempIntCRAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempAdvDBAcct) = "" Then TempAdvDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtACDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempAdvDBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Adv/Collect Debit was changed from " + QPTrim$(TempAdvDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempAdvCRAcct) = "" Then TempAdvCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtACCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempAdvCRAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Adv/Collect Credit was changed from " + QPTrim$(TempAdvCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempLtLstDBAcct) = "" Then TempLtLstDBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtLLDebit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempLtLstDBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Late Listing Debit was changed from " + QPTrim$(TempLtLstDBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempLtLstCRAcct) = "" Then TempLtLstCRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtLLCredit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempLtLstCRAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": Late Listing Credit was changed from " + QPTrim$(TempLtLstCRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempOpt1DBAcct) = "" Then TempOpt1DBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR1Debit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt1DBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label10.Caption + " Debit was changed from " + QPTrim$(TempOpt1DBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt1CRAcct) = "" Then TempOpt1CRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR1Credit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt1CRAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label10.Caption + " Credit was changed from " + QPTrim$(TempOpt1CRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempOpt2DBAcct) = "" Then TempOpt2DBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR2Debit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt2DBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label12.Caption + " Debit was changed from " + QPTrim$(TempOpt2DBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt2CRAcct) = "" Then TempOpt2CRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR2Credit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt2CRAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label12.Caption + " Credit was changed from " + QPTrim$(TempOpt2CRAcct) + " to " + ThisDesc + " and saved.")
  End If

  If QPTrim$(TempOpt3DBAcct) = "" Then TempOpt3DBAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR3Debit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt3DBAcct) <> ThisDesc Then
    TXLog ("CM-frmTaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label13.Caption + " Debit was changed from " + QPTrim$(TempOpt3DBAcct) + " to " + ThisDesc + " and saved.")
  End If
  
  If QPTrim$(TempOpt3CRAcct) = "" Then TempOpt3CRAcct = "BLANK"
  ThisDesc = QPTrim$(fptxtOR3Credit.Text)
  If ThisDesc = "" Then
    ThisDesc = "BLANK"
  End If
  If QPTrim$(TempOpt3CRAcct) <> ThisDesc Then
    TXLog ("frmTaxBillSetUp: For Year " + CStr(TempYear) + ": " + Label13.Caption + " Credit was changed from " + QPTrim$(TempOpt3CRAcct) + " to " + ThisDesc + " and saved.")
  End If

End Sub

Private Function Check4Changes() As Boolean
  Dim RevRec As TaxAcctsType
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
  Dim OR1DBAcct As String
  Dim OR1CRAcct As String
  Dim OR2DBAcct As String
  Dim OR2CRAcct As String
  Dim OR3DBAcct As String
  Dim OR3CRAcct As String
  Dim ThisStr As String
  Dim NewDesc As String
  Dim Thisx As Integer
  Dim choice As String
  
  Check4Changes = False
  If Exist(TxGLInterBill) Then
    OpenTaxGLInterBill RRHandle
    Get RRHandle, 1, RevRec
  Else
    frmTaxMsgWOpts.Label1.Caption = "Are you sure you want to exit without saving?"
    frmTaxMsgWOpts.Label1.Top = 900
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtTPDebit.SetFocus
      Check4Changes = True
      Exit Function
    Else
      Unload frmTaxMsgWOpts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Tax Principle Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Tax Principle Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Interest Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Interest Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Adv/Collect Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Adv/Collect Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Late Listing Debit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Late Listing Credit' field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      RevRec.TaxAcct(Thisx).LtLstCRAcct = QPTrim$(ThisControl.Text)
      Put RRHandle, 1, RevRec
      Call Savemsg(900, "Late Listing Credit has been saved successfully.")
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
    frmTaxMsgW4Opts.Label1.Caption = "The '" + Label10.Caption + "' Debit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The '" + Label10.Caption + "' Credit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The '" + Label12.Caption + "' Debit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The '" + Label12.Caption + "' Credit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The '" + Label13.Caption + "' Debit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The '" + Label13.Caption + "' Credit field has been changed from " + ThisStr + " to " + NewDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
       ' frmTaxBillSetUpMenu.Show
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxBillGLSetUp", "Check4Changes", Erl)
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

