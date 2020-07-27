VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxCustLookup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Lookup"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxCustLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   3210
      Left            =   60
      TabIndex        =   0
      Tag             =   $"frmVATaxCustLookup.frx":08CA
      Top             =   5100
      Width           =   11520
      _Version        =   196608
      _ExtentX        =   20320
      _ExtentY        =   5662
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   2
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
      BorderStyle     =   1
      BorderColor     =   8454143
      BorderWidth     =   3
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
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   3
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
      ColumnHeaderShow=   -1  'True
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
      ColDesigner     =   "frmVATaxCustLookup.frx":0A32
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   420
      Left            =   9960
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxCustLookup.frx":0DD7
   End
   Begin VB.OptionButton OptAll 
      BackColor       =   &H008F8265&
      Caption         =   "All"
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
      Height          =   252
      Left            =   360
      TabIndex        =   30
      Top             =   4680
      Width           =   732
   End
   Begin VB.OptionButton OptNoProp 
      BackColor       =   &H008F8265&
      Caption         =   "Own No Prop"
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
      Height          =   252
      Left            =   4200
      TabIndex        =   29
      Top             =   4680
      Width           =   1572
   End
   Begin VB.OptionButton OptPers 
      BackColor       =   &H008F8265&
      Caption         =   "Pers Only"
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
      Height          =   252
      Left            =   2760
      TabIndex        =   28
      Top             =   4680
      Width           =   1212
   End
   Begin VB.OptionButton OptReal 
      BackColor       =   &H008F8265&
      Caption         =   "Real Only"
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
      Height          =   252
      Left            =   1320
      TabIndex        =   27
      Top             =   4680
      Width           =   1212
   End
   Begin EditLib.fpText fptxtPersPin 
      Height          =   390
      Left            =   2490
      TabIndex        =   1
      Top             =   3840
      Width           =   3255
      _Version        =   196608
      _ExtentX        =   5741
      _ExtentY        =   688
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
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRealPin 
      Height          =   390
      Left            =   2490
      TabIndex        =   2
      Top             =   3330
      Width           =   3255
      _Version        =   196608
      _ExtentX        =   5741
      _ExtentY        =   688
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
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxCustLookup.frx":0FB3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   420
      Left            =   7920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1935
      _Version        =   131072
      _ExtentX        =   3413
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxCustLookup.frx":1191
   End
   Begin EditLib.fpText fptxtOptSearch 
      Height          =   390
      Left            =   8280
      TabIndex        =   5
      Top             =   1695
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4890
      _ExtentY        =   698
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
      MaxLength       =   15
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
   Begin EditLib.fpText fptxtSrvcAdd 
      Height          =   390
      Left            =   2130
      TabIndex        =   8
      Top             =   2355
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   698
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
      MaxLength       =   35
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
   Begin EditLib.fpText fptxtAcctNum 
      Height          =   390
      Left            =   2250
      TabIndex        =   9
      Top             =   1860
      Width           =   3495
      _Version        =   196608
      _ExtentX        =   6165
      _ExtentY        =   688
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
      CharValidationText=   "0 1 2 3 4 5 6 7 8 9"
      MaxLength       =   10
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
   Begin EditLib.fpText fptxtSearchName 
      Height          =   390
      Left            =   1770
      TabIndex        =   10
      Top             =   1365
      Width           =   3975
      _Version        =   196608
      _ExtentX        =   7006
      _ExtentY        =   698
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
      MaxLength       =   35
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
   Begin EditLib.fpText fptxtSSN1 
      Height          =   390
      Left            =   2880
      TabIndex        =   11
      Top             =   2835
      Width           =   795
      _Version        =   196608
      _ExtentX        =   1397
      _ExtentY        =   698
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
      AlignTextH      =   1
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
      MaxLength       =   3
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
   Begin EditLib.fpText fptxtSSN2 
      Height          =   390
      Left            =   3960
      TabIndex        =   12
      Top             =   2835
      Width           =   465
      _Version        =   196608
      _ExtentX        =   825
      _ExtentY        =   698
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
      AlignTextH      =   1
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
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
   Begin EditLib.fpText fptxtSSN3 
      Height          =   390
      Left            =   4680
      TabIndex        =   13
      Top             =   2835
      Width           =   1065
      _Version        =   196608
      _ExtentX        =   1879
      _ExtentY        =   688
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
      AlignTextH      =   1
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtOptRealSrch 
      Height          =   390
      Left            =   8280
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   688
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
      MaxLength       =   15
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
   Begin EditLib.fpText fptxtOptPersSrch 
      Height          =   390
      Left            =   8280
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   688
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
      MaxLength       =   15
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
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   720
      X2              =   6000
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Opt'l:"
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
      Height          =   300
      Left            =   6360
      TabIndex        =   35
      Top             =   1800
      Width           =   1755
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Real Optional:"
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
      Height          =   300
      Left            =   6480
      TabIndex        =   34
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Pers Search Entry"
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
      Height          =   300
      Left            =   8160
      TabIndex        =   33
      Top             =   3480
      Width           =   2955
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Opt'l:"
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
      Height          =   300
      Left            =   6480
      TabIndex        =   32
      Top             =   3960
      Width           =   1755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   6000
      X2              =   6000
      Y1              =   5160
      Y2              =   2820
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   26
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Real Search Entry:"
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
      Height          =   300
      Left            =   8160
      TabIndex        =   25
      Top             =   2400
      Width           =   2955
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3945
      Left            =   70
      Top             =   1200
      Width           =   11490
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1488
      Top             =   180
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Lookup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2928
      TabIndex        =   24
      Top             =   312
      Width           =   6012
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   23
      Top             =   1980
      Width           =   1875
   End
   Begin VB.Label lblPin 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Address:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   2460
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Social Security Number:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   21
      Top             =   2940
      Width           =   2595
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Name:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   20
      Top             =   1470
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Search Entry:"
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
      Height          =   300
      Left            =   8160
      TabIndex        =   19
      Top             =   1335
      Width           =   2955
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3210
      Left            =   6000
      Top             =   1200
      Width           =   5565
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   300
      Left            =   3720
      TabIndex        =   18
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   300
      Left            =   4455
      TabIndex        =   17
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label LabelDel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   16
      Top             =   660
      Width           =   2592
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Real Property Pin#:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   3450
      Width           =   2310
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Prop Pin#:"
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
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   2355
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1488
      Top             =   120
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxCustLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdClear_Click()
  fptxtSearchName.Text = ""
  fptxtAcctNum.Text = ""
  fptxtSrvcAdd.Text = ""
  fptxtSSN1.Text = ""
  fptxtSSN2.Text = ""
  fptxtSSN3.Text = ""
  fptxtRealPin.Text = ""
  fptxtPersPin.Text = ""
  fptxtOptSearch.Text = ""
  fptxtOptRealSrch.Text = ""
  fptxtOptPersSrch.Text = ""
  fpList1.Clear
  OptAll.Value = True
  OptReal.Value = False
  OptPers.Value = False
  OptNoProp.Value = False
End Sub

Private Sub cmdExit_Click()
  If DelAbs = True Then
    frmVATaxAbsMaint.Show
    DoEvents
  ElseIf RPayEntry = True Then
    Unload Me
  ElseIf PPayEntry = True Then
    Unload Me
  ElseIf Exist("C:\CPWork\txradjust.dat") Then
    Unload Me
  ElseIf Exist("C:\CPWork\txpadjust.dat") Then
    Unload Me
  ElseIf Exist("C:\CPWork\rmanualbill.dat") Then
    frmVATaxManualBillEntry.Show
    DoEvents
    Unload Me
  ElseIf Exist("C:\CPWork\pmanualbill.dat") Then
    frmVATaxPManualBillEntry.Show
    DoEvents
    Unload Me
  ElseIf Exist("C:\CPWork\custinq.dat") Then
    DoEvents
    Unload Me
  ElseIf Exist("C:\CPWork\custtranshist.dat") Then
    frmVATaxCustTHistRpt.Show
    DoEvents
    Unload Me
  Else
    frmVATaxCustMaintMenu.Show
    DoEvents
  End If
  KillFile "txpyment.dat"
  EditCust = False
  AddCust = False
  THistRpt = False
  DelAbs = False
  RPayEntry = False
  PPayEntry = False
  
  Unload Me
End Sub

Public Sub cmdSearch_Click()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim x As Long, y As Long, z As Long
  Dim SSN As String
  Dim SrvcAdd As String
  Dim AcctNum As String
  Dim SearchName As String
  Dim RealPin As String
  Dim PersPin As String
  Dim PersSearch As String
  Dim RealSearch As String
  Dim FoundIt As Boolean
  Dim MatSSN As String
  Dim FoundMatch As Long 'Integer 8/31/09
  Dim OneSSN As String * 3
  Dim TwoSSN As String * 2
  Dim ThreeSSN As String * 4
  Dim NewSSN As String
  Dim CopyNewSSN As String
  Dim NewMatSSN As String
  Dim PrintAll As Boolean
  Dim OptSrchFld As String
  Dim SwapCustNum1 As Long
  Dim SwapCustNum2 As Long
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NextRec As Long
  Dim NextPRec As Long 'added 8/22/06
  Dim PersFoundIt As Boolean
  Dim RealFoundIt As Boolean
  Dim RealSrchFoundIt As Boolean
  Dim PersSrchFoundIt As Boolean
  Dim PrintRProp As String * 22
  Dim PrintPProp As String * 22
  Dim PrintRealSrch As String * 20
  Dim PrintPersSrch As String * 20
  Dim IdxFlag As Boolean
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxRec As CustNameIdxType
  Dim ThisCustRec As Long
  Dim PrintYN As Boolean '5/26/06
  Dim NextRRec As Long 'added 8/22/06
  Dim DontPrintIt As Boolean 'added 8/22/06
  Dim NextRSRec As Long
  Dim NextPSRec As Long
  
  'On Error GoTo ERRORSTUFF
  
  OpenPersPropFile PHandle, NumOfPersRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  
  IdxFlag = False
  If QPTrim$(fptxtSearchName.Text) <> "" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  End If
  
  SwapCustNum1 = GCustNum
  fpList1.Clear
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  If NumOfTaxCusts = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no customers on file. Search aborted."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  OneSSN = fptxtSSN1.Text
  TwoSSN = fptxtSSN2.Text
  ThreeSSN = fptxtSSN3.Text
  
  FoundMatch = 0
  SSN = OneSSN + TwoSSN + ThreeSSN
  
  ReDim ThatSSN(1 To 9) As String
  ReDim ThisSSN(1 To 9) As String
  NewSSN = ""
  For x = 1 To 9
    ThisSSN(x) = Mid(SSN, x, 1)
    If Not IsNumeric(ThisSSN(x)) Then
      ThisSSN(x) = "x"
    End If
    NewSSN = NewSSN + ThisSSN(x)
  Next x
  
  CopyNewSSN = NewSSN
  
  SrvcAdd = QPTrim$(fptxtSrvcAdd.Text)
  AcctNum = CStr(fptxtAcctNum.Text)
  SearchName = QPTrim$(fptxtSearchName.Text)
  OptSrchFld = QPTrim$(fptxtOptSearch.Text)
  RealPin = QPTrim$(fptxtRealPin.Text)
  PersPin = QPTrim$(fptxtPersPin.Text)
  RealSearch = QPTrim$(fptxtOptRealSrch.Text)
  PersSearch = QPTrim$(fptxtOptPersSrch.Text)
  
  NewMatSSN = ""
  PrintAll = False
  If Len(QPTrim$(SSN)) = 0 And Len(QPTrim$(SearchName)) = 0 And Len(QPTrim$(AcctNum)) = 0 And Len(QPTrim$(SrvcAdd)) = 0 And Len(QPTrim$(OptSrchFld)) = 0 And Len(QPTrim$(RealPin)) = 0 And Len(QPTrim$(PersPin)) = 0 And Len(QPTrim$(RealSearch)) = 0 And Len(QPTrim$(PersSearch)) = 0 Then
    PrintAll = True
  End If
  
  If IdxFlag = True Then
    NumOfTaxCusts = NumOfIdx
  End If
  
  For x = 1 To NumOfTaxCusts
    If IdxFlag = False Then
      Get TCHandle, x, TaxCustRec
      ThisCustRec = x
    Else
      Get TCHandle, IdxArray(x), TaxCustRec
      ThisCustRec = IdxArray(x)
    End If
    GoSub GetProp '5/26/06
    If PrintYN = False Then GoTo NoMatch '5/26/06
    If TaxCustRec.Deleted <> 0 Then GoTo NoMatch
    If PrintAll = True Then GoTo PrintIt
    FoundIt = False
    PersFoundIt = False
    RealFoundIt = False
    RealSrchFoundIt = False
    PersSrchFoundIt = False
    PrintRProp = ""
    PrintPProp = ""
    NewMatSSN = ""
    NewSSN = CopyNewSSN
    If Len(QPTrim$(SSN)) > 0 Then
      MatSSN = QPTrim$(ReplaceString(TaxCustRec.CSSN, "-", ""))
      For y = 1 To 9
        ThatSSN(y) = Mid(MatSSN, y, 1)
        If ThatSSN(y) <> ThisSSN(y) Then
          ThatSSN(y) = "x"
        End If
        NewMatSSN = NewMatSSN + ThatSSN(y)
      Next y
      If Len(QPTrim$(ReplaceString(NewSSN, "x", ""))) = Len(QPTrim$(ReplaceString(NewMatSSN, "x", ""))) Then
        If InStr(NewSSN, NewMatSSN) > 0 Then
          FoundIt = True
        Else
          FoundIt = False
          GoTo NoMatch
       End If
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(SearchName)) > 0 Then
      If InStr(QPTrim$(UCase(TaxCustRec.SName)), QPTrim$(UCase(SearchName))) Then
        FoundIt = True
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(AcctNum)) > 0 Then
      If InStr(AcctNum, CStr(TaxCustRec.Acct)) And (Len(QPTrim$(CStr(TaxCustRec.Acct))) = Len(QPTrim$(AcctNum))) Then
        FoundIt = True
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(SrvcAdd)) > 0 Then
      If InStr(TaxCustRec.ServiceAdd, SrvcAdd) Then
        FoundIt = True
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(OptSrchFld)) > 0 Then
      If InStr(TaxCustRec.OptSrchDesc, OptSrchFld) Then
        FoundIt = True
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(RealSearch$)) > 0 Then
      If TaxCustRec.FirstPropRec > 0 Then
        NextRSRec = TaxCustRec.FirstPropRec
        Do While NextRSRec > 0
          Get RHandle, NextRSRec, RealPropRec
          If InStr(RealPropRec.OptSearch, RealSearch$) Then
            RSet PrintRealSrch = "O - " + QPTrim$(RealPropRec.OptSearch)
            RealSrchFoundIt = True
            DontPrintIt = False 'added 8/22/06
            GoTo PrintIt 'added 8/22/06
LoopRSAgain:
            DontPrintIt = True 'added 8/22/06
          End If
          NextRSRec = RealPropRec.NextRec
        Loop
        If RealSrchFoundIt = False Then
          FoundIt = False
          GoTo NoMatch
        Else
          FoundIt = True
        End If
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(RealPin)) > 0 Then
      If TaxCustRec.FirstPropRec > 0 Then
        NextRRec = TaxCustRec.FirstPropRec
        Do While NextRRec > 0
          Get RHandle, NextRRec, RealPropRec
          If InStr(RealPropRec.RealPin, RealPin) Then
            RSet PrintRProp = QPTrim$(RealPropRec.RealPin)
            RealFoundIt = True
            DontPrintIt = False 'added 8/22/06
            GoTo PrintIt 'added 8/22/06
LoopRAgain:
            DontPrintIt = True 'added 8/22/06
          End If
          NextRRec = RealPropRec.NextRec
        Loop
        If RealFoundIt = True Then
          FoundIt = True
        Else
          FoundIt = False
          If Len(QPTrim$(PersPin)) > 0 Then GoTo TryPers
          GoTo NoMatch
        End If
      Else
        FoundIt = False
        If Len(QPTrim$(PersPin)) > 0 Then GoTo TryPers
        GoTo NoMatch
      End If
    End If
TryPers:
    If Len(QPTrim$(PersSearch$)) > 0 Then 'added 8/22/06
      If TaxCustRec.FirstPersRec > 0 Then
        NextPSRec = TaxCustRec.FirstPersRec
        Do While NextPSRec > 0
          Get PHandle, NextPSRec, PersPropRec
          If InStr(PersPropRec.OptSearch, PersSearch$) Then
            RSet PrintPersSrch = "O - " + QPTrim$(PersPropRec.OptSearch)
            PersSrchFoundIt = True
            FoundIt = True
            DontPrintIt = False 'added 8/22/06
            GoTo PrintIt 'added 8/22/06
LoopPSAgain:
            DontPrintIt = True 'added 8/22/06
          End If
          NextPSRec = PersPropRec.NextRec
        Loop
        If PersSrchFoundIt = False Then
          FoundIt = False
          GoTo NoMatch
        Else
          FoundIt = True
        End If
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(PersPin)) > 0 Then
      If TaxCustRec.FirstPersRec > 0 Then
        NextPRec = TaxCustRec.FirstPersRec
        Do While NextPRec > 0
          Get PHandle, NextPRec, PersPropRec
          If InStr(PersPropRec.PropPin, PersPin) Then
            RSet PrintPProp = QPTrim$(PersPropRec.PropPin)
            PersFoundIt = True
            DontPrintIt = False 'added 8/22/06
            GoTo PrintIt 'added 8/22/06
LoopPAgain:
'            PersFoundIt = False 'added 8/22/06
            DontPrintIt = True 'added 8/22/06
          End If
          NextPRec = PersPropRec.NextRec
        Loop
        If PersFoundIt = True Then
          FoundIt = True
        Else
          FoundIt = False
          GoTo NoMatch
        End If
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If FoundIt = False Then GoTo NoMatch
    GCustNum = ThisCustRec
PrintIt:
    If DontPrintIt = True Then 'added 8/22/06
      DontPrintIt = False 'added 8/22/06
      GoTo NoMatch 'added 8/22/06
    End If 'added 8/22/06
    
    '------------------------------------------------------------------
    FoundMatch = FoundMatch + 1
    If Len(QPTrim$(TaxCustRec.CSSN)) > 0 Then
      Call ReplaceString(TaxCustRec.CSSN, "-", "") 'added 12/7/07
      Call InsertSSNDashes(TaxCustRec.CSSN)
    End If
    If PersFoundIt = True And RealFoundIt = True Then
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & PrintRProp
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & PrintPProp
    ElseIf PersFoundIt = True Then
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & PrintPProp
    ElseIf RealFoundIt = True Then
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & PrintRProp
    ElseIf RealSrchFoundIt = True Then
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & PrintRealSrch
    ElseIf PersSrchFoundIt = True Then
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & PrintPersSrch
    Else
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) '& Chr(9) & "NA"
    End If
    If NextRRec > 0 Then GoTo LoopRAgain
    If NextRSRec > 0 Then GoTo LoopRSAgain
    If NextPRec > 0 Then GoTo LoopPAgain
    If NextPSRec > 0 Then GoTo LoopPSAgain
NoMatch:
  Next x
  
  Close PHandle
  Close RHandle
  
  If FoundMatch > 1 Then
    fpList1.SetFocus
    fpList1.ListIndex = 0
  End If
  
  If FoundMatch = 1 Then
    If NumOfTaxCusts = 1 Then GCustNum = 1
    If EditCust = True Then
      frmVATaxCustAddEdit.Show
      frmVATaxCustAddEdit.Caption = "Customer Edit"
      frmVATaxCustAddEdit.Label1.Caption = "Customer Edit"
      If QPTrim$(fptxtRealPin.Text) <> "" And QPTrim$(fptxtPersPin.Text) = "" Then '8/16/06
        Call GoToRealPropScreen
      End If
      If QPTrim$(fptxtOptRealSrch.Text) <> "" And QPTrim$(fptxtOptPersSrch.Text) = "" Then '8/16/06
        Call GoToRealPropScreenOpt("SRCH")
      End If
      If QPTrim$(fptxtPersPin.Text) <> "" And QPTrim$(fptxtRealPin.Text) = "" Then '8/16/06
        Call GoToPersPropScreen
      End If
      If QPTrim$(fptxtOptPersSrch.Text) <> "" And QPTrim$(fptxtOptRealSrch.Text) = "" Then '8/16/06
        Call GoToPersPropScreenOpt("SRCH")
      End If
      DoEvents
    ElseIf THistRpt = True Then
      frmVATaxReportOpt.Show vbModal
      If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
        Unload frmVATaxReportOpt
        Call PrintGraphics
      ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
        frmVATaxMsg.Label1.Caption = "Pitch 12 is recommended for this report."
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Unload frmVATaxReportOpt
        Call PrintText
      End If
    ElseIf DelAbs = True Then
      frmVATaxAbsList.Show
      DoEvents
    ElseIf RPayEntry = True Then
      SwapCustNum2 = GCustNum 'the lookup changes the GCustNum  before
      'the TaxPayment save routine can save any changes...so we preserve
      'the old GCustNum long enough to check out the current customer before
      'loading the new one
      GCustNum = SwapCustNum1
      If frmVATaxPaymentEntry.Check4Changes = True Then
        Close
        Unload Me
        Exit Sub
      End If
      If Me.Visible = False Then Exit Sub
      frmVATaxPaymentEntry.GetNewCust = True
      Call frmVATaxPaymentEntry.Clearscreen
      frmVATaxPaymentEntry.NotFirstLoad = False
      GCustNum = SwapCustNum2
      Get TCHandle, GCustNum, TaxCustRec
      frmVATaxPaymentEntry.fpLongAcctNum.Text = TaxCustRec.Acct
      frmVATaxPaymentEntry.TempAcctNum = TaxCustRec.Acct
      frmVATaxPaymentEntry.fptxtName.Text = QPTrim$(TaxCustRec.CustName)
      If QPTrim$(TaxCustRec.Addr1) <> "" Then
        frmVATaxPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr1)
      Else
        frmVATaxPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr2)
      End If
      frmVATaxPaymentEntry.fptxtCity.Text = QPTrim$(TaxCustRec.City)
      frmVATaxPaymentEntry.fptxtState.Text = QPTrim$(TaxCustRec.State)
      frmVATaxPaymentEntry.fptxtZip.Text = QPTrim$(TaxCustRec.Zip)
      frmVATaxPaymentEntry.Lookup = False '2/14/06
      Call frmVATaxPaymentEntry.EnterEditChk
      Unload Me
    ElseIf PPayEntry = True Then
      SwapCustNum2 = GCustNum 'the lookup changes the GCustNum  before
      'the TaxPayment save routine can save any changes...so we preserve
      'the old GCustNum long enough to check out the current customer before
      'loading the new one
      GCustNum = SwapCustNum1
      If frmVATaxPersPaymentEntry.Check4Changes = True Then
        Close
        Unload Me
        Exit Sub
      End If
      If Me.Visible = False Then Exit Sub
      frmVATaxPersPaymentEntry.GetNewCust = True
      Call frmVATaxPersPaymentEntry.Clearscreen
      frmVATaxPersPaymentEntry.NotFirstLoad = False
      GCustNum = SwapCustNum2
      Get TCHandle, GCustNum, TaxCustRec
      frmVATaxPersPaymentEntry.fpLongAcctNum.Text = TaxCustRec.Acct
      frmVATaxPersPaymentEntry.TempAcctNum = TaxCustRec.Acct
      frmVATaxPersPaymentEntry.fptxtName.Text = QPTrim$(TaxCustRec.CustName)
      If QPTrim$(TaxCustRec.Addr1) <> "" Then
        frmVATaxPersPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr1)
      Else
        frmVATaxPersPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr2)
      End If
      frmVATaxPersPaymentEntry.fptxtCity.Text = QPTrim$(TaxCustRec.City)
      frmVATaxPersPaymentEntry.fptxtState.Text = QPTrim$(TaxCustRec.State)
      frmVATaxPersPaymentEntry.fptxtZip.Text = QPTrim$(TaxCustRec.Zip)
      frmVATaxPersPaymentEntry.Lookup = False '2/14/06
      Call frmVATaxPersPaymentEntry.EnterEditChk
      Unload Me
    ElseIf Exist("txAdjust.dat") Then
      Get TCHandle, GCustNum, TaxCustRec
      frmVATaxAdjustments.fpLongAcctNum.Text = TaxCustRec.Acct
      If TaxCustRec.Acct > 0 Then
        Call frmVATaxAdjustments.LoadMeEdit
        frmVATaxAdjustments.fptxtDate.SetFocus
      End If
      Unload Me
    ElseIf Exist("manualbill.dat") Then
      If GCustNum > 0 Then
        Call frmVATaxManualBillEntry.ClearBillFields
        Call frmVATaxManualBillEntry.Clearscreen
        OpenTaxManualBillFile TMHandle, NumOfTMRecs
        For x = 1 To NumOfTMRecs
          Get TMHandle, x, TaxMRec
          If TaxMRec.Deleted = True Then GoTo NoNo
          If TaxMRec.Account = GCustNum Then
            frmVATaxManualBillEntry.PostSaveLoad = True
            ThisMRec = 0
          End If
NoNo:
        Next x
        Close TMHandle
        Call frmVATaxManualBillEntry.EnterEditCheck
        DoEvents
        Unload Me
        If frmVATaxManualBillEntry.PostSaveLoad = True Then
          frmVATaxManualBillEntry.PostSaveLoad = False
        End If
      End If
    ElseIf Exist("C:\CPWork\custinq.dat") Then
      Call frmVATaxCustInq.LoadCust
'      frmVATaxCustInq.Show
      DoEvents
      Me.Hide
    ElseIf Exist("C:\CPWork\custtranshist.dat") Then
'      frmVATaxCustTHistRpt.fptxtName = QPTrim$(TaxCustRec.CustName)
      Call frmVATaxCustTHistRpt.LoadCust
      frmVATaxCustTHistRpt.Show
      DoEvents
      Unload Me
    End If
  End If
  
  If FoundMatch = 0 Then
    frmVATaxMsg.Label1.Caption = "No matches could be found."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  Close TCHandle
   
  Exit Sub
  
GetProp: '5/26/06
  PrintYN = False
  If OptAll.Value = True Then
    PrintYN = True
    Return
  End If
  If OptReal.Value = True Then
    If TaxCustRec.FirstPropRec > 0 Then
      PrintYN = True
    End If
  ElseIf OptPers.Value = True Then
    If TaxCustRec.FirstPersRec > 0 Then
      PrintYN = True
    End If
  ElseIf OptNoProp.Value = True Then
    If TaxCustRec.FirstPersRec = 0 And TaxCustRec.FirstPropRec = 0 Then
      PrintYN = True
    End If
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustLookup", "cmdSearch_Click", Erl)
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
  If KeyCode = vbKeyReturn Then
  'the next line was included to allow the user to have data
  'in the data fields and a selection in the list, use the
  'enter key as a way to process the selection
    If fpList1.ListIndex <> -1 Then GoTo CustAlreadySelected
    If Len(fptxtSSN1.Text) > 0 Or Len(fptxtSSN2.Text) > 0 Or Len(fptxtSSN3.Text) > 0 Or Len(fptxtSrvcAdd.Text) > 0 Or Len(fptxtAcctNum.Text) > 0 Or Len(fptxtSearchName.Text) > 0 Or Len(fptxtOptSearch.Text) > 0 Or Len(fptxtRealPin.Text) > 0 Or Len(fptxtPersPin) > 0 Or Len(fptxtOptRealSrch.Text) > 0 Then
      Call cmdSearch_Click
      KeyCode = 0
      Exit Sub
    End If
CustAlreadySelected:
    fpList1.Col = 1
    If QPTrim$(fpList1.ColText) = "" Then
      MsgBox "No customer has been selected"
      Exit Sub
    Else
      Call fpList1_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
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
      Call cmdSearch_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxCustLookup.")
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

Private Sub fpList1_DblClick()
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim x As Long, y As Long
  Dim SearchName$
  Dim SSNum$
  Dim PinNum$
  Dim City$
  Dim AcctNum$
  Dim Found As Boolean
  Dim FirstPersRec As Long
  Dim FirstRealRec As Long
  Dim SwapCustNum1 As Long
  Dim SwapCustNum2 As Long
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  
'  On Error GoTo ERRORSTUFF
  
  SwapCustNum1 = GCustNum
  
  fpList1.Col = 0
  SearchName$ = QPTrim$(fpList1.ColText)
  If SearchName$ = "" Then
    frmVATaxMsg.Label1.Caption = "No item has been selected"
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  
  fpList1.Col = 1
  City$ = QPTrim$(fpList1.ColText)
  
  fpList1.Col = 2
  SSNum = QPTrim$(fpList1.ColText)
  
  fpList1.Col = 3
  AcctNum = QPTrim$(fpList1.ColText)
  
  OpenTaxCustFile TCHandle, NumOfTaxCusts
'  For x = 1 To NumOfTaxCusts
'    Get TCHandle, x, TaxCustRec
'    If QPTrim$(TaxCustRec.CustName) = SearchName And _
'      Len(QPTrim$(TaxCustRec.CustName)) = Len(SearchName) And _
'        City$ = QPTrim$(TaxCustRec.City) And QPTrim$(ReplaceString(TaxCustRec.CSSN, "-", "")) = QPTrim$(ReplaceString(SSNum, "-", "")) _
'          And AcctNum = CStr(TaxCustRec.Acct) Then
  Get TCHandle, Val(AcctNum), TaxCustRec
    'If AcctNum = CStr(TaxCustRec.Acct) Then
      Found = True
      fpList1.Row = -1
      GCustNum = Val(AcctNum)
'      Exit For
    'Else
    '  Found = False
    '  GoTo NotAMatch
    'End If
NotAMatch:
  'Next x
  
  'Get TCHandle, GCustNum, TaxCustRec
  FirstPersRec = TaxCustRec.FirstPersRec
  FirstRealRec = TaxCustRec.FirstPropRec
  Close TCHandle
  
  If EditCust = True Then
    frmVATaxCustAddEdit.Show
    frmVATaxCustAddEdit.Caption = "Customer Edit"
    frmVATaxCustAddEdit.Label1.Caption = "Customer Edit"
    If QPTrim$(fptxtRealPin.Text) <> "" And QPTrim$(fptxtPersPin.Text) = "" Then '8/16/06
      Call GoToRealPropScreen
    End If
    If QPTrim$(fptxtOptRealSrch.Text) <> "" And QPTrim$(fptxtOptPersSrch.Text) = "" Then '8/16/06
      Call GoToRealPropScreenOpt("DC")
    End If
    If QPTrim$(fptxtPersPin.Text) <> "" And QPTrim$(fptxtRealPin.Text) = "" Then '8/16/06
      Call GoToPersPropScreen
    End If
    If QPTrim$(fptxtOptPersSrch.Text) <> "" And QPTrim$(fptxtOptRealSrch.Text) = "" Then '8/16/06
      Call GoToPersPropScreenOpt("DC")
    End If
    DoEvents
  ElseIf THistRpt = True Then
    frmVATaxReportOpt.Show vbModal
    If frmVATaxReportOpt.fptxtPrintType.Text = "Graphical" Then
      Unload frmVATaxReportOpt
      Call PrintGraphics
    ElseIf frmVATaxReportOpt.fptxtPrintType.Text = "Text" Then
      Unload frmVATaxReportOpt
      Call PrintText
    End If
  ElseIf DelAbs = True Then
    If frmVATaxAbsMaint.fptxtChoice.Text = "real" Then
      If FirstRealRec = 0 Then
        frmVATaxMsg.Label1.Caption = "No Real Abstracts to Delete!"
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Close
        Exit Sub
      End If
      frmVATaxAbsList.Label2.Caption = "Real Estate List"
    ElseIf frmVATaxAbsMaint.fptxtChoice.Text = "pers" Then
      If FirstPersRec = 0 Then
        frmVATaxMsg.Label1.Caption = "No Personal Abstracts to Delete!"
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Close
        Exit Sub
      End If
      frmVATaxAbsList.Label2.Caption = "Personal Property List"
    End If
    frmVATaxAbsList.Show
    DoEvents
  ElseIf RPayEntry = True Then
    SwapCustNum2 = GCustNum 'the lookup changes the GCustNum  before
    'the TaxPayment save routine can save any changes...so we preserve
    'the old GCustNum long enough to check out the current customer before
    'loading the new one
    GCustNum = SwapCustNum1
    If frmVATaxPaymentEntry.Check4Changes = True Then
      Close
      Unload Me
      Exit Sub
    End If
    If Me.Visible = False Then Exit Sub
    frmVATaxPaymentEntry.GetNewCust = True
    Call frmVATaxPaymentEntry.Clearscreen
    GCustNum = SwapCustNum2
    frmVATaxPaymentEntry.fpLongAcctNum.Text = TaxCustRec.Acct
    frmVATaxPaymentEntry.TempAcctNum = TaxCustRec.Acct
    frmVATaxPaymentEntry.fptxtName.Text = QPTrim$(TaxCustRec.CustName)
    If QPTrim$(TaxCustRec.Addr1) <> "" Then
      frmVATaxPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr1)
    Else
      frmVATaxPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr2)
    End If
    frmVATaxPaymentEntry.fptxtCity.Text = QPTrim$(TaxCustRec.City)
    frmVATaxPaymentEntry.fptxtState.Text = QPTrim$(TaxCustRec.State)
    frmVATaxPaymentEntry.fptxtZip.Text = QPTrim$(TaxCustRec.Zip)
    frmVATaxPaymentEntry.Lookup = False
    Call frmVATaxPaymentEntry.EnterEditChk
    frmVATaxPaymentEntry.cmdBills.SetFocus
    Unload Me
  ElseIf PPayEntry = True Then
    SwapCustNum2 = GCustNum 'the lookup changes the GCustNum  before
    'the TaxPayment save routine can save any changes...so we preserve
    'the old GCustNum long enough to check out the current customer before
    'loading the new one
    GCustNum = SwapCustNum1
    If frmVATaxPersPaymentEntry.Check4Changes = True Then
      Close
      Unload Me
      Exit Sub
    End If
    If Me.Visible = False Then Exit Sub
    frmVATaxPersPaymentEntry.GetNewCust = True
    Call frmVATaxPersPaymentEntry.Clearscreen
    GCustNum = SwapCustNum2
    frmVATaxPersPaymentEntry.fpLongAcctNum.Text = TaxCustRec.Acct
    frmVATaxPersPaymentEntry.TempAcctNum = TaxCustRec.Acct
    frmVATaxPersPaymentEntry.fptxtName.Text = QPTrim$(TaxCustRec.CustName)
    If QPTrim$(TaxCustRec.Addr1) <> "" Then
      frmVATaxPersPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr1)
    Else
      frmVATaxPersPaymentEntry.fptxtAddress.Text = QPTrim$(TaxCustRec.Addr2)
    End If
    frmVATaxPersPaymentEntry.fptxtCity.Text = QPTrim$(TaxCustRec.City)
    frmVATaxPersPaymentEntry.fptxtState.Text = QPTrim$(TaxCustRec.State)
    frmVATaxPersPaymentEntry.fptxtZip.Text = QPTrim$(TaxCustRec.Zip)
    frmVATaxPersPaymentEntry.Lookup = False
    Call frmVATaxPersPaymentEntry.EnterEditChk
    frmVATaxPersPaymentEntry.cmdBills.SetFocus
    Unload Me
  ElseIf Exist("C:\CPWork\txradjust.dat") Then
    Unload Me
    frmVATaxAdjustments.fpLongAcctNum.Text = TaxCustRec.Acct
    If TaxCustRec.Acct > 0 Then
      Call frmVATaxAdjustments.LoadMeEdit
      frmVATaxAdjustments.fptxtDate.SetFocus
    End If
    Unload Me
  ElseIf Exist("C:\CPWork\txpadjust.dat") Then
    Unload Me
    frmVATaxPAdjustments.fpLongAcctNum.Text = TaxCustRec.Acct
    If TaxCustRec.Acct > 0 Then
      Call frmVATaxPAdjustments.LoadMeEdit
      frmVATaxPAdjustments.fptxtDate.SetFocus
    End If
    Unload Me
  ElseIf Exist("C:\CPWork\rmanualbill.dat") Then
    If GCustNum > 0 Then
      Call frmVATaxManualBillEntry.ClearBillFields
      Call frmVATaxManualBillEntry.Clearscreen
      OpenTaxManualBillFile TMHandle, NumOfTMRecs
      For x = 1 To NumOfTMRecs
        Get TMHandle, x, TaxMRec
        If TaxMRec.Deleted = True Then GoTo NoNo
        If TaxMRec.Account = GCustNum Then
          frmVATaxManualBillEntry.PostSaveLoad = True
          ThisMRec = 0
        End If
NoNo:
      Next x
      Close TMHandle
      Call frmVATaxManualBillEntry.EnterEditCheck
      DoEvents
      Unload Me
      If frmVATaxManualBillEntry.PostSaveLoad = True Then
        frmVATaxManualBillEntry.PostSaveLoad = False
      End If
    Else
      Call TaxMsg(900, "The customer search failed. Loading aborted.")
      DoEvents
      Unload Me
    End If
  ElseIf Exist("C:\CPWork\pmanualbill.dat") Then
    If GCustNum > 0 Then
      Call frmVATaxPManualBillEntry.ClearBillFields
      Call frmVATaxPManualBillEntry.Clearscreen
      OpenTaxManualBillFile TMHandle, NumOfTMRecs
      For x = 1 To NumOfTMRecs
        Get TMHandle, x, TaxMRec
        If TaxMRec.Deleted = True Then GoTo NoNoP
        If TaxMRec.Account = GCustNum Then
          frmVATaxPManualBillEntry.PostSaveLoad = True
          ThisMRec = 0
        End If
NoNoP:
      Next x
      Close TMHandle
      Call frmVATaxPManualBillEntry.EnterEditCheck
      DoEvents
      Unload Me
      If frmVATaxPManualBillEntry.PostSaveLoad = True Then
        frmVATaxPManualBillEntry.PostSaveLoad = False
      End If
    Else
      Call TaxMsg(900, "The customer search failed. Loading aborted.")
      DoEvents
      Unload Me
    End If
  ElseIf Exist("C:\CPWork\custinq.dat") Then
    Call frmVATaxCustInq.LoadCust
    DoEvents
    Unload Me
  ElseIf Exist("C:\CPWork\custtranshist.dat") Then
    Call frmVATaxCustTHistRpt.LoadCust
    frmVATaxCustTHistRpt.Show
    DoEvents
    Unload Me
  End If
  frmVATaxPaymentEntry.NotFirstLoad = True
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustLookup", "fpList1_DblClick", Erl)
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

Private Sub fptxtSSN1_Change()
  If Len(fptxtSSN1.Text) = 3 Then
    fptxtSSN2.SetFocus
  End If
End Sub

Private Sub fptxtSSN2_Change()
  If Len(fptxtSSN2.Text) = 2 Then
    fptxtSSN3.SetFocus
  End If

End Sub

Private Sub fptxtSSN3_Change()
  If Len(fptxtSSN3.Text) = 4 Then
    fptxtSrvcAdd.SetFocus
  End If

End Sub

Private Sub LoadMe()
  Dim TaxSURec As TaxMasterType
  Dim TMHandle As Integer
  
  OptAll.Value = True '5/26/06
  If Exist("TAXSETUP.Dat") Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxSURec
    Close TMHandle
  End If
  If QPTrim$(TaxSURec.OptSrchCust) <> "" Then
    Label6.Caption = "Search By: " + QPTrim$(TaxSURec.OptSrchCust)
    fptxtOptSearch.Enabled = True
  Else
    Label6.Caption = "No Optional Search Saved"
    fptxtOptSearch.Enabled = False
  End If
  
  If QPTrim$(TaxSURec.OptSrchProp) <> "" Then
    Label10.Caption = "Search By: " + QPTrim$(TaxSURec.OptSrchProp)
    fptxtOptRealSrch.Enabled = True
  Else
    Label10.Caption = "No Real Opt'l Search Saved"
    fptxtOptRealSrch.Enabled = False
  End If
  
  If QPTrim$(TaxSURec.OptSrchPers) <> "" Then
    Label12.Caption = "Search By: " + QPTrim$(TaxSURec.OptSrchPers)
    fptxtOptPersSrch.Enabled = True
  Else
    Label12.Caption = "No Pers Opt'l Search Saved"
    fptxtOptPersSrch.Enabled = False
  End If
  
  LabelDel.Visible = False
  If DelAbs = True Then
    LabelDel.Visible = True
    If frmVATaxAbsMaint.fptxtChoice.Text = "real" Then
      LabelDel.Caption = "Delete Real Abstract"
    ElseIf frmVATaxAbsMaint.fptxtChoice.Text = "pers" Then
      LabelDel.Caption = "Delete Personal Abstract"
    End If
  End If
  
End Sub

Private Sub PrintGraphics()
  Dim TaxTran As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTrans As Long
  Dim TaxTran2 As TaxTransactionType
  Dim TTHandle2 As Integer
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfCusts As Long
  Dim DidCnt As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim PrevTranRec&
  Dim PCnt As Integer
  Dim ThisRec&
  Dim ZCnt As Integer
  Dim TOwed#
  Dim TPaid#
  Dim cnt As Long
  Dim TransDate$
  Dim BillType$
  Dim TaxYear$
  Dim Post2GL$
  Dim GTOwed#
  Dim GTPaid#
  Dim NextRec As Integer
  Dim ThisTransType$
  Dim ThisCust$
  Dim ThisAmtOwed#, thisAmtPaid#
  Dim dlm$, ThisPin$
  Dim Town$
  Dim TaxSURec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrnCnt As Long
  Dim OverAllBal As Double
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxSURec
  Close TMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  dlm = "~"
  RptFile$ = "TAXRPTS\TaxCHIST.RPT"     'Report File Name
  
  PrnCnt = 0
  DidCnt = 0
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenTaxCustFile TCHandle, NumOfCusts
  
  Get TCHandle, GCustNum, TaxCustRec
  Close TCHandle
  OverAllBal = GetCustBalance(GCustNum, -1)
  
  ThisCust = QPTrim$(TaxCustRec.CustName)
  
  OpenTaxTransFile TTHandle, NumOfTrans
  
  PrevTranRec& = TaxCustRec.LastTrans
  ReDim HistRecs(1 To 1) As HistRecInfoType
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get TTHandle, PrevTranRec&, TaxTran
      DidCnt = DidCnt + 1
      ReDim Preserve HistRecs(1 To DidCnt) As HistRecInfoType
      HistRecs(DidCnt).TranRec = PrevTranRec&
      HistRecs(DidCnt).TranType = TaxTran.TranType
      HistRecs(DidCnt).TranDate = TaxTran.TransDate
      HistRecs(DidCnt).BelongTo = TaxTran.BelongTo
      PrevTranRec& = TaxTran.LastTrans
    Loop
  End If
  
  For cnt = 1 To DidCnt
    NextRec = NextRec + 1
    If HistRecs(cnt).TranType = 1 Or HistRecs(cnt).TranType = 22 Or HistRecs(cnt).TranType = 12 Or HistRecs(cnt).TranType = 11 Then 'added 11 on 5/30/08
      Get TTHandle, HistRecs(cnt).TranRec, TaxTran
      GoSub GetTransInfo
      PrnCnt = PrnCnt + 1
      If HistRecs(cnt).TranType = 1 Or HistRecs(cnt).TranType = 22 Or HistRecs(cnt).TranType = 12 Or HistRecs(cnt).TranType = 11 Then
        '                    0            1            2              3
        Print #RptHandle, NextRec; dlm; Town; dlm; ThisCust; dlm; TransDate$;
        '                            4                5              6
        Print #RptHandle, dlm; ThisTransType; dlm; TaxYear$; dlm; Post2GL$;
        If HistRecs(cnt).TranType = 22 Then
          '                         7             8                       9                    10
          Print #RptHandle, dlm; 0; dlm; TaxTran.Amount; dlm; -TaxTran.Amount; dlm; OverAllBal; dlm; 'OldRound(GTOwed# - GTPaid#); dlm;
        ElseIf HistRecs(cnt).TranType = 12 Then
          '                         7             8                       9                    10
          Print #RptHandle, dlm; 0; dlm; -TaxTran.Amount; dlm; 0; dlm; OverAllBal; dlm; 'OldRound(GTOwed# - GTPaid#); dlm;
        ElseIf HistRecs(cnt).TranType = 11 Then
          '                         7             8                       9
          Print #RptHandle, dlm; 0; dlm; -TaxTran.Amount; dlm; 0; dlm; OverAllBal; dlm;
        Else
          '                         7             8                       9                    10
          Print #RptHandle, dlm; TOwed#; dlm; TPaid#; dlm; OldRound(TOwed# - TPaid#); dlm; OverAllBal; dlm; 'OldRound(GTOwed# - GTPaid#); dlm;
        End If
        If QPTrim$(TaxTran.PersPin) <> "0" And QPTrim$(TaxTran.PersPin) <> "" Then
          ThisPin = QPTrim$(TaxTran.PersPin)
          '                                11
          Print #RptHandle, "Personal Property Pin #:  " + ThisPin
        ElseIf QPTrim$(TaxTran.RealPin) <> "0" Then
          ThisPin = QPTrim$(TaxTran.RealPin)
          '                                11
          Print #RptHandle, "Real Property Pin #:  " + ThisPin
        Else
          ThisPin = "None Recorded"
          '                              11
          Print #RptHandle, "Property Pin #:  " + ThisPin
        End If
      End If
      ReDim THistRecs(1 To 1) As HistRecInfoType
      PCnt = 0
      ThisRec& = HistRecs(cnt).TranRec
      For ZCnt = 1 To DidCnt
        If HistRecs(ZCnt).TranType <> 1 And HistRecs(ZCnt).TranType <> 22 And HistRecs(ZCnt).TranType <> 12 Or HistRecs(cnt).TranType = 11 Then 'added 11 on 5/30/08
          If HistRecs(ZCnt).BelongTo = ThisRec& Then
            PCnt = PCnt + 1
            ReDim Preserve THistRecs(1 To PCnt) As HistRecInfoType
            LSet THistRecs(PCnt) = HistRecs(ZCnt)
            If HistRecs(ZCnt).TranType = 21 Then
              HistRecs(ZCnt).TranType = 0 'reset to 0 so it won't be run through this code again
            End If
          End If
        End If
      Next
      If PCnt > 0 Then
        For ZCnt = 1 To PCnt
          Get TTHandle, THistRecs(ZCnt).TranRec, TaxTran
            GoSub GetTransInfo
            PrnCnt = PrnCnt + 1
            Print #RptHandle, NextRec; dlm; Town; dlm; ThisCust; dlm; TransDate$;
            '                            4                5              6
            Print #RptHandle, dlm; ThisTransType; dlm; TaxYear$; dlm; Post2GL$;
            '                         7                    8                       9                         10
            Print #RptHandle, dlm; ThisAmtOwed#; dlm; thisAmtPaid#; dlm; OldRound(TOwed# - TPaid#); dlm; OverAllBal; dlm;
            If QPTrim$(TaxTran.PersPin) <> "" And QPTrim$(TaxTran.PersPin) <> "0" Then
              ThisPin = QPTrim$(TaxTran.PersPin)
             '                                11
              Print #RptHandle, "Personal Property Pin #:  " + ThisPin
            ElseIf QPTrim$(TaxTran.RealPin) <> "" Then
              ThisPin = QPTrim$(TaxTran.RealPin)
              '                                11
              Print #RptHandle, "Real Property Pin #:  " + ThisPin
            Else
              ThisPin = "None Recorded"
              '                              11
              Print #RptHandle, "Property Pin #:  " + ThisPin
            End If
        Next
      End If
    
    End If
  Next
  Close
  
  If PrnCnt = 0 Then
    Call TaxMsg(900, "There are no transactions saved for this customer.")
    Exit Sub
  End If
  
  arVATaxCustTransRpt.Show
  
  Exit Sub
  
PrintBillInfo:
  
Return

GetTransInfo:
  TransDate$ = MakeRegDate(TaxTran.TransDate)
  BillType$ = ""
  TaxYear$ = ""
  Post2GL$ = "N"
  If TaxTran.Posted2GL = "Y" Then
    Post2GL$ = "Y"
  End If
  ThisAmtOwed = 0
  thisAmtPaid = 0
  Select Case TaxTran.TranType
  Case 1
    BillType$ = "Bill #" + ParseBillNum$(TaxTran.Description)
    ThisTransType = BillType
    Select Case TaxTran.BillType
    Case "R"
      BillType$ = "Real-Estate"
    Case "P"
      BillType$ = "Personal Property"
    Case "C"
      BillType$ = "Combined"
    Case "M"
      BillType$ = "Manual"
    End Select
    TaxYear$ = QPTrim$(Str$(TaxTran.TaxYear))
    TPaid# = 0
    TOwed# = TaxTran.Amount
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
  Case 2
    ThisTransType = "Payment"
    TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt) '4/22/05
    thisAmtPaid = OldRound(TaxTran.Amount + TaxTran.DiscAmt) '4/22/05
    GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt) '4/22/05
  Case 3
    BillType$ = "Release"
    ThisTransType = "Release"
    TOwed# = OldRound#(TOwed# - TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
    ThisAmtOwed# = -TaxTran.Amount
  Case 4
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
    ThisTransType = "Interest"
  Case 5
    BillType$ = "Penalty"
    ThisTransType = "Penalty"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
  Case 6
    BillType$ = "Collection/Ad Cost"
    ThisTransType = "Collection/Ad Cost"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    ThisAmtOwed = TaxTran.Amount
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
  Case 7
    If TaxTran.CustPin = 0 Then 'DOS
      TPaid# = OldRound(TPaid# + TaxTran.Amount)
      thisAmtPaid = TaxTran.Amount
      GTPaid# = OldRound(GTPaid# + TaxTran.Amount)
      ThisTransType = "Adjust Paid 'DOS'"
    Else
      TPaid# = OldRound#(TPaid# - TaxTran.Amount)
      thisAmtPaid = -TaxTran.Amount
      GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      ThisTransType = "Adjust Paid Down"
    End If
  Case 8      'This will be the misc addcost adjustment
    BillType$ = "Miscellaneous Cost"
    ThisTransType = "Miscellaneous Cost"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    ThisAmtOwed = TaxTran.Amount
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
  Case 9
    BillType$ = "Credit Applied at Billing"
    ThisTransType = "Credit Applied at Billing"
    TPaid# = OldRound#(TPaid# + TaxTran.Revenue.PrePaidUsed)
    GTPaid# = OldRound#(GTPaid# + TaxTran.Revenue.PrePaidUsed) '9/21/06
    thisAmtPaid# = TaxTran.Revenue.PrePaidUsed
  Case 13
    BillType$ = "Adjust Bill Down"
    ThisTransType = "Adjust Bill Down"
    TOwed# = OldRound#(TOwed# - TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
    ThisAmtOwed# = -TaxTran.Amount
  Case 14
    BillType$ = "Adjust Bill Up"
    ThisTransType = "Adjust Bill Up"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
  Case 21
    BillType$ = "Payment Plus Overpayment"
    ThisTransType = "Payment Plus Overpayment"
    TPaid# = OldRound#(TPaid# + TaxTran.Amount)
    thisAmtPaid = TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
  Case 22
    BillType$ = "Overpayment Only"
    ThisTransType = "Overpayment Only"
    TPaid# = OldRound#(TPaid# + TaxTran.Amount)
    thisAmtPaid = TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
  Case 10
    BillType$ = "Adjust Pay Down Affecting Credit"
    ThisTransType = "Adjust Pay Down Affecting Credit"
    TPaid# = OldRound#(TPaid# - TaxTran.Amount)
    thisAmtPaid = -TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
  Case 11
    BillType$ = "Adjust Prepay Down"
    ThisTransType = "Adjust Prepay Down"
    TPaid# = OldRound#(TPaid# - TaxTran.Amount)
    thisAmtPaid = -TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
  Case 12
    BillType$ = "Refund Prepay"
    ThisTransType = "Refund Prepay"
    TPaid# = OldRound#(TPaid# - TaxTran.Amount)
    thisAmtPaid = -TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
  Case 24
    BillType$ = "Adjust Bill Up With Credit Applied"
    ThisTransType = "Adjust Bill Up With Credit Applied"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = OldRound(TaxTran.Amount)
  Case 30 'added 3/7/07
    BillType$ = "PPTRA Removal"
    ThisTransType = "PPTRA Removal"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
  Case Else
    BillType$ = "?????"
    ThisTransType = "Unknown"
  End Select
Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustLookup", "PrintGraphics", Erl)
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

ExitHistRpt:
  
End Sub

Private Sub PrintText()
  Dim TaxTran As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTrans As Long
  Dim TaxTran2 As TaxTransactionType
  Dim TTHandle2 As Integer
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfCusts As Long
  Dim DidCnt As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim PrevTranRec&
  Dim PCnt As Integer
  Dim ThisRec&
  Dim ZCnt As Integer
  Dim TOwed#
  Dim TPaid#
  Dim cnt As Long
  Dim TransDate$
  Dim BillType$
  Dim TaxYear$
  Dim Post2GL$
  Dim GTOwed#
  Dim GTPaid#, Page As Integer
  Dim NextRec As Integer
  Dim ThisTransType$, FF$
  Dim ThisCust$, MaxLines As Integer
  Dim ThisAmtOwed#, thisAmtPaid#
  Dim Town$, LineCnt As Integer
  Dim TaxSURec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrintIt As Boolean
  Dim GCnt As Integer
  Dim ThisPin$
  Dim PrnCnt As Long
  Dim OverAllBal As Double
  
  On Error GoTo ERRORSTUFF
  
  PrintIt = False
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxSURec
  Close TMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  PrnCnt = 0
  DidCnt = 0
  
  RptFile$ = "ARTxCusTRpt.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenTaxCustFile TCHandle, NumOfCusts
  
  Get TCHandle, GCustNum, TaxCustRec
  Close TCHandle
  
  ThisCust = QPTrim$(TaxCustRec.CustName)
  OverAllBal = GetCustBalance(GCustNum, -1)
  
  OpenTaxTransFile TTHandle, NumOfTrans
  
  PrevTranRec& = TaxCustRec.LastTrans
  ReDim HistRecs(1 To 1) As HistRecInfoType
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get TTHandle, PrevTranRec&, TaxTran
      DidCnt = DidCnt + 1
      ReDim Preserve HistRecs(1 To DidCnt) As HistRecInfoType
      HistRecs(DidCnt).TranRec = PrevTranRec&
      HistRecs(DidCnt).TranType = TaxTran.TranType
      HistRecs(DidCnt).TranDate = TaxTran.TransDate
      HistRecs(DidCnt).BelongTo = TaxTran.BelongTo
      PrevTranRec& = TaxTran.LastTrans
    Loop
  End If
  GoSub PrintHeader
  
  GoSub PrintCustHeader
  GCnt = 0
  For cnt = 1 To DidCnt
    If HistRecs(cnt).TranType = 1 Or HistRecs(cnt).TranType = 22 Or HistRecs(cnt).TranType = 12 Or HistRecs(cnt).TranType = 11 Then  'added 11 on 5/30/08
      Get TTHandle, HistRecs(cnt).TranRec, TaxTran
      PrintIt = True
      GCnt = GCnt + 1
      GoSub GetTransInfo
      If QPTrim$(TaxTran.PersPin) <> "0" Then
        ThisPin = QPTrim$(TaxTran.PersPin)
        Print #RptHandle, "Personal Property Pin #: " + ThisPin
      ElseIf QPTrim$(TaxTran.RealPin) <> "0" Then
        ThisPin = QPTrim$(TaxTran.RealPin)
        Print #RptHandle, "Real Property Pin #: " + ThisPin
      Else
        ThisPin = "None Recorded"
        Print #RptHandle, "Property Pin #: " + ThisPin
      End If
      PrnCnt = PrnCnt + 1
      Print #RptHandle, ThisTransType; Tab(21); TransDate$; Tab(34);
      If HistRecs(cnt).TranType = 22 Then
        Print #RptHandle, TaxYear$; Tab(39); Using$("$###,##0.00", 0); Tab(51);
        Print #RptHandle, Using$("$###,##0.00", TaxTran.Amount); Tab(78); Post2GL$
      ElseIf HistRecs(cnt).TranType = 12 Then
        Print #RptHandle, TaxYear$; Tab(39); Using$("$###,##0.00", 0); Tab(51);
        Print #RptHandle, Using$("$###,##0.00", -TaxTran.Amount); Tab(78); Post2GL$
      ElseIf HistRecs(cnt).TranType = 11 Then
        Print #RptHandle, TaxYear$; Tab(39); Using$("$###,##0.00", 0); Tab(51);
        Print #RptHandle, Using$("$###,##0.00", -TaxTran.Amount); Tab(78); Post2GL$
      Else
        Print #RptHandle, TaxYear$; Tab(39); Using$("$###,##0.00", ThisAmtOwed#); Tab(51);
        Print #RptHandle, Using$("$###,##0.00", thisAmtPaid#); Tab(78); Post2GL$
      End If
      LineCnt = LineCnt + 2
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      ReDim THistRecs(1 To 1) As HistRecInfoType
      PCnt = 0
      ThisRec& = HistRecs(cnt).TranRec
      For ZCnt = 1 To DidCnt
        If HistRecs(ZCnt).TranType <> 1 And HistRecs(ZCnt).TranType <> 22 And HistRecs(ZCnt).TranType <> 12 Or HistRecs(cnt).TranType = 11 Then 'added 11 on 5/30/08
          If HistRecs(ZCnt).BelongTo = ThisRec& Then
            PCnt = PCnt + 1
            ReDim Preserve THistRecs(1 To PCnt) As HistRecInfoType
            LSet THistRecs(PCnt) = HistRecs(ZCnt)
            If HistRecs(ZCnt).TranType = 21 Then
              HistRecs(ZCnt).TranType = 0 'reset to 0 so it won't be run through this code again
            End If
          End If
        End If
      Next
      If PCnt > 0 Then
        For ZCnt = 1 To PCnt
          Get TTHandle, THistRecs(ZCnt).TranRec, TaxTran
          GoSub GetTransInfo
          PrnCnt = PrnCnt + 1
          Print #RptHandle, ThisTransType; Tab(21); TransDate$; Tab(34);
          Print #RptHandle, TaxYear$; Tab(39); Using$("$###,##0.00", ThisAmtOwed); Tab(51);
          Print #RptHandle, Using$("$###,##0.00", thisAmtPaid#); Tab(78); Post2GL$
          LineCnt = LineCnt + 1
          If LineCnt > MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            GoSub PrintCustHeader
          End If
        Next
      End If
    End If
    If PrintIt = True Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      GoSub PrintCustSubEnd
      PrintIt = False
    End If
  Next
  
  If GCnt > 0 Then
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    GoSub PrintCustGrandEnd
  End If
  
  Print #RptHandle, FF$
  Close
  
  If PrnCnt = 0 Then
    Call TaxMsg(900, "There are no transactions saved for this customer.")
    Exit Sub
  End If
  
  ViewPrint RptFile$, "Tax Customer Transaction History", True
  
  KillFile RptFile$
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Tax Billing: Customer Transaction History"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Str(Page)
  Print #RptHandle,
  Print #RptHandle, Tab(34); "Tax"; Tab(75); "Posted"
  Print #RptHandle, "Transaction Type"; Tab(24); "Date"; Tab(34); "Year"; Tab(42); "Amt Owed"; Tab(54); "Amt Paid"; Tab(64); "Total Owed"; Tab(75); "To GL"
  Print #RptHandle, String$(80, "=")
  LineCnt = 7
Return

PrintCustHeader:
  Print #RptHandle, "Customer:  " + ThisCust
  Print #RptHandle, String$(80, "-")
  LineCnt = LineCnt + 2
  Return
  
PrintCustSubEnd:
  If HistRecs(cnt).TranType = 22 Then
    Print #RptHandle, Tab(15); "Totals"; Tab(39); Using$("$###,##0.00", 0); Tab(51); Using$("$###,##0.00", TaxTran.Amount); Tab(63); Using$("$###,##0.00", -TaxTran.Amount)
  ElseIf HistRecs(cnt).TranType = 12 Then
    Print #RptHandle, Tab(15); "Totals"; Tab(39); Using$("$###,##0.00", 0); Tab(51); Using$("$###,##0.00", -TaxTran.Amount); Tab(63); Using$("$###,##0.00", 0)
  ElseIf HistRecs(cnt).TranType = 11 Then
    Print #RptHandle, Tab(15); "Totals"; Tab(39); Using$("$###,##0.00", 0); Tab(51); Using$("$###,##0.00", -TaxTran.Amount); Tab(63); Using$("$###,##0.00", 0)
  Else
    Print #RptHandle, Tab(15); "Totals"; Tab(39); Using$("$###,##0.00", TOwed#); Tab(51); Using$("$###,##0.00", TPaid#); Tab(63); Using$("$###,##0.00", OldRound(TOwed# - TPaid#))
  End If
  Print #RptHandle, String$(80, "-")
  LineCnt = LineCnt + 2
  Return

PrintCustGrandEnd:
  Print #RptHandle, "Grand Totals"; Tab(39); Using$("$###,##0.00", GTOwed#); Tab(51); Using$("$###,##0.00", GTPaid#); Tab(63); Using$("$###,##0.00", OverAllBal) 'OldRound(GTOwed# - GTPaid#))
  Print #RptHandle, String$(80, "-")
  LineCnt = LineCnt + 2
  Return

GetTransInfo:
  TransDate$ = MakeRegDate(TaxTran.TransDate)
  BillType$ = ""
  TaxYear$ = ""
  Post2GL$ = "N"
  If TaxTran.Posted2GL = "Y" Then
    Post2GL$ = "Y"
  End If
  ThisAmtOwed = 0
  thisAmtPaid = 0
  Select Case TaxTran.TranType
  Case 1
    BillType$ = "Bill #" + ParseBillNum$(TaxTran.Description)
    ThisTransType = BillType
    Select Case TaxTran.BillType
    Case "R"
      BillType$ = "Real-Estate"
    Case "P"
      BillType$ = "Personal Property"
    Case "C"
      BillType$ = "Combined"
    Case "M"
      BillType$ = "Manual"
    End Select
    TaxYear$ = QPTrim$(Str$(TaxTran.TaxYear))
    TPaid# = 0
    TOwed# = TaxTran.Amount
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed = TaxTran.Amount
    thisAmtPaid = 0
  Case 2
    ThisTransType = "Payment"
    TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt) '4/22/05
    thisAmtPaid = OldRound(TaxTran.Amount + TaxTran.DiscAmt) '4/22/05
    GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt) '4/22/05
  Case 3
    BillType$ = "Release"
    ThisTransType = "Release"
    TOwed# = OldRound#(TOwed# - TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
    ThisAmtOwed# = -TaxTran.Amount
  Case 4
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
    ThisTransType = "Interest"
  Case 5
    BillType$ = "Penalty"
    ThisTransType = "Penalty"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
  Case 6
    BillType$ = "Collection/Ad Cost"
    ThisTransType = "Collection/Ad Cost"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    ThisAmtOwed = TaxTran.Amount
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
  Case 7 'may need to add discounts here
    If TaxTran.CustPin = 0 Then 'DOS
      TPaid# = OldRound(TPaid# + TaxTran.Amount)
      thisAmtPaid = TaxTran.Amount
      GTPaid# = OldRound(GTPaid# + TaxTran.Amount)
      ThisTransType = "Adjust Paid 'DOS'"
    Else
      TPaid# = OldRound#(TPaid# - TaxTran.Amount)
      thisAmtPaid = -TaxTran.Amount
      GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      ThisTransType = "Adjust Paid Down"
    End If
  Case 8      'This will be the misc addcost adjustment
    BillType$ = "Miscellaneous Cost"
    ThisTransType = "Miscellaneous Cost"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    ThisAmtOwed = TaxTran.Amount
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
  Case 9
    BillType$ = "Credit Applied at Billing"
    ThisTransType = "Credit Applied at Billing"
    TPaid# = OldRound#(TPaid# + TaxTran.Revenue.PrePaidUsed)
    GTPaid# = OldRound#(GTPaid# + TaxTran.Revenue.PrePaidUsed) '9/21/06
    thisAmtPaid# = TaxTran.Revenue.PrePaidUsed
  Case 13
    BillType$ = "Adjust Bill Down"
    ThisTransType = "Adjust Bill Down"
    TOwed# = OldRound#(TOwed# - TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
    ThisAmtOwed# = -TaxTran.Amount
  Case 14
    BillType$ = "Adjust Bill Up"
    ThisTransType = "Adjust Bill Up"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
  Case 21
    BillType$ = "Payment Plus Overpayment"
    ThisTransType = "Payment Plus Overpayment"
    TPaid# = OldRound#(TPaid# + TaxTran.Amount)
    thisAmtPaid = TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
  Case 22
    BillType$ = "Overpayment Only"
    ThisTransType = "Overpayment Only"
    TPaid# = OldRound#(TPaid# + TaxTran.Amount)
    thisAmtPaid = TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
  Case 10
    BillType$ = "Adjust Pay Down Affecting Credit"
    ThisTransType = "Adjust Pay Down Affecting Credit"
    TPaid# = OldRound#(TPaid# - TaxTran.Amount)
    thisAmtPaid = -TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
  Case 11
    BillType$ = "Adjust Prepay Down"
    ThisTransType = "Adjust Prepay Down"
    TPaid# = OldRound#(TPaid# - TaxTran.Amount)
    thisAmtPaid = -TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
  Case 12
    BillType$ = "Refund Prepay"
    ThisTransType = "Refund Prepay"
    TPaid# = OldRound#(TPaid# - TaxTran.Amount)
    thisAmtPaid = -TaxTran.Amount
    GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
  Case 24
    BillType$ = "Adjust Bill Up With Credit Applied"
    ThisTransType = "Adjust Bill Up With Credit Applied"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = OldRound(TaxTran.Amount)
  Case 30 'added 3/7/07
    BillType$ = "PPTRA Removal"
    ThisTransType = "PPTRA Removal"
    TOwed# = OldRound#(TOwed# + TaxTran.Amount)
    GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
    ThisAmtOwed# = TaxTran.Amount
  Case Else
    BillType$ = "?????"
    ThisTransType = "Unknown"
  End Select
Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustLookup", "PrintText", Erl)
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

  
ExitHistRpt:
  
End Sub
Public Sub Reopen()
  cmdExit.Enabled = True
  Me.KeyPreview = True
End Sub

Private Sub GoToRealPropScreen()
  Dim RealRec As PropertyRecType
  Dim NumOfRRecs As Long
  Dim RHandle As Integer
  Dim x As Long
  Dim NextRec As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim WhichRec As Integer
  Dim ThisPin$
  Dim CustName$
  Dim RealRecCnt As Integer
  
  If GCustNum = 0 Then Exit Sub
  
  If fpList1.SelCount > 0 Then
    fpList1.Row = fpList1.ListIndex
    fpList1.Col = 4
    ThisPin = QPTrim$(fpList1.ColText)
  Else
    ThisPin = QPTrim$(fptxtRealPin.Text)
  End If
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  NextRec = TaxCust.FirstPropRec
  ReDim CustRRecs(1 To 1) As Long
  Do While NextRec > 0
    Get RHandle, NextRec, RealRec
    If RealRec.Deleted <> 0 Then GoTo Skip
    RealRecCnt = RealRecCnt + 1
    ReDim Preserve CustRRecs(1 To RealRecCnt) As Long
    CustRRecs(RealRecCnt) = NextRec
Skip:
    NextRec = RealRec.NextRec
  Loop
  NextRec = TaxCust.FirstPropRec
  Do While NextRec > 0
    Get RHandle, NextRec, RealRec
    If RealRec.Deleted <> 0 Then
      GoTo SkipIt
    End If
    WhichRec = WhichRec + 1
    If QPTrim$(RealRec.RealPin) = ThisPin Then
      frmVATaxRealProp.WhichRec = WhichRec
      ReDim RealRecs(0 To 0) As Long
      Call GetRealRecList(RealRecs(), GCustNum, CustName)
      frmVATaxRealProp.fptxtThisCust.Text = CustName
      frmVATaxRealProp.NumOfCustRERecs = RealRecs(0)
      If RealRecs(0) <> RealRecCnt Then
        ReDim RealRecs(0 To 0) As Long
        Call TaxMsg(700, "ERROR: There is a problem reading the real property position. Please access this property through the customer screen.")
        Close
        Exit Sub
      End If
      Call frmVATaxRealProp.LoadAgain(WhichRec)
      frmVATaxRealProp.WhichRec = WhichRec
      Exit Do
    End If
SkipIt:
    NextRec = RealRec.NextRec
  Loop
  
  ReDim RealRecs(0 To RealRecCnt) As Long
  'ReDim RealRecs(0 To frmVATaxRealProp.NumOfCustRERecs) As Long
  RealRecs(0) = RealRecCnt
  'RealRecs(0) = frmVATaxRealProp.NumOfCustRERecs
  For x = 1 To RealRecs(0)
    RealRecs(x) = CustRRecs(x)
  Next x
  frmVATaxRealProp.Show
  
  
  Close RHandle

End Sub


Private Sub GoToPersPropScreen()
  Dim PersRec As PersonalRecType
  Dim NumOfPRecs As Long
  Dim PHandle As Integer
  Dim x As Long
  Dim NextRec As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim WhichRec As Integer
  Dim ThisPin$
  Dim CustName$
  Dim PersRecCnt As Integer
  
  If GCustNum = 0 Then Exit Sub
  
  If fpList1.SelCount > 0 Then
    fpList1.Row = fpList1.ListIndex
    fpList1.Col = 4
    ThisPin = QPTrim$(fpList1.ColText)
  Else
    ThisPin = QPTrim$(fptxtPersPin.Text)
  End If
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  NextRec = TaxCust.FirstPersRec
  ReDim CustPRecs(1 To 1) As Long
  Do While NextRec > 0
    Get PHandle, NextRec, PersRec
    If PersRec.Deleted <> 0 Then GoTo Skip
    PersRecCnt = PersRecCnt + 1
    ReDim Preserve CustPRecs(1 To PersRecCnt) As Long
    CustPRecs(PersRecCnt) = NextRec
Skip:
    NextRec = PersRec.NextRec
  Loop
  NextRec = TaxCust.FirstPersRec
  Do While NextRec > 0
    Get PHandle, NextRec, PersRec
    If PersRec.Deleted <> 0 Then
      GoTo SkipIt
    End If
    WhichRec = WhichRec + 1
    If QPTrim$(PersRec.PropPin) = ThisPin Then
      frmVATaxPersProp.WhichRec = WhichRec
      ReDim PersRecs(0 To 0) As Long
      Call GetPersRecList(PersRecs(), GCustNum, CustName)
      frmVATaxPersProp.fptxtThisCust.Text = CustName
      frmVATaxPersProp.NumOfCustPPRecs = PersRecs(0)
      If PersRecs(0) <> PersRecCnt Then
        ReDim PersRecs(0 To 0) As Long
        Call TaxMsg(700, "ERROR: There is a problem reading the personal property position. Please access this property through the customer screen.")
        Close
        Exit Sub
      End If
      Call frmVATaxPersProp.LoadAgain(WhichRec)
      frmVATaxPersProp.WhichRec = WhichRec
      Exit Do
    End If
    NextRec = PersRec.NextRec
SkipIt:
  Loop
  
  ReDim PersRecs(0 To frmVATaxPersProp.NumOfCustPPRecs) As Long
  PersRecs(0) = frmVATaxPersProp.NumOfCustPPRecs
  For x = 1 To PersRecs(0)
    PersRecs(x) = CustPRecs(x)
  Next x
  frmVATaxPersProp.Show
  
  Close PHandle

End Sub

Private Sub GoToRealPropScreenOpt(FromThis$)
  Dim RealRec As PropertyRecType
  Dim NumOfRRecs As Long
  Dim RHandle As Integer
  Dim x As Long
  Dim NextRec As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim WhichRec As Integer
  Dim ThisOpt$
  Dim CustName$
  Dim RealRecCnt As Integer
  Dim CompareThis$
  
  If GCustNum = 0 Then Exit Sub
  
  If fpList1.SelCount > 0 Then
    fpList1.Row = fpList1.ListIndex
    fpList1.Col = 4
    ThisOpt = QPTrim$(fpList1.ColText)
  Else
    ThisOpt = QPTrim$(fptxtOptRealSrch.Text)
  End If
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  NextRec = TaxCust.FirstPropRec
  ReDim CustRRecs(1 To 1) As Long
  Do While NextRec > 0
    Get RHandle, NextRec, RealRec
    If RealRec.Deleted <> 0 Then GoTo Skip
    RealRecCnt = RealRecCnt + 1
    ReDim Preserve CustRRecs(1 To RealRecCnt) As Long
    CustRRecs(RealRecCnt) = NextRec
    NextRec = RealRec.NextRec
Skip:
  Loop
  NextRec = TaxCust.FirstPropRec
  Do While NextRec > 0
    Get RHandle, NextRec, RealRec
    If RealRec.Deleted <> 0 Then
      GoTo SkipIt
    End If
    WhichRec = WhichRec + 1
    If FromThis = "DC" Then
      CompareThis = "O - " + QPTrim$(RealRec.OptSearch)
    ElseIf FromThis = "SRCH" Then
      CompareThis = QPTrim$(RealRec.OptSearch)
    End If
    If CompareThis = ThisOpt Then
      frmVATaxRealProp.WhichRec = WhichRec
      ReDim RealRecs(0 To 0) As Long
      Call GetRealRecList(RealRecs(), GCustNum, CustName)
      frmVATaxRealProp.fptxtThisCust.Text = CustName
      frmVATaxRealProp.NumOfCustRERecs = RealRecs(0)
      If RealRecs(0) <> RealRecCnt Then
        ReDim RealRecs(0 To 0) As Long
        Call TaxMsg(700, "ERROR: There is a problem reading the real property position. Please access this property through the customer screen.")
        Close
        Exit Sub
      End If
      Call frmVATaxRealProp.LoadAgain(WhichRec)
      frmVATaxRealProp.WhichRec = WhichRec
      Exit Do
    End If
SkipIt:
    NextRec = RealRec.NextRec
  Loop
  
'  Call frmVATaxRealProp.LoadMe
  ReDim RealRecs(0 To frmVATaxRealProp.NumOfCustRERecs) As Long
  RealRecs(0) = frmVATaxRealProp.NumOfCustRERecs
  For x = 1 To RealRecs(0)
    RealRecs(x) = CustRRecs(x)
  Next x
  frmVATaxRealProp.Show
  
  
  Close RHandle

End Sub

Private Sub GoToPersPropScreenOpt(FromThis$)
  Dim PersRec As PersonalRecType
  Dim NumOfPRecs As Long
  Dim PHandle As Integer
  Dim x As Long
  Dim NextRec As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim WhichRec As Integer
  Dim ThisOpt$
  Dim CustName$
  Dim PersRecCnt As Integer
  Dim CompareThis$
  
  If GCustNum = 0 Then Exit Sub
  
  If fpList1.SelCount > 0 Then
    fpList1.Row = fpList1.ListIndex
    fpList1.Col = 4
    ThisOpt = QPTrim$(fpList1.ColText)
  Else
    ThisOpt = QPTrim$(fptxtOptPersSrch.Text)
  End If
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  NextRec = TaxCust.FirstPersRec
  ReDim CustPRecs(1 To 1) As Long
  PersRecCnt = 0
  Do While NextRec > 0
    Get PHandle, NextRec, PersRec
    If PersRec.Deleted <> 0 Then GoTo Skip
    PersRecCnt = PersRecCnt + 1
    ReDim Preserve CustPRecs(1 To PersRecCnt) As Long
    CustPRecs(PersRecCnt) = NextRec
Skip:
    NextRec = PersRec.NextRec
  Loop
  NextRec = TaxCust.FirstPersRec
  Do While NextRec > 0
    Get PHandle, NextRec, PersRec
    If PersRec.Deleted <> 0 Then
      GoTo SkipIt
    End If
    WhichRec = WhichRec + 1
    If FromThis = "DC" Then
      CompareThis = "O - " + QPTrim$(PersRec.OptSearch)
    ElseIf FromThis = "SRCH" Then
      CompareThis = QPTrim$(PersRec.OptSearch)
    End If
    If CompareThis = ThisOpt Then
      frmVATaxPersProp.WhichRec = WhichRec
      ReDim PersRecs(0 To 0) As Long
      Call GetPersRecList(PersRecs(), GCustNum, CustName)
      frmVATaxPersProp.fptxtThisCust.Text = CustName
      frmVATaxPersProp.NumOfCustPPRecs = PersRecs(0)
      If PersRecs(0) <> PersRecCnt Then
        ReDim PersRecs(0 To 0) As Long
        Call TaxMsg(700, "ERROR: There is a problem reading the personal property position. Please access this property through the customer screen.")
        Close
        Exit Sub
      End If
      Call frmVATaxPersProp.LoadAgain(WhichRec)
      frmVATaxPersProp.WhichRec = WhichRec
      Exit Do
    End If
    NextRec = PersRec.NextRec
SkipIt:
  Loop
  
'  Call frmVATaxPersProp.LoadMe
  ReDim PersRecs(0 To frmVATaxPersProp.NumOfCustPPRecs) As Long
  PersRecs(0) = frmVATaxPersProp.NumOfCustPPRecs
  For x = 1 To PersRecs(0)
    PersRecs(x) = CustPRecs(x)
  Next x
  frmVATaxPersProp.Show
  
  Close PHandle

End Sub





