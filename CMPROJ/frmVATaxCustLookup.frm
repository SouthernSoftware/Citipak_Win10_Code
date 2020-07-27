VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
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
      TabIndex        =   12
      Tag             =   $"frmVATaxCustLookup.frx":08CA
      Top             =   4935
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
      Left            =   2760
      TabIndex        =   29
      Top             =   4536
      Width           =   1212
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
      Left            =   4200
      TabIndex        =   28
      Top             =   4536
      Width           =   1212
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
      Left            =   5640
      TabIndex        =   27
      Top             =   4536
      Width           =   1572
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
      Left            =   1800
      TabIndex        =   26
      Top             =   4536
      Width           =   732
   End
   Begin EditLib.fpText fptxtPersPin 
      Height          =   396
      Left            =   3696
      TabIndex        =   7
      Top             =   3840
      Width           =   3132
      _Version        =   196608
      _ExtentX        =   5524
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
      Height          =   396
      Left            =   3696
      TabIndex        =   6
      Top             =   3336
      Width           =   3132
      _Version        =   196608
      _ExtentX        =   5524
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
      Left            =   8160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3180
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   420
      Left            =   8160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3750
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmVATaxCustLookup.frx":0FB5
   End
   Begin EditLib.fpText fptxtOptSearch 
      Height          =   396
      Left            =   7800
      TabIndex        =   8
      Top             =   1700
      Width           =   2772
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
      Height          =   396
      Left            =   3216
      TabIndex        =   2
      Top             =   2352
      Width           =   3612
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
   Begin EditLib.fpText fptxtAcctNum 
      Height          =   396
      Left            =   3216
      TabIndex        =   1
      Top             =   1860
      Width           =   3612
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
      Height          =   396
      Left            =   2856
      TabIndex        =   0
      Top             =   1368
      Width           =   3972
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
      Height          =   396
      Left            =   3960
      TabIndex        =   3
      Top             =   2832
      Width           =   792
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
      Height          =   396
      Left            =   5040
      TabIndex        =   4
      Top             =   2832
      Width           =   468
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
      Height          =   396
      Left            =   5760
      TabIndex        =   5
      Top             =   2844
      Width           =   1068
      _Version        =   196608
      _ExtentX        =   1884
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
      Height          =   396
      Left            =   7800
      TabIndex        =   9
      Top             =   2520
      Width           =   2772
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
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   420
      Left            =   8160
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmVATaxCustLookup.frx":1193
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
      Height          =   252
      Left            =   720
      TabIndex        =   30
      Top             =   4296
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   720
      X2              =   7440
      Y1              =   4368
      Y2              =   4368
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   7440
      X2              =   7440
      Y1              =   4956
      Y2              =   2736
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
      Left            =   7680
      TabIndex        =   24
      Top             =   2160
      Width           =   2952
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3756
      Left            =   720
      Top             =   1200
      Width           =   10200
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
      TabIndex        =   23
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
      Left            =   1320
      TabIndex        =   22
      Top             =   1980
      Width           =   1872
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
      Left            =   1320
      TabIndex        =   21
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
      Left            =   1320
      TabIndex        =   20
      Top             =   2940
      Width           =   2592
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
      Left            =   1320
      TabIndex        =   19
      Top             =   1464
      Width           =   1512
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
      Left            =   7680
      TabIndex        =   18
      Top             =   1340
      Width           =   2952
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   1812
      Left            =   7440
      Top             =   1200
      Width           =   3480
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
      Left            =   4800
      TabIndex        =   17
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
      Left            =   5532
      TabIndex        =   16
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
      TabIndex        =   15
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
      Left            =   1320
      TabIndex        =   14
      Top             =   3444
      Width           =   2316
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
      Left            =   1320
      TabIndex        =   13
      Top             =   3960
      Width           =   2352
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
  fpList1.Clear
  OptAll.Value = True
  OptReal.Value = False
  OptPers.Value = False
  OptNoProp.Value = False
End Sub

Private Sub cmdExit_Click()
'  If DelAbs = True Then
'    'frmVATaxAbsMaint.Show
'    DoEvents
'  ElseIf RPayEntry = True Then
'    Unload Me
'  ElseIf PPayEntry = True Then
'    Unload Me
'  ElseIf Exist("txradjust.dat") Then
'    Unload Me
'  ElseIf Exist("txpadjust.dat") Then
'    Unload Me
'  ElseIf Exist("manualbill.dat") Then
'    KillFile "manualbill.dat"
'    'frmVATaxManualBillEntry.Show
'    DoEvents
'    Unload Me
'  ElseIf Exist("custinq.dat") Then
''    KillFile "custinq.dat"
'    DoEvents
'    Unload Me
'  ElseIf Exist("custtranshist.dat") Then
'    'frmVATaxCustTHistRpt.Show
'    DoEvents
'    Unload Me
'  Else
'  '  frmVATaxCustMaintMenu.Show
'    DoEvents
'  End If
'  KillFile "txpyment.dat"
' ' EditCust = False
' ' AddCust = False
'  THistRpt = False
'  DelAbs = False
'  RPayEntry = False
'  PPayEntry = False
  
  Unload Me
End Sub

Public Sub cmdSearch_Click()
  Dim TaxCustRec As VATaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim x As Long, y As Long, z As Long
  Dim SSN As String
  Dim SrvcAdd As String
  Dim AcctNum As String
  Dim SearchName As String
  Dim RealPin As String
  Dim PersPin As String
  Dim RealSearch As String
  Dim FoundIt As Boolean
  Dim MatSSN As String
  Dim FoundMatch As Integer
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
  Dim TaxMRec As VATaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  Dim PersPropRec As VAPersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealPropRec As VAPropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NextRec As Long
  Dim PersFoundIt As Boolean
  Dim RealFoundIt As Boolean
  Dim RealSrchFoundIt As Boolean
  Dim PrintRProp As String * 22
  Dim PrintPProp As String * 22
  Dim PrintRealSrch As String * 20
  Dim IdxFlag As Boolean
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxRec As TXCustNameIdxType
  Dim ThisCustRec As Long
  Dim PrintYN As Boolean '5/26/06
  On Local Error GoTo ERRORSTUFF
  OpenVAPersPropFile PHandle, NumOfPersRecs
  OpenVARealPropFile RHandle, NumOfRealRecs
  
  IdxFlag = False
  If QPTrim$(fptxtSearchName.Text) <> "" Then
    OpenVANameIdxFile IdxHandle, NumOfIdx
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
  OpenVATaxCustFile TCHandle, NumOfTaxCusts
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
  
  NewMatSSN = ""
  PrintAll = False
  If Len(QPTrim$(SSN)) = 0 And Len(QPTrim$(SearchName)) = 0 And Len(QPTrim$(AcctNum)) = 0 And Len(QPTrim$(SrvcAdd)) = 0 And Len(QPTrim$(OptSrchFld)) = 0 And Len(QPTrim$(RealPin)) = 0 And Len(QPTrim$(PersPin)) = 0 And Len(QPTrim$(RealSearch)) = 0 Then
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
'          GCustNum = x
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
'        GCustNum = x
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(AcctNum)) > 0 Then
      If InStr(AcctNum, CStr(TaxCustRec.Acct)) And (Len(QPTrim$(CStr(TaxCustRec.Acct))) = Len(QPTrim$(AcctNum))) Then
        FoundIt = True
'        GCustNum = x
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(SrvcAdd)) > 0 Then
      If InStr(TaxCustRec.ServiceAdd, SrvcAdd) Then
        FoundIt = True
'        GCustNum = x
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(OptSrchFld)) > 0 Then
      If InStr(TaxCustRec.OptSrchDesc, OptSrchFld) Then
        FoundIt = True
'        GCustNum = x
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(RealSearch$)) > 0 Then
      If TaxCustRec.FirstPropRec > 0 Then
        NextRec = TaxCustRec.FirstPropRec
        Do While NextRec > 0
          Get RHandle, NextRec, RealPropRec
          If InStr(RealPropRec.OptSearch, RealSearch$) Then
            RSet PrintRealSrch = "O - " + QPTrim$(RealPropRec.OptSearch)
            RealSrchFoundIt = True
            FoundIt = True
            Exit Do
          End If
          NextRec = RealPropRec.NextRec
        Loop
        If RealSrchFoundIt = False Then
          FoundIt = False
          GoTo NoMatch
        End If
      Else
        FoundIt = False
        GoTo NoMatch
      End If
    End If
    If Len(QPTrim$(RealPin)) > 0 Then
      If TaxCustRec.FirstPropRec > 0 Then
        NextRec = TaxCustRec.FirstPropRec
        Do While NextRec > 0
          Get RHandle, NextRec, RealPropRec
          If InStr(RealPropRec.RealPin, RealPin) Then
            RSet PrintRProp = "R - " + QPTrim$(RealPropRec.RealPin)
            RealFoundIt = True
            Exit Do
          End If
          NextRec = RealPropRec.NextRec
        Loop
        If RealFoundIt = True Then
          FoundIt = True
'          GCustNum = x
        Else
          FoundIt = False
          If Len(QPTrim$(PersPin)) > 0 Then GoTo TryPers
          GoTo NoMatch
        End If
      End If
    End If
TryPers:
    If Len(QPTrim$(PersPin)) > 0 Then
      If TaxCustRec.FirstPersRec > 0 Then
        NextRec = TaxCustRec.FirstPersRec
        Do While NextRec > 0
          Get PHandle, NextRec, PersPropRec
          If InStr(PersPropRec.PropPin, PersPin) Then
            RSet PrintPProp = "P - " + QPTrim$(PersPropRec.PropPin)
            PersFoundIt = True
            Exit Do
          End If
          NextRec = PersPropRec.NextRec
        Loop
        If PersFoundIt = True Then
          FoundIt = True
'          GCustNum = x
        Else
          FoundIt = False
          GoTo NoMatch
        End If
      End If
    End If
    If FoundIt = False Then GoTo NoMatch
    GCustNum = ThisCustRec
PrintIt:
    FoundMatch = FoundMatch + 1
    If Len(QPTrim$(TaxCustRec.CSSN)) > 0 Then
     ' Call InsertSSNDashes(TaxCustRec.CSSN)
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
    Else
      fpList1.InsertRow = QPTrim$(TaxCustRec.CustName) & Chr(9) & QPTrim$(TaxCustRec.City) & Chr(9) & QPTrim$(TaxCustRec.CSSN) & Chr(9) & CStr(TaxCustRec.Acct) & Chr(9) & "NA"
    End If
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
'    If EditCust = True Then
'      frmVATaxCustAddEdit.Show
'      frmVATaxCustAddEdit.Caption = "Customer Edit"
'      frmVATaxCustAddEdit.Label1.Caption = "Customer Edit"
'      DoEvents
    If THistRpt = True Then
      frmReportOpt.Show vbModal
      If rptopt = 1 Then
        Unload frmReportOpt
        Call PrintGraphics
      ElseIf rptopt = 2 Then
        frmTaxMsg.Label1.Caption = "Pitch 12 is recommended for this report."
        frmTaxMsg.Label1.Top = 900
        frmTaxMsg.Show vbModal
        Unload frmReportOpt
        Call PrintText
      End If
'    ElseIf DelAbs = True Then
'      frmVATaxAbsList.Show
'      DoEvents
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
      frmVATaxPaymentEntry.LookUp = False '2/14/06
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
      frmVATaxPersPaymentEntry.LookUp = False '2/14/06
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
'    ElseIf Exist("manualbill.dat") Then
'      If GCustNum > 0 Then
'        Call frmVATaxManualBillEntry.ClearBillFields
'        Call frmVATaxManualBillEntry.Clearscreen
'        OpenVATaxManualBillFile TMHandle, NumOfTMRecs
'        For x = 1 To NumOfTMRecs
'          Get TMHandle, x, TaxMRec
'          If TaxMRec.Deleted = True Then GoTo NoNo
'          If TaxMRec.Account = GCustNum Then
'            frmVATaxManualBillEntry.PostSaveLoad = True
'            ThisMRec = 0
'          End If
'NoNo:
'        Next x
'        Close TMHandle
'        Call frmVATaxManualBillEntry.EnterEditCheck
'        DoEvents
'        Unload Me
'        If frmVATaxManualBillEntry.PostSaveLoad = True Then
'          frmVATaxManualBillEntry.PostSaveLoad = False
'        End If
'      End If
'    ElseIf Exist("custinq.dat") Then
'      Call frmVATaxCustInq.LoadCust
'      frmVATaxCustInq.Show
'      DoEvents
'      Me.Hide
'    ElseIf Exist("custtranshist.dat") Then
''      frmVATaxCustTHistRpt.fptxtName = QPTrim$(TaxCustRec.CustName)
'      Call frmVATaxCustTHistRpt.LoadCust
'      frmVATaxCustTHistRpt.Show
'      DoEvents
'      Unload Me
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
 '   ClearInUse PWcnt
 '   CMTerminate
  
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
    fpList1.col = 1
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
      TXLog ("Cm.exe terminated via menu bar on frmVATaxCustLookup.")
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

Private Sub fpList1_DblClick()
  Dim TaxCustRec As VATaxCustType
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
  Dim TaxMRec As VATaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  
  On Local Error GoTo ERRORSTUFF
  
 'SwapCustNum1 = GCustNum
  
  fpList1.col = 0
  SearchName$ = QPTrim$(fpList1.ColText)
  If SearchName$ = "" Then
    frmVATaxMsg.Label1.Caption = "No item has been selected"
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Exit Sub
  End If
  
  fpList1.col = 1
  City$ = QPTrim$(fpList1.ColText)
  
  fpList1.col = 2
  SSNum = QPTrim$(fpList1.ColText)
  
  fpList1.col = 3
  AcctNum = QPTrim$(fpList1.ColText)
  
  OpenVATaxCustFile TCHandle, NumOfTaxCusts
  For x = 1 To NumOfTaxCusts
    Get TCHandle, x, TaxCustRec
    If QPTrim$(TaxCustRec.CustName) = SearchName And _
      Len(QPTrim$(TaxCustRec.CustName)) = Len(SearchName) And _
        City$ = QPTrim$(TaxCustRec.City) And QPTrim$(ReplaceString(TaxCustRec.CSSN, "-", "")) = QPTrim$(ReplaceString(SSNum, "-", "")) _
          And AcctNum = CStr(TaxCustRec.Acct) Then
      Found = True
      fpList1.Row = -1
      GCustNum = x
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
NotAMatch:
  Next x
  
  Get TCHandle, GCustNum, TaxCustRec
  FirstPersRec = TaxCustRec.FirstPersRec
  FirstRealRec = TaxCustRec.FirstPropRec
  Close TCHandle
  
  If THistRpt = True Then
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
    frmVATaxPaymentEntry.LookUp = False
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
    frmVATaxPersPaymentEntry.LookUp = False
    Call frmVATaxPersPaymentEntry.EnterEditChk
    frmVATaxPersPaymentEntry.cmdBills.SetFocus
    Unload Me
  ElseIf Exist("c:\CPWork\txradjust.dat") Then
    frmVATaxAdjustments.fpLongAcctNum.Text = TaxCustRec.Acct
    If TaxCustRec.Acct > 0 Then
      Call frmVATaxAdjustments.LoadMeEdit
      frmVATaxAdjustments.fptxtDate.SetFocus
    End If
    Unload Me
  ElseIf Exist("c:\CPWork\txpadjust.dat") Then
    frmVATaxPAdjustments.fpLongAcctNum.Text = TaxCustRec.Acct
    If TaxCustRec.Acct > 0 Then
      Call frmVATaxPAdjustments.LoadMeEdit
      frmVATaxPAdjustments.fptxtDate.SetFocus
    End If
    Unload Me
'  ElseIf Exist("rmanualbill.dat") Then
'    If GCustNum > 0 Then
'      Call frmVATaxManualBillEntry.ClearBillFields
'      Call frmVATaxManualBillEntry.Clearscreen
'      OpenVATaxManualBillFile TMHandle, NumOfTMRecs
'      For x = 1 To NumOfTMRecs
'        Get TMHandle, x, TaxMRec
'        If TaxMRec.Deleted = True Then GoTo NoNo
'        If TaxMRec.Account = GCustNum Then
'          frmVATaxManualBillEntry.PostSaveLoad = True
'          ThisMRec = 0
'        End If
'NoNo:
'      Next x
'      Close TMHandle
'      Call frmVATaxManualBillEntry.EnterEditCheck
''      Call frmVATaxManualBillEntry.EnterEditCheck
'      DoEvents
'      Unload Me
'      If frmVATaxManualBillEntry.PostSaveLoad = True Then
'        frmVATaxManualBillEntry.PostSaveLoad = False
'      End If
'    Else
'      Call VATaxMsg(900, "The customer search failed. Loading aborted.")
'      DoEvents
'      Unload Me
'    End If
'  ElseIf Exist("pmanualbill.dat") Then
'    If GCustNum > 0 Then
'      Call frmVATaxPManualBillEntry.ClearBillFields
'      Call frmVATaxPManualBillEntry.Clearscreen
'      OpenVATaxManualBillFile TMHandle, NumOfTMRecs
'      For x = 1 To NumOfTMRecs
'        Get TMHandle, x, TaxMRec
'        If TaxMRec.Deleted = True Then GoTo NoNoP
'        If TaxMRec.Account = GCustNum Then
'          frmVATaxPManualBillEntry.PostSaveLoad = True
'          ThisMRec = 0
'        End If
'NoNoP:
'      Next x
'      Close TMHandle
'      Call frmVATaxPManualBillEntry.EnterEditCheck
''      Call frmVATaxManualBillEntry.EnterEditCheck
'      DoEvents
'      Unload Me
'      If frmVATaxPManualBillEntry.PostSaveLoad = True Then
'        frmVATaxPManualBillEntry.PostSaveLoad = False
'      End If
'    Else
'      Call VATaxMsg(900, "The customer search failed. Loading aborted.")
'      DoEvents
'      Unload Me
'    End If
'  ElseIf Exist("custinq.dat") Then
'    Call frmVATaxCustInq.LoadCust
'    DoEvents
'    Unload Me
'  ElseIf Exist("custtranshist.dat") Then
''    frmVATaxCustTHistRpt.fptxtName = QPTrim$(TaxCustRec.CustName)
'    Call frmVATaxCustTHistRpt.LoadCust
'    frmVATaxCustTHistRpt.Show
'    DoEvents
'    Unload Me
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
 '   ClearInUse PWcnt
 '   CMTerminate
  
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
  Dim TaxSURec As VATaxMasterType
  Dim TMHandle As Integer
  OptAll.Value = True '5/26/06
  If Exist("TAXSETUP.Dat") Then
    OpenVATaxSetUpFile TMHandle
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
    Label10.Caption = "No Real Opt Search Saved"
    fptxtOptRealSrch.Enabled = False
  End If
'  fptxtOptSearchDesc.Text = QPTrim$(TaxSURec.OptSrchCust)
  LabelDel.Visible = False
'  If DelAbs = True Then
'    LabelDel.Visible = True
'    If frmVATaxAbsMaint.fptxtChoice.Text = "real" Then
'      LabelDel.Caption = "Delete Real Abstract"
'    ElseIf frmVATaxAbsMaint.fptxtChoice.Text = "pers" Then
'      LabelDel.Caption = "Delete Personal Abstract"
'    End If
'  End If
  
End Sub

Private Sub PrintGraphics()
  Dim TaxTran As VATaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTrans As Long
  Dim TaxTran2 As VATaxTransactionType
  Dim TTHandle2 As Integer
  Dim TaxCustRec As VATaxCustType
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
  Dim TaxSURec As VATaxMasterType
  Dim TMHandle As Integer
  Dim PrnCnt As Long
  
  On Local Error GoTo ERRORSTUFF
  
  OpenVATaxSetUpFile TMHandle
  Get TMHandle, 1, TaxSURec
  Close TMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  dlm = "~"
  RptFile$ = "TAXRPTS\TaxCHIST.RPT"     'Report File Name
  
  PrnCnt = 0
  DidCnt = 0
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenVATaxCustFile TCHandle, NumOfCusts
  
  Get TCHandle, GCustNum, TaxCustRec
  Close TCHandle
  ThisCust = QPTrim$(TaxCustRec.CustName)
  
  OpenVATaxTransFile TTHandle, NumOfTrans
  
  PrevTranRec& = TaxCustRec.LastTrans
  ReDim HistRecs(1 To 1) As VAHistRecInfoType
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get TTHandle, PrevTranRec&, TaxTran
      DidCnt = DidCnt + 1
      ReDim Preserve HistRecs(1 To DidCnt) As VAHistRecInfoType
      HistRecs(DidCnt).TranRec = PrevTranRec&
      HistRecs(DidCnt).TranType = TaxTran.TranType
      HistRecs(DidCnt).TranDate = TaxTran.TransDate
      HistRecs(DidCnt).BelongTo = TaxTran.BelongTo
      PrevTranRec& = TaxTran.LastTrans
    Loop
  End If
  
  For cnt = 1 To DidCnt
    NextRec = NextRec + 1
    If HistRecs(cnt).TranType = 1 Then
      Get TTHandle, HistRecs(cnt).TranRec, TaxTran
      GoSub GetTransInfo
      PrnCnt = PrnCnt + 1
      '                    0            1            2              3
      Print #RptHandle, NextRec; dlm; Town; dlm; ThisCust; dlm; TransDate$;
      '                            4                5              6
      Print #RptHandle, dlm; ThisTransType; dlm; TaxYear$; dlm; Post2GL$;
      '                         7             8                       9                            10
      Print #RptHandle, dlm; TOwed#; dlm; TPaid#; dlm; OldRound(TOwed# - TPaid#); dlm; OldRound(GTOwed# - GTPaid#); dlm;
      If QPTrim$(TaxTran.PersPin) <> "0" Then
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
      
      ReDim THistRecs(1 To 1) As VAHistRecInfoType
      PCnt = 0
      ThisRec& = HistRecs(cnt).TranRec
      For ZCnt = 1 To DidCnt
        If HistRecs(ZCnt).TranType <> 1 Then
          If HistRecs(ZCnt).BelongTo = ThisRec& Then
            PCnt = PCnt + 1
            ReDim Preserve THistRecs(1 To PCnt) As VAHistRecInfoType
            LSet THistRecs(PCnt) = HistRecs(ZCnt)
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
            '                         7                    8                       9                                10
            Print #RptHandle, dlm; ThisAmtOwed#; dlm; thisAmtPaid#; dlm; OldRound(TOwed# - TPaid#); dlm; OldRound(GTOwed# - GTPaid#); dlm;
            If QPTrim$(TaxTran.PersPin) <> "" Then
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
    Call VATaxMsg(900, "There are no transactions saved for this customer.")
    Exit Sub
  Else
    arVATaxCustTransRpt.Show
  End If
  
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
    GTPaid# = OldRound#(GTPaid# + TaxTran.Revenue.PrePaidUsed) '9/21/06 per Bob
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
  Dim TaxTran As VATaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTrans As Long
  Dim TaxTran2 As VATaxTransactionType
  Dim TTHandle2 As Integer
  Dim TaxCustRec As VATaxCustType
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
  Dim TaxSURec As VATaxMasterType
  Dim TMHandle As Integer
  Dim PrintIt As Boolean
  Dim GCnt As Integer
  Dim ThisPin$
  Dim PrnCnt As Long
  
  On Local Error GoTo ERRORSTUFF
  
  PrintIt = False
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  OpenVATaxSetUpFile TMHandle
  Get TMHandle, 1, TaxSURec
  Close TMHandle
  
  Town = QPTrim$(TaxSURec.Name)
  PrnCnt = 0
  DidCnt = 0
  
  RptFile$ = "ARTxCusTRpt.PRN"
  
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  OpenVATaxCustFile TCHandle, NumOfCusts
  
  Get TCHandle, GCustNum, TaxCustRec
  Close TCHandle
  
  ThisCust = QPTrim$(TaxCustRec.CustName)
  
  OpenVATaxTransFile TTHandle, NumOfTrans
  
  PrevTranRec& = TaxCustRec.LastTrans
  ReDim HistRecs(1 To 1) As VAHistRecInfoType
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get TTHandle, PrevTranRec&, TaxTran
      DidCnt = DidCnt + 1
      ReDim Preserve HistRecs(1 To DidCnt) As VAHistRecInfoType
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
    If HistRecs(cnt).TranType = 1 Then
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
      Print #RptHandle, TaxYear$; Tab(39); Using$("$###,##0.00", ThisAmtOwed#); Tab(51);
      Print #RptHandle, Using$("$###,##0.00", thisAmtPaid#); Tab(78); Post2GL$
      LineCnt = LineCnt + 2
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      ReDim THistRecs(1 To 1) As VAHistRecInfoType
      PCnt = 0
      ThisRec& = HistRecs(cnt).TranRec
      For ZCnt = 1 To DidCnt
        If HistRecs(ZCnt).TranType <> 1 Then
          If HistRecs(ZCnt).BelongTo = ThisRec& Then
            PCnt = PCnt + 1
            ReDim Preserve THistRecs(1 To PCnt) As VAHistRecInfoType
            LSet THistRecs(PCnt) = HistRecs(ZCnt)
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
    Call VATaxMsg(900, "There are no transactions saved for this customer.")
    Exit Sub
  Else
    ViewPrint RptFile$, "Tax Customer Transaction History", True
  End If
  
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
  Print #RptHandle, Tab(15); "Totals"; Tab(39); Using$("$###,##0.00", TOwed#); Tab(51); Using$("$###,##0.00", TPaid#); Tab(63); Using$("$###,##0.00", OldRound(TOwed# - TPaid#))
  Print #RptHandle, String$(80, "-")
  LineCnt = LineCnt + 2
  Return

PrintCustGrandEnd:
  Print #RptHandle, "Grand Totals"; Tab(39); Using$("$###,##0.00", GTOwed#); Tab(51); Using$("$###,##0.00", GTPaid#); Tab(63); Using$("$###,##0.00", OldRound(GTOwed# - GTPaid#))
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
    GTPaid# = OldRound#(GTPaid# + TaxTran.Revenue.PrePaidUsed) '9/21/06 per Bob
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




