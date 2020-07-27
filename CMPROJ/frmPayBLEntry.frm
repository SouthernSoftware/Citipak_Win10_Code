VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPayBLEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transaction Entry"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ForeColor       =   &H00000000&
   Icon            =   "frmPayBLEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   360
      Left            =   8928
      TabIndex        =   1
      Tag             =   $"frmPayBLEntry.frx":08CA
      Top             =   1968
      Width           =   2268
      _Version        =   196608
      _ExtentX        =   4000
      _ExtentY        =   635
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
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmPayBLEntry.frx":0A55
   End
   Begin LpLib.fpCombo fpcmbSetFlag 
      Height          =   360
      Left            =   10320
      TabIndex        =   4
      Tag             =   $"frmPayBLEntry.frx":0D4C
      Top             =   3936
      Width           =   876
      _Version        =   196608
      _ExtentX        =   1545
      _ExtentY        =   635
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
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmPayBLEntry.frx":0E5E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustList 
      Height          =   300
      Left            =   3690
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":1155
      Top             =   1590
      Width           =   1830
      _Version        =   131072
      _ExtentX        =   3228
      _ExtentY        =   529
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
      ButtonDesigner  =   "frmPayBLEntry.frx":125B
   End
   Begin EditLib.fpText fptxtName 
      Height          =   324
      Left            =   1704
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "This field contains the customer's business name. It cannot be edited."
      Top             =   1920
      Width           =   4812
      _Version        =   196608
      _ExtentX        =   8488
      _ExtentY        =   572
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
   Begin EditLib.fpText fptxtAddress 
      Height          =   324
      Left            =   1704
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "This field contains the primary address of this business. This field cannot be edited."
      Top             =   2256
      Width           =   4812
      _Version        =   196608
      _ExtentX        =   8488
      _ExtentY        =   572
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
   Begin EditLib.fpText fptxtCity 
      Height          =   324
      Left            =   1704
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "This field contains the name of the city where this business receives mail. This field cannot be edited."
      Top             =   2592
      Width           =   4812
      _Version        =   196608
      _ExtentX        =   8488
      _ExtentY        =   572
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
      MaxLength       =   20
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
   Begin EditLib.fpText fptxtAccount 
      Height          =   324
      Left            =   1704
      TabIndex        =   0
      Tag             =   $"frmPayBLEntry.frx":143F
      Top             =   1584
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
      _ExtentY        =   572
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
   Begin EditLib.fpText fptxtState 
      Height          =   324
      Left            =   1704
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "This field contains the state where this business receives mail. This field cannot be edited."
      Top             =   2928
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
      _ExtentY        =   572
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"
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
   Begin EditLib.fpMask fptxtZip 
      Height          =   324
      Left            =   3312
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   2928
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin EditLib.fpDateTime fptxtTDate 
      Height          =   348
      Left            =   8988
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":162F
      Top             =   984
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   614
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
      ButtonStyle     =   2
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
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
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
      Text            =   "11/20/2002"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
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
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpcurrLicBal 
      Height          =   324
      Index           =   0
      Left            =   3600
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "This field contains the total outstanding balance for license category #1. It is not editable."
      Top             =   3744
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrLicAmt 
      Height          =   324
      Index           =   0
      Left            =   5184
      TabIndex        =   6
      Tag             =   $"frmPayBLEntry.frx":16B8
      Top             =   3744
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   324
      Left            =   7296
      TabIndex        =   5
      Tag             =   $"frmPayBLEntry.frx":189F
      Top             =   4704
      Width           =   3900
      _Version        =   196608
      _ExtentX        =   6879
      _ExtentY        =   572
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
      MaxLength       =   20
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
      Height          =   396
      Left            =   9684
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":193B
      Top             =   7680
      Width           =   1368
      _Version        =   131072
      _ExtentX        =   2413
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmPayBLEntry.frx":1A37
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   396
      Left            =   8172
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":1C15
      Top             =   7680
      Width           =   1368
      _Version        =   131072
      _ExtentX        =   2413
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmPayBLEntry.frx":1EB0
   End
   Begin EditLib.fpCurrency fpcurrAmtDue1 
      Height          =   324
      Left            =   9600
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "This field contains the current outstanding balance for this customer. It cannot be edited."
      ToolTipText     =   "This is a read only field. It indicates the total outstanding balance for this customer."
      Top             =   1584
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrCashPaid 
      Height          =   324
      Left            =   9600
      TabIndex        =   2
      Tag             =   $"frmPayBLEntry.frx":208C
      Top             =   2400
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrChkPaid 
      Height          =   324
      Left            =   9600
      TabIndex        =   3
      Tag             =   $"frmPayBLEntry.frx":2154
      Top             =   2784
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrTotRecd 
      Height          =   324
      Left            =   9600
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":2257
      Top             =   3168
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrChange 
      Height          =   324
      Left            =   9600
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":2388
      Top             =   3552
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   572
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
      CurrencyNegFormat=   1
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
   Begin EditLib.fpCurrency fpcurrLicBal 
      Height          =   324
      Index           =   1
      Left            =   3600
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "This field contains the total outstanding balance for license category #2. It is not editable."
      Top             =   4080
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrLicBal 
      Height          =   324
      Index           =   2
      Left            =   3600
      TabIndex        =   46
      TabStop         =   0   'False
      Tag             =   "This field contains the total outstanding balance for license category #3. It is not editable."
      Top             =   4416
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrLicBal 
      Height          =   324
      Index           =   3
      Left            =   3600
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "This field contains the total outstanding balance for license category #4. It is not editable."
      Top             =   4752
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrLicBal 
      Height          =   324
      Index           =   4
      Left            =   3600
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "This field contains the total outstanding balance for license category #5. It is not editable."
      Top             =   5088
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrLicTotDue 
      Height          =   324
      Left            =   3600
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "This field contains the accumulated total of all outstanding license fees. This field cannot be edited."
      Top             =   5568
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrLicAmt 
      Height          =   324
      Index           =   1
      Left            =   5184
      TabIndex        =   7
      Tag             =   $"frmPayBLEntry.frx":24BD
      Top             =   4080
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrLicAmt 
      Height          =   324
      Index           =   2
      Left            =   5184
      TabIndex        =   8
      Tag             =   $"frmPayBLEntry.frx":26F5
      Top             =   4416
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrLicAmt 
      Height          =   324
      Index           =   3
      Left            =   5184
      TabIndex        =   9
      Tag             =   $"frmPayBLEntry.frx":292C
      Top             =   4752
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrLicAmt 
      Height          =   324
      Index           =   4
      Left            =   5184
      TabIndex        =   10
      Tag             =   $"frmPayBLEntry.frx":2B63
      Top             =   5088
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrLicTotPay 
      Height          =   324
      Left            =   5184
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":2D9A
      Top             =   5568
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
      ButtonColor     =   8421504
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpcurrPenTotDue 
      Height          =   324
      Left            =   3600
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "This field contains the outstanding penalty balance for this customer. This field is not editable."
      Top             =   6192
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
   Begin EditLib.fpCurrency fpcurrPenAmtTot 
      Height          =   324
      Left            =   5184
      TabIndex        =   11
      Tag             =   $"frmPayBLEntry.frx":2E5F
      Top             =   6192
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
   Begin EditLib.fpCurrency fpcurrRevGTDue 
      Height          =   324
      Left            =   3600
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   "This field contains the entire outstanding balance for this customer. This field is not editable. "
      Top             =   7056
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
   Begin EditLib.fpCurrency fpcurrRevGTPay 
      Height          =   324
      Left            =   5184
      TabIndex        =   58
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":2F58
      Top             =   7056
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency fpcurrIssDue 
      Height          =   324
      Left            =   3600
      TabIndex        =   60
      TabStop         =   0   'False
      Tag             =   "This field contains the current outstanding balance this customer has for issuance fees. This field is not editable."
      Top             =   6528
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
      Enabled         =   0   'False
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
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
   Begin EditLib.fpCurrency fpcurrIssAmt 
      Height          =   324
      Left            =   5184
      TabIndex        =   12
      Tag             =   $"frmPayBLEntry.frx":2FE0
      Top             =   6528
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   572
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCheck 
      Height          =   396
      Left            =   5148
      TabIndex        =   63
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":310A
      Top             =   7680
      Width           =   1368
      _Version        =   131072
      _ExtentX        =   2413
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmPayBLEntry.frx":321A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCharge 
      Height          =   396
      Left            =   6660
      TabIndex        =   64
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":33F6
      Top             =   7680
      Width           =   1368
      _Version        =   131072
      _ExtentX        =   2413
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmPayBLEntry.frx":3509
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInfo 
      Height          =   390
      Left            =   2130
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1380
      _Version        =   131072
      _ExtentX        =   2434
      _ExtentY        =   688
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
      ButtonDesigner  =   "frmPayBLEntry.frx":36E6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCash 
      Height          =   390
      Left            =   3630
      TabIndex        =   70
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":38C1
      Top             =   7680
      Width           =   1380
      _Version        =   131072
      _ExtentX        =   2434
      _ExtentY        =   688
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
      ButtonDesigner  =   "frmPayBLEntry.frx":39CD
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDrawer 
      Height          =   396
      Left            =   624
      TabIndex        =   71
      Top             =   7680
      Width           =   1368
      _Version        =   131072
      _ExtentX        =   2413
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmPayBLEntry.frx":3BA8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   435
      Left            =   7635
      TabIndex        =   72
      TabStop         =   0   'False
      Tag             =   "Press the 'Clear Payment Totals' button to restore all fields to their amounts as they appeared when the screen loaded initially."
      Top             =   6300
      Width           =   3195
      _Version        =   131072
      _ExtentX        =   5636
      _ExtentY        =   767
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
      ButtonDesigner  =   "frmPayBLEntry.frx":3D85
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdReset 
      Height          =   450
      Left            =   7635
      TabIndex        =   73
      TabStop         =   0   'False
      Tag             =   $"frmPayBLEntry.frx":3F72
      Top             =   5760
      Width           =   3195
      _Version        =   131072
      _ExtentX        =   5636
      _ExtentY        =   794
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
      ButtonDesigner  =   "frmPayBLEntry.frx":40AD
   End
   Begin VB.Shape Shape4 
      Height          =   612
      Left            =   168
      Top             =   7560
      Width           =   11364
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H8000000E&
      Height          =   6084
      Left            =   192
      Top             =   1464
      Width           =   11364
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
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
      Height          =   312
      Left            =   744
      TabIndex        =   69
      Top             =   1032
      Width           =   1824
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2688
      TabIndex        =   68
      Top             =   984
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   312
      Left            =   4704
      TabIndex        =   67
      Top             =   1032
      Width           =   1824
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6600
      TabIndex        =   66
      Top             =   984
      Width           =   732
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CM Business License Payment Entry"
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
      Left            =   2964
      TabIndex        =   65
      Top             =   360
      Width           =   5724
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   192
      X2              =   6864
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Set Renewal Flag (Y/N)?:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7200
      TabIndex        =   61
      Top             =   3984
      Width           =   3036
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Fee Revenue"
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
      Height          =   300
      Left            =   816
      TabIndex        =   59
      Top             =   6576
      Width           =   2076
   End
   Begin VB.Line Line3 
      X1              =   192
      X2              =   6864
      Y1              =   6048
      Y2              =   6048
   End
   Begin VB.Line Line2 
      X1              =   6912
      X2              =   11496
      Y1              =   5184
      Y2              =   5184
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   864
      TabIndex        =   56
      Top             =   7056
      Width           =   2028
   End
   Begin VB.Label CatDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   4
      Left            =   336
      TabIndex        =   54
      Top             =   5136
      UseMnemonic     =   0   'False
      Width           =   3228
   End
   Begin VB.Label CatDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   3
      Left            =   336
      TabIndex        =   53
      Top             =   4800
      UseMnemonic     =   0   'False
      Width           =   3228
   End
   Begin VB.Label CatDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   2
      Left            =   336
      TabIndex        =   52
      Top             =   4464
      UseMnemonic     =   0   'False
      Width           =   3228
   End
   Begin VB.Label CatDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   1
      Left            =   336
      TabIndex        =   51
      Top             =   4128
      UseMnemonic     =   0   'False
      Width           =   3228
   End
   Begin VB.Label CatDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   0
      Left            =   336
      TabIndex        =   50
      Top             =   3792
      UseMnemonic     =   0   'False
      Width           =   3228
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due Back:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7728
      TabIndex        =   44
      Top             =   3600
      Width           =   1836
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tendered:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7536
      TabIndex        =   43
      Top             =   3216
      Width           =   2028
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amount Tendered:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7104
      TabIndex        =   42
      Top             =   2832
      Width           =   2460
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Tendered:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7104
      TabIndex        =   41
      Top             =   2448
      Width           =   2460
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8016
      TabIndex        =   40
      Top             =   1632
      Width           =   1500
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7308
      TabIndex        =   39
      Top             =   1032
      Width           =   1596
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7344
      TabIndex        =   38
      Top             =   2064
      Width           =   1500
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   8544
      TabIndex        =   37
      Top             =   4368
      Width           =   1308
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2856
      TabIndex        =   35
      Top             =   2976
      Width           =   396
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   768
      TabIndex        =   34
      Top             =   2976
      Width           =   828
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   576
      TabIndex        =   33
      Top             =   1968
      Width           =   1020
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   816
      TabIndex        =   32
      Top             =   2640
      Width           =   636
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   480
      TabIndex        =   31
      Top             =   2304
      Width           =   1116
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   336
      TabIndex        =   30
      Top             =   1632
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   564
      Index           =   1
      Left            =   1464
      Top             =   240
      Width           =   8652
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid"
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
      Left            =   5232
      TabIndex        =   29
      Top             =   3456
      Width           =   1356
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total License Revenue"
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
      Height          =   300
      Left            =   480
      TabIndex        =   28
      Top             =   5568
      Width           =   2412
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Revenue"
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
      Height          =   300
      Left            =   816
      TabIndex        =   27
      Top             =   6240
      Width           =   2076
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   648
      Left            =   1464
      Top             =   180
      Width           =   8652
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due"
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
      Left            =   3696
      TabIndex        =   36
      Top             =   3456
      Width           =   1308
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   6888
      X2              =   6888
      Y1              =   1488
      Y2              =   7536
   End
End
Attribute VB_Name = "frmPayBLEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim NotFirstLoad As Boolean
  Dim TempTotRecd As Double
  Dim ThisDate$, Answer As Integer
  Dim TempAcctNum$
  Dim TempAcctBal As Double
  Dim TempLicBal() As Double
  Dim TempPenBal As Double
  Dim TempIssBal As Double
  Dim TempTotBal As Double
  Dim TempChkAmt As Double
  Dim TempCashAmt As Double
  Dim TempCreditAmt As Double
  Dim TempLicPaid() As Double
  Dim TempPenPaid As Double
  Dim TempIssPaid As Double
  Dim TempTotPaid As Double
  Dim TempTotDue As Double
  Dim TempChange As Double
  Dim TempPrintFlag As String * 1
  Dim CatDesc0$
  Dim CatDesc1$
  Dim CatDesc2$
  Dim CatDesc3$
  Dim CatDesc4$
  Dim NumOfCodes As Integer, uselook As Boolean
  Dim fromform As Form, toform As Form, codeopt As Integer, noreset As Boolean
  Dim TotNegBal As Double, RecpPort As String, CmNum As Long, cntout As Integer
  Dim NegFlag As Boolean, Oper As String, DefPayDate As String, RctValidate As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional DDate As String)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  uselook = True
  If DDate <> "" Then
    DefPayDate = DDate
  End If
End Sub
'Private Sub Form_Activate()
'  If Val(fpCustRecNo) > 0 And Not BeenDone Then
'    BeenDone = True
'    fpAcct = fpCustRecNo
'    GetCustinfo
'    DoEvents
'  End If
'End Sub

Private Sub cmdCash_Click()
  If fpcurrAmtDue1.DoubleValue > 0 Then
    fpcmbType.Text = "Cash"
    fpcurrCashPaid = fpcurrAmtDue1.DoubleValue
    fpCurrTotRecd = fpcurrAmtDue1.DoubleValue
    fptxtDesc.SetFocus
   ' Call cmdSave_Click
  End If
End Sub

Private Sub cmdCharge_Click()
  If fpcurrAmtDue1.DoubleValue > 0 Then
    fpcmbType.Text = "Charge"
    fpcurrChkPaid = fpcurrAmtDue1.DoubleValue
    fpCurrTotRecd = fpcurrAmtDue1.DoubleValue
    fptxtDesc.SetFocus
   ' Call cmdSave_Click
  End If

End Sub

Private Sub cmdCheck_Click()
  If fpcurrAmtDue1.DoubleValue > 0 Then
    fpcmbType.Text = "Check"
    fpcurrChkPaid = fpcurrAmtDue1.DoubleValue
    fpCurrTotRecd = fpcurrAmtDue1.DoubleValue
    fptxtDesc.SetFocus
  '  Call cmdSave_Click
  End If

End Sub

Private Sub cmdClear_Click()
  Dim x As Integer
    
  fpcurrRevGTPay = 0
  fpcurrLicTotPay = 0
  fpcurrPenAmtTot = 0
  fpcurrIssAmt = 0
  fpcurrChange = 0
  fpCurrTotRecd = 0
  fpcurrCashPaid = 0
  fpcurrChkPaid = 0
  For x = 0 To 4 'NumOfCodes - 1
    fpcurrLicAmt(x) = 0
  Next x
  fpcurrAmtDue1 = fpcurrRevGTDue.DoubleValue
  If fpcmbType.Enabled = True Then
    fpcmbType.SetFocus
  End If
End Sub

Private Sub cmdCustList_Click()
  frmBLCustomerList.Wheretogo frmPayBLEntry, frmPayBLEntry, 1
  frmBLCustomerList.Show vbModal
End Sub

Public Sub EnterEditChk()
  Dim ONum$
  Dim ThisRec As Integer
  Dim CustNum$
  
  'in conjunction with BegBalCheck this set of code determines the
  'current status of the customer the user is attempting to bring up
  'on the screen
  ONum = OperNum
  ThisRec = 0
  CustNum$ = Str(GCustNum)
  CustNum$ = QPTrim$(CustNum$)
'  Select Case BegBalCheck(CustNum$, ONum$, ThisRec)
'    Case 1 'normal first time transaction for this customer
'      EditFlag = False
      Call LoadMe
'      Exit Sub
'    Case 2 'edit a transaction that is in progress
'      EditFlag = True
'      GPayNum = ThisRec
'      Call LoadMe
'      Exit Sub
'    Case 3 'edit a transaction in progress started by
'    'a different operator ...(this one was deleted 12/31/03)
'      EditFlag = False
'      Call LoadMe
'      Exit Sub
'    Case 4 'a transaction for this customer is already in progress
'    'so abort this attempt
'      GCustNum = 0
'      EditFlag = False
'      Call LoadMe
'      Exit Sub
'    Case 5 'a transaction is in progress so don't edit it...rather
'    'start a brand new one ...(this one was deleted 12/31/03)
'      EditFlag = False
'      Call LoadMe
'      Exit Sub
'    Case Else
'      frmBLMessageBoxJr.Label1.Caption = "Error: This customer's data could not be retrieved."
'      frmBLMessageBoxJr.Label1.Top = 700
'      frmBLMessageBoxJr.Show vbModal
'      Close
'      Exit Sub
'  End Select
  
End Sub


Private Sub cmdExit_Click()
 ' EditFlag = False
  Chk4Change
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    cmdSave_Click
    If cntout > 0 Then
      Exit Sub
    End If
  End If

  ThisCustXNum = 0
  GCustNum = 0
  'KillFile "transentry.dat"
    Load frmCMPaySource
    DoEvents
    frmCMPaySource.Show
  
  BLLog ("Pay entry screen exited.")
  CMLog "OUT: CMBL Payment" + " Oper:" + Oper$
  Unload Me
  
End Sub
Private Sub Chk4Change()
  Answer = 0
  If fpCurrTotRecd <> 0 Or fpcurrRevGTPay <> 0 Then
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      Answer = 3
    Case True
      Answer = 2
    Case 1
      Answer = 1
    End Select
  Else
    Answer = 0
  End If
End Sub

''
''Private Sub cmdHelp_Click()
''  If InStr(cmdHelp.Text, "On") Then
''    cmdHelp.Text = "F1 &Turn Help Off"
''    btnHelp.AutoScan = fpAutoScanPopupOnly
''    lblBalloon.Visible = True
''    fptxtTDate.ToolTipText = ""
''    fptxtAccount.ToolTipText = ""
''    cmdInfo.ToolTipText = ""
''    cmdCustList.ToolTipText = ""
''    fptxtName.ToolTipText = ""
''    fptxtAddress.ToolTipText = ""
''    fptxtCity.ToolTipText = ""
''    fptxtState.ToolTipText = ""
''    fptxtZip.ToolTipText = ""
''    fpcurrAmtDue1.ToolTipText = ""
''    fpcurrCashPaid.ToolTipText = ""
''    fpcurrChkPaid.ToolTipText = ""
''    fpcmbType.ToolTipText = ""
''    fpcurrTotRecd.ToolTipText = ""
''    fpcurrChange.ToolTipText = ""
''    fptxtDesc.ToolTipText = ""
''    cmdClear.ToolTipText = ""
''    cmdReset.ToolTipText = ""
''    fpcurrRevGTDue.ToolTipText = ""
''    fpcurrRevGTPay.ToolTipText = ""
''    fpcurrPenTotDue.ToolTipText = ""
''    fpcurrPenAmtTot.ToolTipText = ""
''    fpcurrIssDue.ToolTipText = ""
''    fpcurrIssAmt.ToolTipText = ""
''    fpcurrLicTotDue.ToolTipText = ""
''    fpcurrLicTotPay.ToolTipText = ""
''    fpcurrLicBal(0).ToolTipText = ""
''    fpcurrLicBal(1).ToolTipText = ""
''    fpcurrLicBal(2).ToolTipText = ""
''    fpcurrLicBal(3).ToolTipText = ""
''    fpcurrLicBal(4).ToolTipText = ""
''    fpcurrLicAmt(0).ToolTipText = ""
''    fpcurrLicAmt(1).ToolTipText = ""
''    fpcurrLicAmt(2).ToolTipText = ""
''    fpcurrLicAmt(3).ToolTipText = ""
''    fpcurrLicAmt(4).ToolTipText = ""
''    cmdExit.ToolTipText = ""
''    cmdSave.ToolTipText = ""
''    cmdHelp.ToolTipText = ""
''    cmdDelete.ToolTipText = ""
''  ElseIf InStr(cmdHelp.Text, "Off") Then
''    cmdHelp.Text = "F1 &Turn Help On"
''    btnHelp.AutoScan = fpAutoScanOff
''    lblBalloon.Visible = False
'''    fptxtTDate.ToolTipText = "Today's date."
'''    fptxtAccount.ToolTipText = "Either enter a valid customer number here or select a customer number from the customer list brought up by pressing F7. Then pess F4 to populate this screen."
'''    cmdGetCust.ToolTipText = "Press this button to retrieve the data for the customer whose number is entered in the 'Account #' field."
'''    cmdCustList.ToolTipText = "Press this button to bring up a list of all currently saved customers."
'''    fptxtName.ToolTipText = "This is a read only field."
'''    fptxtAddress.ToolTipText = "This is a read only field."
'''    fptxtCity.ToolTipText = "This is a read only field."
'''    fptxtState.ToolTipText = "This is a read only field."
'''    fptxtZip.ToolTipText = "This is a read only field."
'''    fpcurrAmtDue1.ToolTipText = "This is a read only field. It indicates the total outstanding balance for this customer."
'''    fpcmbType.ToolTipText = "Select cash, check or charge depending on the payment form."
'''    fpcurrCashPaid.ToolTipText = "Enter the amount of cash the customer is remitting for this transaction."
'''    fpcurrChkPaid.ToolTipText = "Enter the check/charge amount this customer is remitting for this transaction."
'''    fpcurrTotRecd.ToolTipText = "This field tallies up the cash amount and the check amount tendered by this customer for this transaction."
'''    fpcurrChange.ToolTipText = "This field automatically calculates the difference between the amount tendered by the customer and the amount owed for this transaction."
'''    fptxtDesc.ToolTipText = "This field allows an optional comment regarding this transaction."
'''    cmdClear.ToolTipText = "Press to reset all the amounts entered to zero."
'''    cmdReset.ToolTipText = "Use this button if you have overridden automatic distribution and wish to have the program to redistribute the amounts entered."
'''    fpcurrRevGTDue.ToolTipText = "This field contains the total outstanding balance for this customer."
'''    fpcurrRevGTPay.ToolTipText = "Keeps a running total of amounts paid."
'''    fpcurrPenTotDue.ToolTipText = "Displays the outstanding penalty fee balance for this customer."
'''    fpcurrPenAmtTot.ToolTipText = "Enter the amount of payment earmarked for penalty fees."
'''    fpcurrIssDue.ToolTipText = "Displays the outstanding issuance fee balance for this customer."
'''    fpcurrIssAmt.ToolTipText = "Enter the amount of payment earmarked for issuance fees."
'''    fpcurrLicTotDue.ToolTipText = "This field contains the accumulated total of all outstanding license fees."
'''    fpcurrLicTotPay.ToolTipText = "This field contains the accumulated amounts paid for license fees."
'''    fpcurrLicBal(0).ToolTipText = "This field contains the total outstanding balance for license category #1."
'''    fpcurrLicBal(1).ToolTipText = "This field contains the total outstanding balance for license category #2."
'''    fpcurrLicBal(2).ToolTipText = "This field contains the total outstanding balance for license category #3."
'''    fpcurrLicBal(3).ToolTipText = "This field contains the total outstanding balance for license category #4."
'''    fpcurrLicBal(4).ToolTipText = "This field contains the total outstanding balance for license category #5."
'''    fpcurrLicAmt(0).ToolTipText = "Enter the amount earmarked for payment for license category #1 here."
'''    fpcurrLicAmt(1).ToolTipText = "Enter the amount earmarked for payment for license category #2 here."
'''    fpcurrLicAmt(2).ToolTipText = "Enter the amount earmarked for payment for license category #3 here."
'''    fpcurrLicAmt(3).ToolTipText = "Enter the amount earmarked for payment for license category #4 here."
'''    fpcurrLicAmt(4).ToolTipText = "Enter the amount earmarked for payment for license category #5 here."
'''    cmdExit.ToolTipText = "Press to exit this screen."
'''    cmdSave.ToolTipText = "Press to commmit the data on this screen to a temporary file. Any transaction can be edited until it is posted."
'''    cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
'''    cmdDelete.ToolTipText = "Remove this transaction from the pending transaction list."
''  End If
''
'''  frmBLMessageBox.Label1.Caption = "Business license automatically calculates the amount of change due to the customer. It also automatically distributes the amount tendered among the amounts owed for penalty, issuance fee and license fees (up to five separate license fees)."
'''  frmBLMessageBox.Label2.Caption = "Automatic distribution prioritizes penalty and issuance fees first and then begins with the first license fee owed and moves down the license list. Automatic distribution amounts, however, can be overridden. The program does require that any issuance fee or penalty fee is paid before license fees are paid."
'''  frmBLMessageBox.Label2.Height = 1500
'''  frmBLMessageBox.Label3.Caption = "Any transaction can be edited up until the time it is posted."
'''  frmBLMessageBox.Label3.Top = 3300
'''  frmBLMessageBox.Show vbModal
''End Sub

Private Sub cmdInfo_Click()
  If QPTrim$(fptxtAccount.Text) = "" Then
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
    Exit Sub
  End If
  
  If Check4ValidCustNum(QPTrim$(fptxtAccount.Text)) = True Then
    Call GetInfo
  Else
    frmBLMessageBoxJr.Label1.Caption = "The customer number entered is not valid. Please enter a valid customer number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Call Clearscreen
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
    Exit Sub
  End If
  
End Sub

Private Sub cmdInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call cmdInfo_Click
'  SkipLostFocus = True
End Sub

Private Sub cmdReset_Click()
  Call fpcurrTotRecd_Change
  If fpcmbType.Enabled = True Then
    fpcmbType.SetFocus
  End If
End Sub
'Public Function ChkOKFlag()
'Dim ThisNeg As Double, x As Integer, TotPaid As Double, TOTDIST As Double
'  If QPTrim$(fptxtAccount.Text) = "" Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  ThisNeg = OldRound(fpcurrTotRecd.DoubleValue - fpcurrRevGTPay.DoubleValue)
'  If ThisNeg < 0 Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  If OldRound(fpcurrIssDue.DoubleValue - fpcurrIssAmt.DoubleValue) < 0 Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  If OldRound(fpcurrPenTotDue.DoubleValue - fpcurrPenAmtTot.DoubleValue) < 0 Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  If fpcurrTotRecd.DoubleValue <= 0 Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  If fpcurrPenAmtTot.DoubleValue = 0 Then
'    If fpcurrIssAmt.DoubleValue = 0 Then
'      For x = 0 To 4
'        If fpcurrLicAmt(x).DoubleValue > 0 Then Exit For
'      Next x
'    End If
'  End If
'  If x > 4 And fpcurrTotRecd.DoubleValue > 0 Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  If QPTrim$(fpcmbType.Text) = "" Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'  If OldRound(TotPaid#) <> OldRound(TOTDIST#) Then
'    ChkOKFlag = False
'    Exit Function
'  End If
'
'End Function
Private Sub cmdSave_Click()

  Dim TotPaid#
  Dim AmtPaid#
  Dim Change#
  Dim TOTDIST#
  Dim NextRec As Integer
  Dim RctNum As Integer
  Dim ThisTot As Double
  Dim ThisNeg As Double
  Dim x As Integer
  Dim SaveFlag As Integer
  Dim FntSize As Integer, Msg1 As String, Msg2 As String
  Dim cnt As Integer
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  cntout = 0
 ' On Error GoTo ERRORSTUFF

  SaveFlag = 2
  'a valid account number is required to save the transaction
  If QPTrim$(fptxtAccount.Text) = "" Then
    fptxtAccount.BackColor = &HFFFF&
    Msg1$ = "Please enter a valid account number."
    fptxtAccount.BackColor = &HFFFFFF
    If fptxtAccount.Enabled = True Then
'      SkipLostFocus = False
      fptxtAccount.SetFocus
    End If
    cntout = cntout + 1
    GoTo msgstuff
    Exit Sub
  End If
  ThisNeg = OldRound(fpCurrTotRecd.DoubleValue - fpcurrRevGTPay.DoubleValue)
  If ThisNeg < 0 Then
    fpcurrChange.BackColor = &H80FFFF
    Msg1$ = "Payment amts entered exceed amt tendered"
    Msg2$ = "Please re-enter payment amounts correctly."
    DoEvents
    fpcurrChange.BackColor = &HFFFFFF
    cntout = cntout + 1
    GoTo msgstuff
    Exit Sub
  End If

  'looks to make sure the customer number matches the
  'customer name...the user could enter a different customer
  'number without fetching that data
'  If CompareAcctNumWData = False Then
'    Exit Sub
'  End If

  'issuance fees cannot be entered such that after posting the issuance
  'balance would be less than zero
  If OldRound(fpcurrIssDue.DoubleValue - fpcurrIssAmt.DoubleValue) < 0 Then
    fpcurrIssAmt.BackColor = &H80FFFF
    Msg1$ = "You May NOT Exceed Issuance Fee Amt Owed"
    Msg2$ = "Credit Issuance Fee Balance NOT Allowed."
    fpcurrIssAmt.BackColor = &HFFFFFF
    fpcurrIssAmt = fpcurrIssDue.DoubleValue
    If fpcurrIssAmt.Enabled = True Then
      fpcurrIssAmt.SetFocus
    End If
    'resets the issuance amount to the issuance balance
    ThisTot = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue
    fpcurrRevGTPay = OldRound(fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue)
    Call MakeChange
    cntout = cntout + 1
    GoTo msgstuff
    Exit Sub
  End If

  'penalty fees cannot be entered such that after posting the penalty
  'balance would be less than zero
  If OldRound(fpcurrPenTotDue.DoubleValue - fpcurrPenAmtTot.DoubleValue) < 0 Then
    fpcurrPenAmtTot.BackColor = &H80FFFF
    Msg1$ = "Invalid Penalty Amount"
    Msg2$ = "A Credit Penalty Balance NOT Allowed."
    fpcurrPenAmtTot.BackColor = &HFFFFFF
    fpcurrPenAmtTot = fpcurrPenTotDue.DoubleValue
    If fpcurrPenAmtTot.Enabled = True Then
      fpcurrPenAmtTot.SetFocus
    End If
    'the penalty amount is reset to the penalty balance
    ThisTot = OldRound(fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue)
    fpcurrRevGTPay = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue
    Call MakeChange
    cntout = cntout + 1
    GoTo msgstuff
    Exit Sub
  End If

  'if the total amount received entered is zero then the transaction is
  'automatically deleted
  'If EditFlag = True Then
    If fpCurrTotRecd.DoubleValue = 0 Then
      Msg1$ = "The total amount received is zero."
      Msg2$ = "Please Correct"
      cntout = cntout + 1
      GoTo msgstuff
      Exit Sub
    End If

'  Else 'User has not entered an amount in cash amount or check amount paid
    If fpCurrTotRecd.DoubleValue <= 0 Then
      fpCurrTotRecd.BackColor = &HFFFF&
'      If fpcurrRevGTPay.DoubleValue >= 0 Then 'total of amounts entered is more than zero
       Msg1$ = "Amount Received Must Be Greater Than Zero"
        fpCurrTotRecd.BackColor = &HFFFFFF
       Msg2$ = "Please Correct"
       cntout = cntout + 1
       GoTo msgstuff
        Exit Sub
     End If
'        frmBLSpecMsgBox.Label1.Top = 700 'total amounts tally to less than zero...
'        'indicates this is probably a transaction where the user is attempting to
'        'give the customer a refund
'        frmBLSpecMsgBox.Label1.Caption = "Since the amount received is zero the save procedure is aborted. If a refund is in order then please use the 'Adjust Customer Balance' screen for that procedure."
'        frmBLSpecMsgBox.Show vbModal
'        fpcurrTotRecd.BackColor = &HFFFFFF
'        Exit Sub
'      End If
'    End If
'  End If


  'this pop-up comes into play if the customer owes nothing but
  'wants to prepay
'  If fpcurrAmtDue1.DoubleValue = 0 Then
'    frmBLMessageBoxJrWOpts.Label1.Caption = "This customer has a zero balance. Do you want to continue saving anyway?"
'    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
'    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC No"
'    frmBLMessageBoxJrWOpts.Label1.Top = 800
'    frmBLMessageBoxJrWOpts.Show vbModal
'    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
'      Unload frmBLMessageBoxJrWOpts
'      Close
'      If fptxtAccount.Enabled = True Then
''        SkipLostFocus = False
'        fptxtAccount.SetFocus
'      End If
'      Exit Sub
'    Else
'      Unload frmBLMessageBoxJrWOpts
'    End If
'  End If

  'If the user enters an amount in the amount tendered field
  'then fails to put any amounts in any payment field then a
  'pop-up alerts the user that he cannot save the transaction
  'as entered
  If fpcurrPenAmtTot.DoubleValue = 0 Then
    If fpcurrIssAmt.DoubleValue = 0 Then
      For x = 0 To 4
        If fpcurrLicAmt(x).DoubleValue > 0 Then Exit For
      Next x
    End If
  End If
  If x > 4 And fpCurrTotRecd.DoubleValue > 0 Then
    Msg1$ = "The amount received must be distributed."
    Msg2$ = "Please distribute this amount before saving."
    cntout = cntout + 1
    GoTo msgstuff
    If fpCurrTotRecd.Enabled = True Then
      If fpCurrTotRecd.Enabled = True Then
        fpCurrTotRecd.SetFocus
      ElseIf fpcmbType.Enabled = True Then
        fpcmbType.SetFocus
      End If
    Else
      If fpcurrLicAmt(0).Enabled = True Then
        fpcurrLicAmt(0).SetFocus
      End If
    End If
    Exit Sub
  End If

  If TotalsOK = False Then Exit Sub

  AmtPaid# = fpCurrTotRecd.DoubleValue
  Change# = fpcurrChange.DoubleValue
  TOTDIST# = fpcurrRevGTPay.DoubleValue
  If fpcmbType.ListIndex = 2 Then
    If fpcurrCashPaid <= 0 Or fpcurrChkPaid <= 0 Then
      Msg1$ = "Cash and Check amounts required."
     cntout = cntout + 1
     GoTo msgstuff
     Exit Sub
    End If
  End If
  If QPTrim$(fpcmbType.Text) = "" Then
    fpcmbType.BackColor = &HFFFF&
    Msg1$ = "Please enter the type of payment method."
    cntout = cntout + 1
    GoTo msgstuff
    fpcmbType.BackColor = &HFFFFFF
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
    Exit Sub
  End If

  TotPaid# = OldRound(AmtPaid# - Change#)

  'can't save if the amount tendered does not add up to the totals
  'entered in the payment fields
  If OldRound(TotPaid#) <> OldRound(TOTDIST#) Then
    fpCurrTotRecd.BackColor = &HFFFF&
    fpcurrRevGTPay.BackColor = &HFFFF&
    If fpcmbType.Text = "Charge" Then
      Msg1$ = "Invalid Amounts."
      Msg2$ = "No change is allowed for charge payments."
    Else
      Msg1$ = "Invalid Amounts."
      Msg2$ = "Please correct before saving."
    End If
    fpCurrTotRecd.BackColor = &HFFFFFF
    cntout = cntout + 1
    GoTo msgstuff
    If fpcurrRevGTPay.Enabled = True Then
      fpcurrRevGTPay.SetFocus
      fpcurrRevGTPay.BackColor = &HFFFFFF
    ElseIf fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
    Exit Sub
  End If
    If fpcmbType.ListIndex = 1 Or fpcmbType.ListIndex = 2 Then
      frmPrintReceipt.setvallist = 1
    Else
      frmPrintReceipt.setvallist = 0
    End If

  frmPrintReceipt.Show vbModal
    If SavePay = True Then
      BLSaveTrans
      If PrnRecp = True Or PrnVali = True Then
        PrintReceipt
      End If
      MsgBox "Transaction Complete.", vbOKOnly, "Complete"
      Clearscreen
    End If

  
    
msgstuff:
  
  If cntout > 0 Then
    
    frmMsgDialog.RetLabel = "-2"
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:Entry Error"
    MsgText(1) = ""
    MsgText(2) = Msg1$
    MsgText(3) = Msg2$
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    
  End If

  
  
Exit Sub
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "cmdSave_Click", Erl)
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
  '  CitiTerminate

End Sub
Private Sub BLSaveTrans()

  Dim PayHandle As Integer

  Dim ListFile As Integer, CHandle As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  Dim NumOfRecs As Long, CMTrRecLen As Integer
  Dim cnt As Integer

  'ReDim PayRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)

  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"

  ReDim PayRec(1) As AREditPaymentRecType
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer
  Dim CustSrchIdxRec As CustSearchNameIdxType
  Dim SearchHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim CustRec(1) As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim TransRec As ARTransRecType
  Dim TransHandle As Integer
  Dim NumOfTransRecs As Long
  Dim NextTransRec As Long
  Dim CRec$
  Dim SCnt As Long
  Dim ARCode$
  Dim NewBalance#
  Dim AmtPaid#
  Dim TotPaid#
  Dim Change#
  Dim TOTDIST#
  Dim CustRecd As Integer
  Dim Prev&
  Dim ThisNumOfCodes As Integer
  Dim NextCode As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim LoopStop As Integer
  On Error GoTo ERRORSTUFF
  If Exist("artownsu.dat") Then
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
  End If
  AmtPaid# = fpCurrTotRecd.DoubleValue
  Change# = fpcurrChange.DoubleValue
  TOTDIST# = fpcurrRevGTPay.DoubleValue
  TotPaid# = OldRound(AmtPaid# - Change#)
  
  OpenBLCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec(1))
  
'  OpenPayFile PayHandle, OperNum 'file saved from TransEntry
'  NumOfPayRecs = LOF(PayHandle) / Len(PayRec)

'update temporary trans file
  PayRec(1).TranType = 20
  PayRec(1).SetFee = "N" 'old variable no longer used
  PayRec(1).ISSueFEE = 0 'old variable no longer used
  PayRec(1).AmtPaid = fpCurrTotRecd.DoubleValue
  PayRec(1).TranDate = Date2Num(fptxtTDate.Text)
  PayRec(1).CustNumber = QPTrim$(fptxtAccount.Text)
  PayRec(1).CustName = QPTrim$(fptxtName.Text)
  PayRec(1).Add1 = QPTrim$(fptxtAddress.Text)
  PayRec(1).City = QPTrim$(fptxtCity.Text)
  PayRec(1).State = QPTrim$(fptxtState.Text)
  PayRec(1).ZIPCODE = QPTrim$(fptxtZip.Text)

  If QPTrim$(fpcmbType.Text) = "Cash" Then
    PayRec(1).CASHCHK = "Cash"
  ElseIf QPTrim$(fpcmbType.Text) = "Check" Then
    PayRec(1).CASHCHK = "Check"
  ElseIf QPTrim$(fpcmbType.Text) = "Cash & Check" Then
    PayRec(1).CASHCHK = "Both"
  ElseIf QPTrim$(fpcmbType.Text) = "Charge" Then
    PayRec(1).CASHCHK = "Charge"
  Else
    PayRec(1).CASHCHK = "Not Saved"
  End If

  PayRec(1).CashAmt = fpcurrCashPaid.DoubleValue           'Cash Amount
  PayRec(1).ChkAmt = 0
  PayRec(1).CREDITAM = 0
  If QPTrim$(fpcmbType.Text) <> "Charge" Then
    PayRec(1).ChkAmt = fpcurrChkPaid.DoubleValue         'Cash Amount
  Else
    PayRec(1).CREDITAM = fpcurrChkPaid.DoubleValue
  End If
  PayRec(1).Change = fpcurrChange.DoubleValue
  If fpcmbSetFlag.Text = "Yes" Then
    PayRec(1).ISSUELIC = "Y"
  Else
    PayRec(1).ISSUELIC = "N"
  End If
  PayRec(1).Desc = QPTrim$(fptxtDesc.Text)
  PayRec(1).LICDUE = fpcurrLicTotDue.DoubleValue            'lic due
  PayRec(1).LICPAID = fpcurrLicTotPay.DoubleValue           'amt to lic
  PayRec(1).CatDesc1 = QPTrim$(CatDesc0)
  PayRec(1).CatDesc2 = QPTrim$(CatDesc1)
  PayRec(1).CatDesc3 = QPTrim$(CatDesc2)
  PayRec(1).CatDesc4 = QPTrim$(CatDesc3)
  PayRec(1).CatDesc5 = QPTrim$(CatDesc4)
  PayRec(1).LICDUE1 = fpcurrLicBal(0).DoubleValue
  PayRec(1).LICDUE2 = fpcurrLicBal(1).DoubleValue
  PayRec(1).LICDUE3 = fpcurrLicBal(2).DoubleValue
  PayRec(1).LICDUE4 = fpcurrLicBal(3).DoubleValue
  PayRec(1).LICDUE5 = fpcurrLicBal(4).DoubleValue
  PayRec(1).LICPAID1 = fpcurrLicAmt(0).DoubleValue
  PayRec(1).LICPAID2 = fpcurrLicAmt(1).DoubleValue
  PayRec(1).LICPAID3 = fpcurrLicAmt(2).DoubleValue
  PayRec(1).LICPAID4 = fpcurrLicAmt(3).DoubleValue
  PayRec(1).LICPAID5 = fpcurrLicAmt(4).DoubleValue
  PayRec(1).PENDUE = fpcurrPenTotDue.DoubleValue
  PayRec(1).PENPAID = fpcurrPenAmtTot.DoubleValue
  PayRec(1).TotDue = fpcurrAmtDue1.DoubleValue            'sum of (due)
  PayRec(1).TotPaid = fpcurrRevGTPay.DoubleValue         'sum of (paid)
  PayRec(1).Amount = TotPaid#
  PayRec(1).ISSDUE = fpcurrIssDue.DoubleValue
  PayRec(1).ISSPAID = fpcurrIssAmt.DoubleValue
'  If EditFlag = True Then
'    RctNum = GPayNum
'    Put PayHandle, GPayNum, EditPayRec 'opened already above
'  Else
'    EditFlag = True
  ListFile = FreeFile
  UBPayRecLen = Len(PayRec(1))
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    'NextRec = LOF(PayHandle) / Len(EditPayRec)
    'RctNum = NextRec + 1
    Put ListFile, 1, PayRec(1)
'  End If
  CMLog "Oper:" + Oper$ + " Updated TempFile for BL Pay"
'  frmBLMessageBoxJr.Label1.Top = 900
'  frmBLMessageBoxJr.Label1.Caption = "Your data has been successfully saved."
'  frmBLMessageBoxJr.Show vbModal
'  Call LogSaves(RctNum)
'
'  If SaveFlag = 2 Then
'    Call PrintReceipt(RctNum)
'    MainLog ("Receipt printed for " + QPTrim$(fptxtName.Text) + ".")
'  End If
'
'  GCustNum = 0
'  GPayNum = 0
'  Call LoadMe
'  If fptxtAccount.Enabled = True Then
''    SkipLostFocus = False
'    fptxtAccount.SetFocus
'  End If
'

  ReDim CMTrRec(1) As CMTransRecType

  ' See if any records to post
  OpenBLTransFile TransHandle
  NumOfTransRecs = LOF(TransHandle) / Len(TransRec)
  NextTransRec = NumOfTransRecs + 1
  
    ThisNumOfCodes = 1
    cnt = cnt + 1
    'Get PayHandle, cnt, PayRec
    CRec = Val(PayRec(1).CustNumber)
    If CRec > 0 Then
      Get CustHandle, Val(PayRec(1).CustNumber), CustRec(1)
      'Set New Balance
      NewBalance# = CustRec(1).AcctBal - PayRec(1).Amount

      ' Post Transaction Record First
      TransRec.CustomerNumber = PayRec(1).CustNumber
      TransRec.TransDate = PayRec(1).TranDate
      TransRec.TransAmount = PayRec(1).Amount
      TransRec.TransType = 2               ' Type 2 = Payment
      TransRec.TransDesc = "CM-" + QPTrim$(PayRec(1).Desc)
      TransRec.CashAmount = PayRec(1).Amount
      TransRec.ChkAmount = 0
      TransRec.BalanceAfterTrans = NewBalance#
      TransRec.ExtraRoom = ""
      TransRec.Posted2GL = "N"
      TransRec.CatCodeRec1 = GetCatRecNum(QPTrim$(CustRec(1).BILLCAT1)) 'CatRecord(1) 'record # for category code
      TransRec.CatCodeRec2 = GetCatRecNum(QPTrim$(CustRec(1).BILLCAT2)) 'CatRecord(2)
      TransRec.CatCodeRec3 = GetCatRecNum(QPTrim$(CustRec(1).BILLCAT3)) '
      TransRec.CatCodeRec4 = GetCatRecNum(QPTrim$(CustRec(1).BILLCAT4)) '
      TransRec.CatCodeRec5 = GetCatRecNum(QPTrim$(CustRec(1).BILLCAT5)) '
  'First five revs - store Cat rec no
  CMTrRec(1).TransRevAmt(1) = TransRec.CatCodeRec1
  CMTrRec(1).TransRevAmt(2) = TransRec.CatCodeRec2
  CMTrRec(1).TransRevAmt(3) = TransRec.CatCodeRec3
  CMTrRec(1).TransRevAmt(4) = TransRec.CatCodeRec4
  CMTrRec(1).TransRevAmt(5) = TransRec.CatCodeRec5
  '00000000000000000000
      If PayRec(1).LICPAID1 > 0 Or PayRec(1).LICPAID2 > 0 Or PayRec(1).LICPAID3 > 0 Or PayRec(1).LICPAID4 > 0 Or PayRec(1).LICPAID5 > 0 Or PayRec(1).ISSPAID > 0 Then
        TransRec.DetailTransType = 210
      End If
      TransRec.CatLicAmt1 = OldRound(PayRec(1).LICPAID1)
      TransRec.CatLicAmt2 = OldRound(PayRec(1).LICPAID2)
      TransRec.CatLicAmt3 = OldRound(PayRec(1).LICPAID3)
      TransRec.CatLicAmt4 = OldRound(PayRec(1).LICPAID4)
      TransRec.CatLicAmt5 = OldRound(PayRec(1).LICPAID5)
      TransRec.CatLicBal1 = OldRound(CustRec(1).FeeLicBal1 - PayRec(1).LICPAID1)
      TransRec.CatLicBal2 = OldRound(CustRec(1).FeeLicBal2 - PayRec(1).LICPAID2)
      TransRec.CatLicBal3 = OldRound(CustRec(1).FeeLicBal3 - PayRec(1).LICPAID3)
      TransRec.CatLicBal4 = OldRound(CustRec(1).FeeLicBal4 - PayRec(1).LICPAID4)
      TransRec.CatLicBal5 = OldRound(CustRec(1).FeeLicBal5 - PayRec(1).LICPAID5)
      TransRec.FeeAmt = 0
      TransRec.PenAmt = OldRound(PayRec(1).PENPAID)
      TransRec.IssAmt = OldRound(PayRec(1).ISSPAID)
      TransRec.LicAmt = OldRound(PayRec(1).LICPAID1 + PayRec(1).LICPAID2 + PayRec(1).LICPAID3 + PayRec(1).LICPAID4 + PayRec(1).LICPAID5)
      TransRec.IssBal = OldRound(PayRec(1).ISSDUE - PayRec(1).ISSPAID)
      TransRec.LicBal = OldRound(PayRec(1).LICDUE - PayRec(1).LICPAID)
      TransRec.PenBal = OldRound(PayRec(1).PENDUE - PayRec(1).PENPAID)
      If PayRec(1).PENPAID > 0 Then
        If TransRec.DetailTransType = 210 Then
          TransRec.DetailTransType = 211
        Else
          TransRec.DetailTransType = 201
        End If
      End If
      TransRec.NextTrans = 0
      Put TransHandle, NextTransRec, TransRec
      'Update Customer
      CustRecd = Val(PayRec(1).CustNumber)
      Get CustHandle, CustRecd, CustRec(1)
      CustRec(1).IssueLicense = PayRec(1).ISSUELIC
      CustRec(1).AcctBal = OldRound(CustRec(1).AcctBal - PayRec(1).Amount)
      CustRec(1).LicBal = OldRound(CustRec(1).LicBal - PayRec(1).LICPAID)
      CustRec(1).FeeLicBal1 = OldRound(CustRec(1).FeeLicBal1 - PayRec(1).LICPAID1)
      CustRec(1).FeeLicBal2 = OldRound(CustRec(1).FeeLicBal2 - PayRec(1).LICPAID2)
      CustRec(1).FeeLicBal3 = OldRound(CustRec(1).FeeLicBal3 - PayRec(1).LICPAID3)
      CustRec(1).FeeLicBal4 = OldRound(CustRec(1).FeeLicBal4 - PayRec(1).LICPAID4)
      CustRec(1).FeeLicBal5 = OldRound(CustRec(1).FeeLicBal5 - PayRec(1).LICPAID5)
      CustRec(1).FeeLicPay1 = PayRec(1).LICPAID1
      CustRec(1).FeeLicPay2 = PayRec(1).LICPAID2
      CustRec(1).FeeLicPay3 = PayRec(1).LICPAID3
      CustRec(1).FeeLicPay4 = PayRec(1).LICPAID4
      CustRec(1).FeeLicPay5 = PayRec(1).LICPAID5
      CustRec(1).PenBal = OldRound(CustRec(1).PenBal - PayRec(1).PENPAID)
      CustRec(1).IssuanceFee = PayRec(1).ISSueFEE
      CustRec(1).IssuanceBal = OldRound(CustRec(1).IssuanceBal - PayRec(1).ISSPAID)
      CustRec(1).IssuancePay = PayRec(1).ISSPAID
      If PayRec(1).SetFee = "Y" Then
        CustRec(1).FeeAmt = PayRec(1).Amount
      End If

      Put CustHandle, CustRecd, CustRec(1)

      If CustRec(1).FirstTrans = 0 Then
        CustRec(1).FirstTrans = NextTransRec
        CustRec(1).LastTrans = NextTransRec
        Put CustHandle, CustRecd, CustRec(1)
      Else
        Prev& = CustRec(1).LastTrans
        CustRec(1).LastTrans = NextTransRec
        Put CustHandle, CustRecd, CustRec(1)
        Get TransHandle, Prev&, TransRec
        TransRec.NextTrans = NextTransRec
        Put TransHandle, Prev&, TransRec
      End If
      'NextTransRec = NextTransRec + 1
    End If
  Close CustHandle
  Close TransHandle
  CMLog ("CMOper " + Str$(OperNum) + " BL Pay posted for Acct# " + Str$(CustRecd))
  CMTrRecLen = Len(CMTrRec(1))
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CHandle Len = CMTrRecLen
  'CmNum = (LOF(CHandle) \ CMTrRecLen) + 1
  
  If Len(QPTrim$(PayRec(1).Desc)) = 0 Then
    CMTrRec(1).TransDesc = "Business License Payment"
  Else
    CMTrRec(1).TransDesc = (QPTrim$(PayRec(1).Desc))
  End If
  Select Case QPTrim$(fpcmbType.Text)
    Case "Cash":
      CMTrRec(1).TransTender = 1
    Case "Check":
      CMTrRec(1).TransTender = 2
    Case "Cash & Check":
      CMTrRec(1).TransTender = 3
    Case "Charge":
      CMTrRec(1).TransTender = 4
    Case Else:
  End Select
  CMTrRec(1).TransCash = PayRec(1).CashAmt
  CMTrRec(1).TransDate = PayRec(1).TranDate
  CMTrRec(1).TransAmount = PayRec(1).TotPaid
  If PayRec(1).ChkAmt <> 0 Then
    CMTrRec(1).TransCheck = PayRec(1).ChkAmt
  Else
    CMTrRec(1).TransCheck = PayRec(1).CREDITAM
  End If
  CMTrRec(1).TransAmtOwed = PayRec(1).TotDue
  CMTrRec(1).TransDesc = QPTrim$(PayRec(1).Desc)
  CMTrRec(1).TransSource = 141
  CMTrRec(1).TransName = PayRec(1).CustName
  CMTrRec(1).TransAcctNum = PayRec(1).CustNumber
  CMTrRec(1).TransDetNum = NextTransRec
  CMTrRec(1).TransOperNum = OperNum
  CMTrRec(1).TransPad = ""
  'First five revs - store Cat rec no
  'do these above with the BL Transaction update
'''  CMTrRec(1).TransRevAmt(1) = PayRec(1).CatDesc1
'''  CMTrRec(1).TransRevAmt(2) = PayRec(1).CatDesc2
'''  CMTrRec(1).TransRevAmt(3) = PayRec(1).CatDesc3
'''  CMTrRec(1).TransRevAmt(4) = PayRec(1).CatDesc4
'''  CMTrRec(1).TransRevAmt(5) = PayRec(1).CatDesc5
  '6 to 10 revs store amt paid
  CMTrRec(1).TransRevAmt(6) = PayRec(1).LICPAID1
  CMTrRec(1).TransRevAmt(7) = PayRec(1).LICPAID2
  CMTrRec(1).TransRevAmt(8) = PayRec(1).LICPAID3
  CMTrRec(1).TransRevAmt(9) = PayRec(1).LICPAID4
  CMTrRec(1).TransRevAmt(10) = PayRec(1).LICPAID5
  'penalty
  CMTrRec(1).TransRevAmt(11) = PayRec(1).PENPAID
  'issuefee
  CMTrRec(1).TransRevAmt(12) = PayRec(1).ISSPAID
  CMTrRec(1).ChkByte = Chr$(1)
  CmNum = (LOF(CHandle) / CMTrRecLen) + 1
  Put CHandle, CmNum, CMTrRec(1)
  Close CHandle

  Close

  BLLog ("PayTrans posted. " + Str$(OperNum))
  CMLog ("Oper " + Str$(OperNum) + " BL Pay postedCM for trans# " + Str$(CmNum))
'  frmBLMessageBoxJr.Label1.Caption = "Posting completed successfully!"
'  frmBLMessageBoxJr.Label1.FontSize = 16
'  frmBLMessageBoxJr.Label1.Top = 900
'  frmBLMessageBoxJr.Show vbModal
  
'  frmBLEnterPayments.Show
'  DoEvents
'  Unload frmBLPostTrans
  
  Exit Sub

ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLSaveTrans", "cmdSave_Click", Erl)
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
 '   CitiTerminate
  

End Sub

Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get RP1, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    ElseIf RcptPrnFile.PrnDefYN = 1 Then
      On Local Error GoTo nofound
      RP = FreeFile
      Open RecpPort For Output As RP
      Close RP
      RecpDef = 1
    End If
    If RcptPrnFile.CtlDefYN = 0 Then
      CntrlDef = 0
    Else
      CntrlDef = 1
    End If
    If RcptPrnFile.RValidate = 1 Then
      ValiDef = 1
      RctValidate = True
      GetUBBankINfo
    Else
      ValiDef = 0
      RctValidate = False
    End If
  Close RP1
  Else
    RecpDef = 99
    ValiDef = 0
  End If
Exit Sub
nofound:
  RecpDef = 99
  ValiDef = 0
End Sub



Private Sub fpcmdDrawer_Click()
  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  Port$ = QPTrim$(RecpPort)
  CMLog "Oper: " + Oper$ + "BL Pay-Open Drawer"
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  Print #PortFile, Chr$(7)
  Close PortFile
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    If fptxtTDate.Enabled = True Then
      fptxtTDate.SetFocus
    End If
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  GCustNum = 0
  lblOperator = OperNum
  lblOperName.Caption = PWUser
 
  Oper$ = QPTrim(lblOperator.Caption)

  Call LoadMe
  GetRcpInfo
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  SkipLostFocus = False
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      Call cmdExit_Click
      SendKeys "%C"
    Case vbKeyF11:
      Call cmdClear_Click
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdSave_Click
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF7:
      Call cmdCustList_Click
      SendKeys "%L"
      KeyCode = 0
    Case vbKeyF5:
      Call cmdCash_Click
      SendKeys "%a"
      KeyCode = 0
    Case vbKeyF6:
      Call cmdCheck_Click
      SendKeys "%k"
      KeyCode = 0
    Case vbKeyF8:
      Call cmdCharge_Click
      SendKeys "%g"
      KeyCode = 0
    Case vbKeyF3:
      Call cmdReset_Click
      SendKeys "%R"
      KeyCode = 0
    Case vbKeyF4:
      Call cmdInfo_Click
      SendKeys "%u"
      KeyCode = 0
    Case vbKeyF2:
      KeyCode = 0
      DoEvents
      fpcmdDrawer_Click
'    Case vbKeyF1:
'      Call cmdHelp_Click
'      SendKeys "%T"
'      KeyCode = 0
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
      BLLog ("terminated via menu bar on frmBLPayEntry.")
      CMLog ("terminated via menu bar on frmBLPayEntry.")
      CitiTerminate
      End
    End If
  End If
End Sub

Public Sub LoadMe()
  
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim LicBal As Double
  Dim PenBal As Double
  Dim LicPenBal As Double
  Dim One As Integer
  Dim DHandle As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As AREditPaymentRecType
  Dim x As Integer
  Dim WarnLicFlag As Boolean
  Dim WarnPenFlag As Boolean
  Dim ThisCatDesc As String * 26
  
  On Error GoTo ERRORSTUFF
  
  NotFirstLoad = False
  ThisCustXNum = 0

'  lblBalloon.Visible = False
'  fptxtTDate.ToolTipText = "Today's date."
'  fptxtAccount.ToolTipText = "Either enter a valid customer number here or select a customer number from the customer list brought up by pressing F7. Then pess F4 to populate this screen."
'  cmdGetCust.ToolTipText = "Press this button to retrieve the data for the customer whose number is entered in the 'Account #' field."
'  cmdCustList.ToolTipText = "Press this button to bring up a list of all currently saved customers."
'  fptxtName.ToolTipText = "This is a read only field."
'  fptxtAddress.ToolTipText = "This is a read only field."
'  fptxtCity.ToolTipText = "This is a read only field."
'  fptxtState.ToolTipText = "This is a read only field."
'  fptxtZip.ToolTipText = "This is a read only field."
'  fpcurrAmtDue1.ToolTipText = "This is a read only field. It indicates the total outstanding balance for this customer."
'  fpcmbType.ToolTipText = "Select cash, check or charge depending on the payment form."
'  fpcurrCashPaid.ToolTipText = "Enter the amount of cash the customer is remitting for this transaction."
'  fpcurrChkPaid.ToolTipText = "Enter the check/charge amount this customer is remitting for this transaction."
'  fpcurrTotRecd.ToolTipText = "This field tallies up the cash amount and the check amount tendered by this customer for this transaction."
'  fpcurrChange.ToolTipText = "This field automatically calculates the difference between the amount tendered by the customer and the amount owed for this transaction."
'  fptxtDesc.ToolTipText = "This field allows an optional comment regarding this transaction."
'  cmdClear.ToolTipText = "Press to reset all the amounts entered to zero."
'  cmdReset.ToolTipText = "Use this button if you have overridden automatic distribution and wish to have the program to redistribute the amounts entered."
'  fpcurrRevGTDue.ToolTipText = "This field contains the total outstanding balance for this customer."
'  fpcurrRevGTPay.ToolTipText = "Keeps a running total of amounts paid."
'  fpcurrPenTotDue.ToolTipText = "Displays the outstanding penalty fee balance for this customer."
'  fpcurrPenAmtTot.ToolTipText = "Enter the amount of payment earmarked for penalty fees."
'  fpcurrIssDue.ToolTipText = "Displays the outstanding issuance fee balance for this customer."
'  fpcurrIssAmt.ToolTipText = "Enter the amount of payment earmarked for issuance fees."
'  fpcurrLicTotDue.ToolTipText = "This field contains the accumulated total of all outstanding license fees."
'  fpcurrLicTotPay.ToolTipText = "This field contains the accumulated amounts paid for license fees."
'  fpcurrLicBal(0).ToolTipText = "This field contains the total outstanding balance for license category #1."
'  fpcurrLicBal(1).ToolTipText = "This field contains the total outstanding balance for license category #2."
'  fpcurrLicBal(2).ToolTipText = "This field contains the total outstanding balance for license category #3."
'  fpcurrLicBal(3).ToolTipText = "This field contains the total outstanding balance for license category #4."
'  fpcurrLicBal(4).ToolTipText = "This field contains the total outstanding balance for license category #5."
'  fpcurrLicAmt(0).ToolTipText = "Enter the amount earmarked for payment for license category #1 here."
'  fpcurrLicAmt(1).ToolTipText = "Enter the amount earmarked for payment for license category #2 here."
'  fpcurrLicAmt(2).ToolTipText = "Enter the amount earmarked for payment for license category #3 here."
'  fpcurrLicAmt(3).ToolTipText = "Enter the amount earmarked for payment for license category #4 here."
'  fpcurrLicAmt(4).ToolTipText = "Enter the amount earmarked for payment for license category #5 here."
'  cmdExit.ToolTipText = "Press to exit this screen."
'  cmdSave.ToolTipText = "Press to commmit the data on this screen to a temporary file. Any transaction can be edited until it is posted."
'  cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
'  cmdDelete.ToolTipText = "Remove this transaction from the pending transaction list."
  
  
  
  If Exist("custinfotrans.dat") Then Exit Sub
  WarnLicFlag = False
  WarnPenFlag = False
  NegFlag = False
  'lblOpNum.Caption = "Operator Name: " + PWUser

  For x = 0 To 4
    If fpcurrLicBal(x).DoubleValue < 0 Then
      NegFlag = True
      Exit For
    End If
  Next x
  
  'Label18.Caption = "Operator Number: " + CStr(OperNum)
'  One = 1
'  DHandle = FreeFile
'  Open "transentry.dat" For Output As DHandle Len = 2
'  Print #DHandle, One
'  Close DHandle
  
  fptxtTDate = DefPayDate$
  ThisDate$ = DefPayDate$
  
  fpcmbType.Clear
  fpcmbType.Text = "Cash"
  fpcmbType.AddItem "Cash"
  fpcmbType.AddItem "Check"
  fpcmbType.AddItem "Cash & Check"
  fpcmbType.AddItem "Charge"
  fpcmbSetFlag.Clear
  fpcmbSetFlag.Text = "No"
  fpcmbSetFlag.AddItem "No"
  fpcmbSetFlag.AddItem "Yes"
  
  Call Clearscreen
  
  'cmdDelete.Enabled = False
  
'  If EditFlag = True Then
'    cmdDelete.Enabled = True
'    If GPayNum = 0 Then GoTo PayNumIsZero 'means that the
'    'user pulled up a customer that has an existing beginning
'    'balance done under another operator's number and opted to
'    'go ahead and add another transaction anyway
'    OpenPayFile PayHandle, OperNum
'    Get PayHandle, GPayNum, EditPayRec
'    GoSub GetNumOfCats
'    GoSub LoadEdit
'    Close
'    Exit Sub
'  End If
  
PayNumIsZero:
  If GCustNum = 0 Then
    fptxtAccount.TabIndex = 0
    Exit Sub
  End If
  
  OpenBLCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  
  'the following code checks to see if the customer
  'for whom this transaction entry is targeted is
  'included in a temporary post file...if so then
  'the user can either abort this transaction entry
  'attempt or continue and in so doing, delete the
  'unposted file
  If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
    If Exist("artmppen.dat") Then
      If EmpInPenProcess(CStr(GCustNum)) = True Then
        Get CHandle, GCustNum, CustRec
        frmBLMessageBoxJr.Label1.Caption = QPTrim$(CustRec.CustName) + " is currently involved in an unposted penalty fee file."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        BLLog ("User alerted that " + QPTrim$(CustRec.CustName) + " will is included in an unposted penalty file.")
      End If
    End If
    If Exist("artmppst.dat") Then
      Unload frmBLCustomerList
      If EmpInLicProcess(CStr(GCustNum)) = True Then
        WarnLicFlag = True
        Get CHandle, GCustNum, CustRec
        frmBLMessageBoxJrWOpts.Label1.Caption = QPTrim$(CustRec.CustName) + " is currently involved in an unposted business license fee file. These files would be rendered inaccurate if a transaction is entered here and posted. If you wish to continue then the unposted business license fee file WILL BE DELETED. To abort this transaction entry attempt press ESC. Otherwise, press F10 to continue and DELETE the unposted business license fee file."
        frmBLMessageBoxJrWOpts.Label1.Top = 350
        frmBLMessageBoxJrWOpts.Label1.Height = 1500
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          Close
          Exit Sub
        Else
          Unload frmBLMessageBoxJrWOpts
          KillFile "artmppst.dat"
          frmBLMessageBoxJr.Label1.Caption = "The unposted business license fee file has been deleted."
          frmBLMessageBoxJr.Label1.Top = 800
          frmBLMessageBoxJr.Show vbModal
          BLLog ("User warned that continuing with entering a transaction for " + QPTrim$(CustRec.CustName) + " will delete the unposted business license fee file because that customer is included in that business license fee file. The user elected to continue and the file was deleted.")
        End If
      End If
    End If
  End If
  
  If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
    Get CHandle, GCustNum, CustRec
    If CustRec.Deleted = "Y" Then
      frmBLMessageBoxJr.Label1.Caption = "This customer has been deleted"
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      If fptxtAccount.Enabled = True Then
        fptxtAccount.SetFocus
      End If
      Close
      Exit Sub
    End If
    If WarnLicFlag = True Then
      BLLog ("User warned that the customer, " + QPTrim$(CustRec.CustName) + ", is involved in an unposted license fee calculations file and the balances displayed do not include the new fees. Screen not loaded for this customer.")
    End If
    If WarnPenFlag = True Then
      BLLog ("User warned that the customer, " + QPTrim$(CustRec.CustName) + ", is involved in an unposted penalty calculations file and the balances displayed do not include the new fees. Screen not loaded for this customer.")
    End If
  End If
  
  Close CHandle
  
  GoSub GetNumOfCats
  
  fptxtAccount.Text = QPTrim$(CustRec.CUSTNUMB)
  ThisCustXNum = CInt(CustRec.CUSTNUMB)
  TempAcctNum$ = QPTrim$(CustRec.CUSTNUMB)
  fptxtName.Text = QPTrim$(CustRec.CustName)
  fptxtAddress.Text = QPTrim$(CustRec.ADDRESS1)
  fptxtCity.Text = QPTrim$(CustRec.City)
  fptxtState.Text = QPTrim$(CustRec.State)
  fptxtZip.Text = QPTrim$(CustRec.ZIPCODE)
  
  If CustRec.AcctBal < 1000000000 And CustRec.AcctBal > -1000000000 Then
    fpcurrAmtDue1 = OldRound(CustRec.AcctBal)
  Else
    fpcurrAmtDue1 = 0
  End If
  TempAcctBal = fpcurrAmtDue1.DoubleValue
  
  ReDim TempLicBal(0 To 4)
  fpcurrLicBal(0) = CustRec.FeeLicBal1
  TempLicBal(0) = CustRec.FeeLicBal1
  
  fpcurrLicBal(1) = CustRec.FeeLicBal2
  TempLicBal(1) = CustRec.FeeLicBal2
  
  fpcurrLicBal(2) = CustRec.FeeLicBal3
  
  TempLicBal(2) = CustRec.FeeLicBal3
  
  fpcurrLicBal(3) = CustRec.FeeLicBal4
  TempLicBal(3) = CustRec.FeeLicBal4
  
  fpcurrLicBal(4) = CustRec.FeeLicBal5
  
  TempLicBal(4) = CustRec.FeeLicBal5
  
  fpcurrLicTotDue = fpcurrLicBal(0).DoubleValue + fpcurrLicBal(1).DoubleValue + fpcurrLicBal(2).DoubleValue + fpcurrLicBal(3).DoubleValue + fpcurrLicBal(4).DoubleValue
  
  fpcurrLicAmt(0) = 0
  fpcurrLicAmt(1) = 0
  fpcurrLicAmt(2) = 0
  fpcurrLicAmt(3) = 0
  fpcurrLicAmt(4) = 0
  
  fpcurrPenTotDue = CustRec.PenBal
  TempPenBal = CustRec.PenBal
  TempPenPaid = 0
  fpcurrIssDue = CustRec.IssuanceBal
  TempIssBal = CustRec.IssuanceBal
  TempIssPaid = 0
  fpcurrRevGTDue = CustRec.PenBal + fpcurrLicTotDue.DoubleValue + CustRec.IssuanceBal
  fpcurrRevGTPay = 0
  ReDim TempLicPaid(0 To 4)
  
  TempLicPaid(0) = 0
  TempLicPaid(1) = 0
  TempLicPaid(2) = 0
  TempLicPaid(3) = 0
  TempLicPaid(4) = 0
  
  CatDesc(0) = GetCatDesc(QPTrim$(CustRec.BILLCAT1))
  CatDesc(1) = GetCatDesc(QPTrim$(CustRec.BILLCAT2))
  CatDesc(2) = GetCatDesc(QPTrim$(CustRec.BILLCAT3))
  CatDesc(3) = GetCatDesc(QPTrim$(CustRec.BILLCAT4))
  CatDesc(4) = GetCatDesc(QPTrim$(CustRec.BILLCAT5))
  
  CatDesc0 = GetCatDesc(QPTrim$(CustRec.BILLCAT1))
  CatDesc1 = GetCatDesc(QPTrim$(CustRec.BILLCAT2))
  CatDesc2 = GetCatDesc(QPTrim$(CustRec.BILLCAT3))
  CatDesc3 = GetCatDesc(QPTrim$(CustRec.BILLCAT4))
  CatDesc4 = GetCatDesc(QPTrim$(CustRec.BILLCAT5))
  
  TempTotPaid = 0
  TempTotDue = fpcurrRevGTDue.DoubleValue
  TempTotBal = fpcurrRevGTDue.DoubleValue
  TempChkAmt = 0
  TempCashAmt = 0
  TempTotRecd = 0
  TempChange = 0
  If CustRec.IssueLicense = "Y" Then
    TempPrintFlag = "Y"
    fpcmbSetFlag.Text = "Yes"
  ElseIf CustRec.IssueLicense = "N" Then
    TempPrintFlag = "N"
    fpcmbSetFlag.Text = "No"
  End If
  
  NotFirstLoad = True
  
  fpcmbType.TabIndex = 0
  
  BLLog ("PayBLTrans entry screen opened.")
  Close
  
  DoEvents
  NotFirstLoad = True
  Exit Sub
    
LoadEdit:
  Close CHandle
  fptxtAccount.Text = QPTrim$(EditPayRec.CustNumber)
  TempAcctNum$ = QPTrim$(EditPayRec.CustNumber)
  fptxtName.Text = QPTrim$(EditPayRec.CustName)
  fptxtAddress.Text = QPTrim$(EditPayRec.Add1)
  fptxtCity.Text = QPTrim$(EditPayRec.City)
  fptxtState.Text = QPTrim$(EditPayRec.State)
  fptxtZip.Text = QPTrim$(EditPayRec.ZIPCODE)
  
  If EditPayRec.TotDue < 1000000000 And EditPayRec.TotDue > -1000000000 Then
    fpcurrAmtDue1 = OldRound(EditPayRec.TotDue)
  Else
    fpcurrAmtDue1 = 0
  End If
  TempAcctBal = fpcurrAmtDue1.DoubleValue
  
  If QPTrim$(EditPayRec.CASHCHK) = "Both" Then
    fpcmbType.Text = "Cash & Check"
  Else
    fpcmbType.Text = QPTrim$(EditPayRec.CASHCHK)
  End If
  
  fpcurrCashPaid = EditPayRec.CashAmt
  TempCashAmt = EditPayRec.CashAmt
  
  If EditPayRec.ChkAmt > 0 Then
    fpcurrChkPaid = EditPayRec.ChkAmt
  ElseIf EditPayRec.CREDITAM > 0 Then
    fpcurrChkPaid = EditPayRec.CREDITAM
  End If
  
  TempChkAmt = EditPayRec.ChkAmt
  TempCreditAmt = EditPayRec.CREDITAM
  fpCurrTotRecd = EditPayRec.AmtPaid
  TempTotRecd = EditPayRec.AmtPaid
  fpcurrChange = EditPayRec.Change
  TempChange = EditPayRec.Change
  If EditPayRec.ISSUELIC = "N" Then
    fpcmbSetFlag.Text = "No"
  Else
    fpcmbSetFlag.Text = "Yes"
  End If
  TempPrintFlag = Mid(fpcmbSetFlag.Text, 1, 1)
  fptxtDesc.Text = QPTrim$(EditPayRec.Desc)
  
  CatDesc(0) = QPTrim$(EditPayRec.CatDesc1)
  CatDesc(1) = QPTrim$(EditPayRec.CatDesc2)
  CatDesc(2) = QPTrim$(EditPayRec.CatDesc3)
  CatDesc(3) = QPTrim$(EditPayRec.CatDesc4)
  CatDesc(4) = QPTrim$(EditPayRec.CatDesc5)
  CatDesc0 = QPTrim$(EditPayRec.CatDesc1)
  CatDesc1 = QPTrim$(EditPayRec.CatDesc2)
  CatDesc2 = QPTrim$(EditPayRec.CatDesc3)
  CatDesc3 = QPTrim$(EditPayRec.CatDesc4)
  CatDesc4 = QPTrim$(EditPayRec.CatDesc5)

  ReDim TempLicBal(0 To 4)
  fpcurrLicBal(0) = EditPayRec.LICDUE1
  TempLicBal(0) = EditPayRec.LICDUE1
  
  fpcurrLicBal(1) = EditPayRec.LICDUE2
  TempLicBal(1) = EditPayRec.LICDUE2
  
  fpcurrLicBal(2) = EditPayRec.LICDUE3
  TempLicBal(2) = EditPayRec.LICDUE3
  
  fpcurrLicBal(3) = EditPayRec.LICDUE4
  TempLicBal(3) = EditPayRec.LICDUE4
  
  fpcurrLicBal(4) = EditPayRec.LICDUE5
  TempLicBal(4) = EditPayRec.LICDUE5
  
  fpcurrLicTotDue = fpcurrLicBal(0).DoubleValue + fpcurrLicBal(1).DoubleValue + fpcurrLicBal(2).DoubleValue + fpcurrLicBal(3).DoubleValue + fpcurrLicBal(4).DoubleValue
  
  fpcurrLicAmt(0) = EditPayRec.LICPAID1
  fpcurrLicAmt(1) = EditPayRec.LICPAID2
  fpcurrLicAmt(2) = EditPayRec.LICPAID3
  fpcurrLicAmt(3) = EditPayRec.LICPAID4
  fpcurrLicAmt(4) = EditPayRec.LICPAID5
  fpcurrLicTotPay = EditPayRec.LICPAID1 + EditPayRec.LICPAID2 + EditPayRec.LICPAID3 + EditPayRec.LICPAID4 + EditPayRec.LICPAID5
  fpcurrPenTotDue = EditPayRec.PENDUE
  TempPenBal = EditPayRec.PENDUE
  TempPenPaid = EditPayRec.PENPAID
  fpcurrRevGTDue = EditPayRec.TotDue
  fpcurrPenAmtTot = EditPayRec.PENPAID
  fpcurrIssDue = EditPayRec.ISSDUE
  fpcurrIssAmt = EditPayRec.ISSPAID
  TempIssBal = EditPayRec.ISSDUE
  TempIssPaid = EditPayRec.ISSPAID
  fpcurrRevGTPay = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue
  
  ReDim TempLicPaid(0 To 4)
  
  TempLicPaid(0) = EditPayRec.LICPAID1
  TempLicPaid(1) = EditPayRec.LICPAID2
  TempLicPaid(2) = EditPayRec.LICPAID3
  TempLicPaid(3) = EditPayRec.LICPAID4
  TempLicPaid(4) = EditPayRec.LICPAID5
  TempTotPaid = EditPayRec.TotPaid
  TempTotDue = fpcurrRevGTDue.DoubleValue
  TempTotBal = fpcurrRevGTDue.DoubleValue
'  TempChkAmt = EditPayRec.CHKAMT
'  TempCashAmt = EditPayRec.CASHAMT
'  TempCreditAmt = EditPayRec.CREDITAM
'  TempTotRecd = EditPayRec.Amount
  NotFirstLoad = True
  BLLog ("PayBLTrans entry screen opened.")
  
  Close
  
  Return
GetNumOfCats:
  NumOfCodes = 0
'  If EditFlag = True Then
'    If Len(QPTrim(EditPayRec.CatDesc1)) > 0 Then
'      NumOfCodes = NumOfCodes + 1
'      fpcurrLicAmt(0).Enabled = True
'    Else
'      fpcurrLicAmt(0).Enabled = False
'    End If
'    If Len(QPTrim$(EditPayRec.CatDesc2)) > 0 Then
'      NumOfCodes = NumOfCodes + 1
'      fpcurrLicAmt(1).Enabled = True
'    Else
'      fpcurrLicAmt(1).Enabled = False
'    End If
'    If Len(QPTrim$(EditPayRec.CatDesc3)) > 0 Then
'      NumOfCodes = NumOfCodes + 1
'      fpcurrLicAmt(2).Enabled = True
'    Else
'      fpcurrLicAmt(2).Enabled = False
'    End If
'    If Len(QPTrim$(EditPayRec.CatDesc4)) > 0 Then
'      NumOfCodes = NumOfCodes + 1
'      fpcurrLicAmt(3).Enabled = True
'    Else
'      fpcurrLicAmt(3).Enabled = False
'    End If
'    If Len(QPTrim$(EditPayRec.CatDesc5)) > 0 Then
'      NumOfCodes = NumOfCodes + 1
'      fpcurrLicAmt(4).Enabled = True
'    Else
'      fpcurrLicAmt(4).Enabled = False
'    End If
'  Else
    If Len(QPTrim$(CustRec.DESC1)) > 0 Then
      NumOfCodes = NumOfCodes + 1
      fpcurrLicAmt(0).Enabled = True
    Else
      fpcurrLicAmt(0).Enabled = False
    End If
    If Len(QPTrim$(CustRec.DESC2)) > 0 Then
      NumOfCodes = NumOfCodes + 1
      fpcurrLicAmt(1).Enabled = True
    Else
      fpcurrLicAmt(1).Enabled = False
    End If
    If Len(QPTrim$(CustRec.DESC3)) > 0 Then
      NumOfCodes = NumOfCodes + 1
      fpcurrLicAmt(2).Enabled = True
    Else
      fpcurrLicAmt(2).Enabled = False
    End If
    If Len(QPTrim$(CustRec.Desc4)) > 0 Then
      NumOfCodes = NumOfCodes + 1
      fpcurrLicAmt(3).Enabled = True
    Else
      fpcurrLicAmt(3).Enabled = False
    End If
    If Len(QPTrim$(CustRec.Desc5)) > 0 Then
      NumOfCodes = NumOfCodes + 1
      fpcurrLicAmt(4).Enabled = True
    Else
      fpcurrLicAmt(4).Enabled = False
    End If
'  End If
  
  Return
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "LoadMe", Erl)
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
 '   CitiTerminate

End Sub

Private Sub fpcmbSetFlag_Change()
  If QPTrim$(fpcmbSetFlag.Text) = "" Then
    If TempPrintFlag = "N" Then
      fpcmbSetFlag.Text = "No"
      Exit Sub
    ElseIf TempPrintFlag = "Y" Then
      fpcmbSetFlag.Text = "Yes"
      Exit Sub
    End If
  End If
End Sub

Private Sub fpcmbSetFlag_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbSetFlag.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbSetFlag.ListIndex = -1
  End If
  If fpcmbSetFlag.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtDesc.Enabled = True Then
        fptxtDesc.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbType_Change()
  If QPTrim$(fpcmbType.Text) = "" Then
    fpcmbType.Text = "Cash"
  End If
  If fpcmbType.Text = "Cash" Then
    fpcurrChkPaid = 0
    fpcurrCashPaid.Enabled = True
    Label20.Caption = "Check Amount Tendered:"
    fpcurrChkPaid.Enabled = False
'    If NotFirstLoad = True Then 'this sub is activated before
'    'the screen is loaded so the setfocus function is not
'    'allowed
'      If fpcurrCashPaid.Enabled = True Then
'        fpcurrCashPaid.SetFocus
'      End If
'    End If
  ElseIf fpcmbType.Text = "Check" Then
    fpcurrCashPaid = 0
    fpcurrCashPaid.Enabled = False
    Label20.Caption = "Check Amount Tendered:"
    fpcurrChkPaid.Enabled = True
'    If NotFirstLoad = True Then
'      If fpcurrChkPaid.Enabled = True Then
'        fpcurrChkPaid.SetFocus
'      End If
'    End If
  ElseIf fpcmbType.Text = "Cash & Check" Then
    fpcurrCashPaid.Enabled = True
    fpcurrChkPaid.Enabled = True
    Label20.Caption = "Check Amount Tendered:"
'    If NotFirstLoad = True Then
'      If fpcurrCashPaid.Enabled = True Then
'        fpcurrCashPaid.SetFocus
'      End If
'    End If
  ElseIf fpcmbType.Text = "Charge" Then
    fpcurrCashPaid = 0
    fpcurrCashPaid.Enabled = False
    Label20.Caption = "Charge Amount Tendered:"
    fpcurrChkPaid.Enabled = True
'    If NotFirstLoad = True Then
'      If fpcurrChkPaid.Enabled = True Then
'        fpcurrChkPaid.SetFocus
'      End If
'    End If
  End If
  
End Sub
'Private Sub fpcmbType_SelChange(ItemIndex As Long)
'  'If BeenDone Then
'    fixamts
' 'End If
'End Sub
'Private Sub fixamts()
'
'  fpcmbType.Action = ActionClearSearchBuffer
'  If noreset = False Then
'
'    If fpcmbType.ListIndex = 0 Then
'      fpcurrCashPaid.Enabled = True
'      fpcurrChkPaid.Enabled = False
'    ElseIf fpcmbType.ListIndex = 1 Then
'      fpcurrCashPaid.Enabled = False
'      fpcurrChkPaid.Enabled = True
'    ElseIf fpcmbType.ListIndex = 2 Then
'      fpcurrCashPaid.Enabled = True
'      fpcurrChkPaid.Enabled = True
'    ElseIf fpcmbType.ListIndex = 3 Then
'      fpcurrCashPaid.Enabled = False
'      fpcurrChkPaid.Enabled = True
'      fpcurrChange.Enabled = False
'    End If
'  End If
'  DoEvents
'  noreset = False
'End Sub

Private Sub fpcmbType_GotFocus()
  If QPTrim$(fptxtAccount.Text) <> "" Then
    fptxtAccount.TabStop = False
  Else
    If fptxtAccount.Enabled = True Then
      fptxtAccount.SetFocus
      Exit Sub
    End If
  End If
  
'  SkipLostFocus = False
  fpcmbType.TabIndex = 1
  fptxtAccount.TabIndex = 0
End Sub

Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcurrCashPaid.Enabled = True Then
        fpcurrCashPaid.SetFocus
      ElseIf fpcurrChkPaid.Enabled = True Then
        fpcurrChkPaid.SetFocus
      Else
        If fpcmbSetFlag.Enabled = True Then
          fpcmbSetFlag.SetFocus
        End If
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If fpcurrIssAmt.Enabled = True Then
          fpcurrIssAmt.SetFocus
        End If
      End If
      KeyCode = 0
    End If
  End If

End Sub


Private Sub fpcurrCashPaid_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpcurrChkPaid.Enabled = True Then
      fpcurrChkPaid.SetFocus
    Else
      If fpcmbSetFlag.Enabled = True Then
        fpcmbSetFlag.SetFocus
      End If
    End If
  End If
End Sub

Private Sub fpcurrCashPaid_KeyPress(KeyAscii As Integer)
  If QPTrim$(fpcmbType.Text) = "" Then
    fpcmbType.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please enter a transaction type before entering a cash amount."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbType.BackColor = &HFFFFFF
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
  End If
End Sub

Private Sub fpcurrCashPaid_LostFocus()
  Dim CashPaid As Double
  Dim ChkPaid As Double
  
  CashPaid = fpcurrCashPaid.DoubleValue
  ChkPaid = fpcurrChkPaid.DoubleValue
  
  fpCurrTotRecd = CashPaid + ChkPaid
  
  Call MakeChange

End Sub

Private Sub fpcurrChange_Change()
  If fpcurrChange.DoubleValue < 0 Then
    fpcurrChange.BackColor = &H80FFFF
  Else
    fpcurrChange.BackColor = &HFFFFFF
  End If
End Sub

Private Sub fpcurrChkPaid_KeyPress(KeyAscii As Integer)
  If QPTrim$(fpcmbType.Text) = "" Then
    fpcmbType.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please enter a transaction type before entering a check amount."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbType.BackColor = &HFFFFFF
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
  End If

End Sub

Private Sub fpcurrChkPaid_LostFocus()
  Dim CashPaid As Double
  Dim ChkPaid As Double
  
  CashPaid = fpcurrCashPaid.DoubleValue
  ChkPaid = fpcurrChkPaid.DoubleValue
  fpCurrTotRecd = CashPaid + ChkPaid
  
  Call MakeChange
End Sub

Private Sub Clearscreen()
  Dim x As Integer
  
  fptxtAccount.Text = ""
  fptxtName.Text = ""
  fptxtAddress.Text = ""
  fptxtCity.Text = ""
  fptxtState.Text = ""
  fptxtZip.Text = ""
  fpcurrCashPaid = 0
  fpcurrChkPaid = 0
  fpcurrAmtDue1 = 0
  fpCurrTotRecd = 0
  fpcurrChange = 0
  For x = 0 To 4
    fpcurrLicBal(x) = 0
    fpcurrLicAmt(x) = 0
    fpcurrLicAmt(x).ControlType = ControlTypeNormal
    CatDesc(x) = ""
  Next x
  fptxtDesc.Text = ""
  fpcmbType.Text = "Cash"
  fpcurrRevGTPay = 0
  fpcurrPenAmtTot = 0
  fpcurrRevGTDue = 0
  fpcurrPenTotDue = 0
  fpcurrLicTotDue = 0
  fpcurrLicTotPay = 0
  fpcurrIssDue = 0
  fpcurrIssAmt = 0
End Sub

Private Sub fpcurrIssAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcurrPenAmtTot.Enabled = True Then
      fpcurrPenAmtTot.SetFocus
    End If
  End If
End Sub

Private Sub fpcurrIssAmt_LostFocus()
  Dim ThisTot As Double
  
  If fpcurrIssAmt.DoubleValue < 0 Then
    fpcurrIssAmt.BackColor = &H80FFFF
    frmBLMessageBox.Label1.Caption = "Only positive values are valid for amounts paid."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    fpcurrIssAmt.BackColor = &HFFFFFF
    fpcurrIssAmt = 0
    If fpcurrIssAmt.Enabled = True Then
      fpcurrIssAmt.SetFocus
    End If
    Exit Sub
  End If
  
  ThisTot = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue
  
  fpcurrRevGTPay = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue

End Sub

Private Sub fpcurrLicAmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If Index <> 4 Then
      If fpcurrLicAmt(Index + 1).Enabled = True Then
        fpcurrLicAmt(Index + 1).SetFocus
      ElseIf fpcurrPenAmtTot.Enabled = True Then
        fpcurrPenAmtTot.SetFocus
      End If
    ElseIf Index = 4 Then
      If fpcurrPenAmtTot.Enabled = True Then
        fpcurrPenAmtTot.SetFocus
      End If
    End If
  End If
  
  If KeyCode = vbKeyUp Then
    If Index <> 0 Then
      If fpcurrLicAmt(Index - 1).Enabled = True Then
        fpcurrLicAmt(Index - 1).SetFocus
      End If
    ElseIf Index = 0 Then
      If fptxtDesc.Enabled = True Then
        fptxtDesc.SetFocus
      End If
    End If
  End If
   
End Sub

Private Sub fpcurrLicAmt_LostFocus(Index As Integer)
  Dim x As Integer
  Dim ThisLPay As Double
  Dim ThisTPay As Double
  Dim ThisNegAmt As Double
  Dim ThisNegBal As Double
  Dim ThisDif As Double
  
  If fpcurrLicAmt(Index).DoubleValue < 0 Then 'user entered a negative
    If fpcurrLicBal(Index).DoubleValue >= 0 Then 'amount due is positive
      frmBLMessageBoxJr.Label1.Caption = "Only positive values are valid for amounts tendered."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      If fpcurrLicAmt(Index).Enabled = True Then
        fpcurrLicAmt(Index).SetFocus
      End If
      Exit Sub
    Else
      If Abs(fpcurrLicAmt(Index).DoubleValue) > Abs(fpcurrLicBal(Index).DoubleValue) Then
        fpcurrLicAmt(Index).BackColor = &H80FFFF
        ThisNegBal = fpcurrLicBal(Index).DoubleValue
        ThisNegAmt = fpcurrLicAmt(Index).DoubleValue
        ThisDif = OldRound(ThisNegBal - ThisNegAmt)
        frmBLMessageBox.Label1.Top = 600
        frmBLMessageBox.Label1.Caption = "The maximum negative amount allowed for payment when a category already has a negative balance is the amount of the negative balance. Entering a negative amount less than the negative balance is, in effect, charging the customer the positive difference between the credit balance and the amount entered below the credit balance."
        frmBLMessageBox.Label1.Height = 1500
        frmBLMessageBox.Label2.Top = 2484
        frmBLMessageBox.Label2.Caption = "In this case: " + QPTrim$(Using("$##,##0.00", CStr(ThisNegBal))) + " - " + QPTrim$(Using("$##,##0.00", CStr(ThisNegAmt))) + " would make the outstanding balance equal " + QPTrim$(Using("$##,##0.00", CStr(ThisDif))) + ". This increases the outstanding balance such that it becomes positive or, in effect, assessing a fee of " + QPTrim$(Using("$##,##0.00", CStr(ThisDif))) + "."
        frmBLMessageBox.Label2.Height = 1500
        frmBLMessageBox.Show vbModal
        If fpcurrLicAmt(Index).Enabled = True Then
          fpcurrLicAmt(Index).SetFocus
          fpcurrLicAmt(Index).BackColor = &HFFFFFF
        End If
        Exit Sub
      End If
    End If
  End If
  
  fpcurrLicTotPay = 0
  
  For x = 0 To 4 'NumOfCodes - 1
    ThisLPay = ThisLPay + fpcurrLicAmt(x).DoubleValue
  Next x
  
  ThisTPay = ThisLPay + fpcurrPenAmtTot.DoubleValue + fpcurrIssAmt.DoubleValue
  
  fpcurrLicTotPay = ThisLPay
  fpcurrRevGTPay = ThisTPay
  
  Call MakeChange

End Sub

Private Sub fpcurrPenAmtTot_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpcurrIssAmt.Enabled = True Then
      fpcurrIssAmt.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcurrLicAmt(4).Enabled = True Then
      fpcurrLicAmt(4).SetFocus
    ElseIf fpcurrLicAmt(3).Enabled = True Then
      fpcurrLicAmt(3).SetFocus
    ElseIf fpcurrLicAmt(2).Enabled = True Then
      fpcurrLicAmt(2).SetFocus
    ElseIf fpcurrLicAmt(1).Enabled = True Then
      fpcurrLicAmt(1).SetFocus
    ElseIf fpcurrLicAmt(0).Enabled = True Then
      fpcurrLicAmt(0).SetFocus
    Else
      If fptxtDesc.Enabled = True Then
        fptxtDesc.SetFocus
      End If
    End If
  End If
End Sub

Private Sub fpcurrPenAmtTot_LostFocus()
  Dim ThisTot As Double
  
  If fpcurrPenAmtTot.DoubleValue < 0 Then
    fpcurrPenAmtTot.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Only positive values are valid for amounts tendered."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcurrPenAmtTot.BackColor = &HFFFFFF
    fpcurrPenAmtTot = 0
    If fpcurrPenAmtTot.Enabled = True Then
      fpcurrPenAmtTot.SetFocus
    End If
    Exit Sub
  End If
  
  ThisTot = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue
  
  fpcurrRevGTPay = fpcurrPenAmtTot.DoubleValue + fpcurrLicTotPay.DoubleValue + fpcurrIssAmt.DoubleValue
End Sub

Private Sub fpcurrRevGTPay_Change()
  Call MakeChange
End Sub

Private Sub fpcurrTotRecd_Change()
  Dim x As Integer, y As Integer
  Dim Collect As Double
  Dim Dif As Double
  
  On Error GoTo ERRORSTUFF
  'if the amount due is 0 then automatic amount distribution
  'doesn't work and the user will have to enter amounts manually
  
  'If EditFlag = False Then
    If fpcurrAmtDue1.DoubleValue = 0 Then
      If fpCurrTotRecd.DoubleValue > 0 Then
        fpcurrAmtDue1.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "Automatic amount distribution does not take place if the Amount Due equals $0.00. Please make sure the appropriate amounts are entered manually."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fpcurrAmtDue1.BackColor = &HFFFFFF
        If fpcmbType.Text = "Cash & Check" Then
          If fpcurrChkPaid.Enabled = True Then
            fpcurrChkPaid.SetFocus
          End If
        Else
          If fpcurrLicAmt(0).Enabled = True Then
            fpcurrLicAmt(0).SetFocus
          End If
        End If
        
        Exit Sub
        
        BLLog ("The amount due is zero for this customer but a value has been entered for amount received. The user was warned to make sure the appropriate amounts were manually entered.")
      End If
    End If
  'End If
  
  fpcurrPenAmtTot = 0
  fpcurrLicTotPay = 0
  fpcurrIssAmt = 0
  fpcurrRevGTPay = 0
  'Object here is to distribute the total amount received so
  'that penalty balance is paid first then each license balance
  'is paid in total from 1 to 5
  For x = 0 To 4 'NumOfCodes - 1
    fpcurrLicAmt(x) = 0
  Next x
  
  Collect = fpCurrTotRecd.DoubleValue
  
  If fpcurrPenTotDue.DoubleValue > 0 Then
    If Collect > fpcurrPenTotDue.DoubleValue Then
      fpcurrPenAmtTot = fpcurrPenTotDue.DoubleValue
      Collect = OldRound(Collect - fpcurrPenTotDue.DoubleValue)
    Else
      fpcurrPenAmtTot = Collect
      Collect = 0
    End If
  End If
      
  If fpcurrIssDue.DoubleValue > 0 Then
    If Collect > fpcurrIssDue.DoubleValue Then
      fpcurrIssAmt = fpcurrIssDue.DoubleValue
      Collect = OldRound(Collect - fpcurrIssDue.DoubleValue)
    Else
      fpcurrIssAmt = Collect
      Collect = 0
    End If
  End If
  
  'do not auto distribute if there are any
  'negative license balances
  
  For x = 0 To 4
    If fpcurrLicBal(x).DoubleValue < 0 Then
      Exit For
    End If
  Next x
  If x < 5 Then
    Exit Sub
  End If
  
  'auto distribute for license amounts
  
  For x = 0 To 4 ' NumOfCodes - 1
    If fpcurrLicBal(x).DoubleValue > Collect Then
      fpcurrLicAmt(x) = Collect
      Collect = 0
      Exit For
    Else
      If fpcurrLicBal(x).DoubleValue < 0 Then
        fpcurrLicAmt(x) = 0
      Else
        fpcurrLicAmt(x) = fpcurrLicBal(x).DoubleValue
      End If
      Collect = OldRound(Collect - fpcurrLicAmt(x).DoubleValue)
    End If
  Next x
  
  For x = 0 To 4 ' NumOfCodes - 1
    fpcurrLicTotPay = OldRound(fpcurrLicTotPay.DoubleValue + fpcurrLicAmt(x).DoubleValue)
  Next x
  
  fpcurrRevGTPay = OldRound(fpcurrLicTotPay.DoubleValue + fpcurrPenAmtTot.DoubleValue + fpcurrIssAmt.DoubleValue)
  
  'handling negative balance distribution...12/12/03
  
  For x = 0 To 4 'NumOfCodes - 1
    Collect = 0
    If fpcurrLicBal(x).DoubleValue < 0 Then
      Collect = Abs(fpcurrLicBal(x).DoubleValue)
      For y = 0 To 4 'NumOfCodes - 1
        If fpcurrLicAmt(y).DoubleValue < fpcurrLicBal(y).DoubleValue Then
          Dif = fpcurrLicBal(y).DoubleValue - fpcurrLicAmt(y).DoubleValue
          If Collect >= Dif Then
            fpcurrLicAmt(y) = fpcurrLicAmt(y).DoubleValue + Dif
            fpcurrLicAmt(x) = fpcurrLicAmt(x).DoubleValue - Dif
            Collect = Collect - Dif
            If Collect <= 0 Then Exit For
          Else
            fpcurrLicAmt(y) = Dif - Collect
            Collect = 0
            Exit For
          End If
        End If
      Next y
    End If
  Next x
      
  Exit Sub
    
NegLicBal:
  
  For x = 0 To 4 ' NumOfCodes - 1
    If fpcurrLicBal(x).DoubleValue > 0 Then 'the total license balance is negative
    'but if one of the license balances is still greater than zero then look for it
    'and apply any amount received to that balance...this should not happen with
    'non converted towns because this program forces all license overages to be
    'distributed to unpaid balances for other licenses (or penalty charges) before
    'allowing a debit balance
      If Collect <= fpcurrLicBal(x).DoubleValue Then
        fpcurrLicAmt(x) = Collect
        Collect = 0
        Exit For
      Else
        fpcurrLicAmt(x) = fpcurrLicBal(x).DoubleValue
        Collect = Collect - fpcurrLicBal(x).DoubleValue
      End If
    End If
  Next x
  

  'Collect has now been distributed so if there is any left over
  'then dump it into the first license balance
  
  If Collect > 0 Then
    fpcurrLicAmt(0) = fpcurrLicAmt(0).DoubleValue + Collect
  End If
  
  For x = 0 To 4 ' NumOfCodes - 1
    fpcurrLicTotPay = OldRound(fpcurrLicTotPay.DoubleValue + fpcurrLicAmt(x).DoubleValue)
  Next x
  
  fpcurrRevGTPay = OldRound(fpcurrLicTotPay.DoubleValue + fpcurrPenAmtTot.DoubleValue + fpcurrIssAmt.DoubleValue)
 
  Exit Sub
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "fpcurrTotRecd_Change", Erl)
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
 '   CitiTerminate
  
End Sub

Private Sub fptxtAccount_Change()
  If QPTrim$(fptxtAccount.Text) <> TempAcctNum$ Then
    NotFirstLoad = False
  End If
End Sub

Private Sub fptxtAccount_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcurrIssAmt.Enabled = True Then
      fpcurrIssAmt.SetFocus
    End If
  End If
End Sub

Private Sub GetCust()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  
  On Error Resume Next
  
  If QPTrim$(fptxtAccount.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a customer account number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    If fptxtAccount.Enabled = True Then
      fptxtAccount.SetFocus
    End If
    Exit Sub
  End If
  
  Number = QPTrim$(fptxtAccount.Text)
  
  OpenBLCustFile CHandle
  TotalAccts = LOF(CHandle) \ Len(CustRec)
  
  If TotalAccts = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To TotalAccts
    Get CHandle, x, CustRec
    If Number$ = QPTrim$(CustRec.CUSTNUMB) Then 'match the selected
    'row with the right code
      Found = True
      GCustNum = x 'now you can assign the correct global
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close CHandle
  
  If Found = True Then
    Call EnterEditChk
  End If
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
'Private Sub PrintReceipt(RctNum As Integer)
'  Dim TownName$
'  Dim PayHandle As Integer
'  Dim PayRec As AREditPaymentRecType
'  Dim NumOfPayRecs As Integer
'  Dim TownRec As TownSetUpType
'  Dim THandle As Integer
'  Dim PayRecpName$
'  Dim RctHandle As Integer
'  Dim RptHandle As Integer, LPTHandle As Integer
'  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
'  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
'
'  'On Error GoTo ERRORSTUFF
'
'  OpenTownFile THandle
'  Get THandle, 1, TownRec
'  Close THandle
'
'  TownName$ = QPTrim$(TownRec.TownName)
'  OpenPayFile PayHandle, OperNum
'  Get PayHandle, RctNum, PayRec
'  Close '1/29/04
'  PayRecpName$ = "RECPT.PRN"
'  RctHandle = FreeFile
'  Open PayRecpName$ For Output As #RctHandle
'  Print #RctHandle, ""
'  Print #RctHandle, QPTrim$(TownName$)
'  Print #RctHandle, "LICENSE PAYMENT"
'  Print #RctHandle, "Date: "; Num2Date(PayRec.TranDate)
'  Print #RctHandle, "Time: "; Time
'  Print #RctHandle,
'  Print #RctHandle, "Account #"; QPTrim$(PayRec.CustNumber)
'  Print #RctHandle, QPTrim$(PayRec.CustName)
'  Print #RctHandle, QPTrim$(PayRec.Add1)
'  Print #RctHandle, QPTrim$(PayRec.DESC)
'  Print #RctHandle,
'  Print #RctHandle, "Total Owed: "; Using("$#,###,0.00", PayRec.TOTDUE)
'  Print #RctHandle, ""
'  Print #RctHandle, "  Cash Amt: "; Using("$#,###,0.00", PayRec.CASHAMT)
'  If PayRec.CREDITAM > 0 Then
'    Print #RctHandle, "Charge Amt: "; Using("$#,###,0.00", PayRec.CREDITAM)
'  Else
'    Print #RctHandle, " Check Amt: "; Using("$#,###,0.00", PayRec.CHKAMT)
'  End If
'  Print #RctHandle, "             -----------"
'  Print #RctHandle, "Total Paid: "; Using("$#,###,0.00", PayRec.AMTPAID#)
'  Print #RctHandle, ""
'  Print #RctHandle, "    Change: "; Using("$#,###,0.00", PayRec.CHANGE)
'  Print #RctHandle,
'  Print #RctHandle,
'  Print #RctHandle, Tab(7); "T H A N K   Y O U !"
'  Print #RctHandle,
'  Print #RctHandle,
'  Print #RctHandle,
'  Print #RctHandle,
'  Print #RctHandle,
'  Print #RctHandle,
'  Print #RctHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
'  Close
'
'10:
'  DefPrinter = RecpPort
'20:
'
'  For CopyLoop = 1 To 1 'Copies
'    LPTHandle = FreeFile
'    Open DefPrinter For Output As LPTHandle
'    RptHandle = FreeFile
'30:
'    Open PayRecpName$ For Input As RptHandle
'40:
'    Do
'      If frmPrint.cmdCancel = False Then
'45:
'        Line Input #RptHandle, ToPrint$
'
'        ToPrint$ = RTrim$(ToPrint$)
'        Print #LPTHandle, ToPrint$
'      Else
'50:
'        Exit Do
'        'Printer.EndDoc
'      End If
'  Loop Until eof(RptHandle)
'60:
'  Close RptHandle
'62:
'  Close LPTHandle
'65:
'  Next CopyLoop
'68:
'  Printer.EndDoc
'
'  KillFile PayRecpName$
'
'  BLLog ("Payment receipt processed for " + QPTrim$(PayRec.CustName) + "... Total Owed: " + Using("$#,###,0.00", PayRec.TOTDUE) + "  Cash Amt: " + Using("$#,###,0.00", PayRec.CASHAMT))
'
'  If PayRec.CREDITAM > 0 Then
'    BLLog (" (Payment receipt cont.) Charge Amt: " + Using("$#,###,0.00", PayRec.CREDITAM) + " Total Paid: " + Using("$#,###,0.00", PayRec.AMTPAID#) + "    Change: " + Using("$#,###,0.00", PayRec.CHANGE))
'  Else
'    BLLog (" (Payment receipt cont.) Check Amt: " + Using("$#,###,0.00", PayRec.CHKAMT) + " Total Paid: " + Using("$#,###,0.00", PayRec.AMTPAID#) + "    Change: " + Using("$#,###,0.00", PayRec.CHANGE))
'  End If
'
'  Exit Sub
'
'ERRORSTUFF:
'   Unload FrmShowPctComp
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "PrintReceipt", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'    ClearInUse PWcnt
'    'Terminate
'
'End Sub

Private Sub LogSaves(ThisRec As Integer)
  Dim EditPayRec As AREditPaymentRecType
  Dim PayHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenPayFile PayHandle, OperNum
  Get PayHandle, ThisRec, EditPayRec
  Close PayHandle
  
  If OldRound(TempTotRecd) <> OldRound(EditPayRec.AmtPaid) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + "Total Received changed from " + QPTrim$(Using("$##,###,##0.00", TempTotRecd)) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.AmtPaid)) + ".")
  End If
  
  If OldRound(TempLicPaid(0)) <> OldRound(EditPayRec.LICPAID1) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + CatDesc0 + " License Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempLicPaid(0))) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICPAID1)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICDUE1)) + ".")
  End If

  If OldRound(TempLicPaid(1)) <> OldRound(EditPayRec.LICPAID2) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + CatDesc1 + " License Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempLicPaid(1))) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICPAID2)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICDUE2)) + ".")
  End If

  If OldRound(TempLicPaid(2)) <> OldRound(EditPayRec.LICPAID3) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + CatDesc2 + " License Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempLicPaid(2))) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICPAID3)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICDUE3)) + ".")
  End If

  If OldRound(TempLicPaid(3)) <> OldRound(EditPayRec.LICPAID4) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + CatDesc3 + " License Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempLicPaid(3))) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICPAID4)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICDUE4)) + ".")
  End If

  If OldRound(TempLicPaid(4)) <> OldRound(EditPayRec.LICPAID5) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + CatDesc4 + " License Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempLicPaid(4))) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICPAID5)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.LICDUE5)) + ".")
  End If

  If OldRound(TempIssPaid) <> OldRound(EditPayRec.ISSPAID) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Issuance Fee Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempIssPaid)) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.ISSPAID)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.ISSDUE)) + ".")
  End If
  
  If OldRound(TempPenPaid) <> OldRound(EditPayRec.PENPAID) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Penalty Paid changed from " + QPTrim$(Using("$##,###,##0.00", TempPenPaid)) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.PENPAID)) + ". Amount owed before posting is " + QPTrim$(Using("$##,###,##0.00", EditPayRec.PENDUE)) + ".")
  End If
  
  If OldRound(TempChkAmt) <> OldRound(EditPayRec.ChkAmt) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Check amount changed from " + QPTrim$(Using("$##,###,##0.00", TempChkAmt)) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.ChkAmt)) + ".")
  End If
  
  If OldRound(TempCashAmt) <> OldRound(EditPayRec.CashAmt) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Cash amount changed from " + QPTrim$(Using("$##,###,##0.00", TempCashAmt)) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.CashAmt)) + ".")
  End If
  
  If OldRound(TempCreditAmt) <> OldRound(EditPayRec.CREDITAM) Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Credit amount changed from " + QPTrim$(Using("$##,###,##0.00", TempCreditAmt)) + " to " + QPTrim$(Using("$##,###,##0.00", EditPayRec.CREDITAM)) + ".")
  End If
  
  If TempPrintFlag <> EditPayRec.ISSUELIC Then
    BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Print flag changed from " + TempPrintFlag + " to " + EditPayRec.ISSUELIC + ".")
  End If
  
  BLLog ("For " + QPTrim$(EditPayRec.CustName) + ": " + " Change issued was " + QPTrim$(Using("$##,###,##0.00", EditPayRec.Change)) + ".")
  Exit Sub
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "LogSaves", Erl)
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
  '  CitiTerminate
  
End Sub

Private Function TotalsOK() As Boolean
  Dim x As Integer
  Dim CodeCnt As Integer
  Dim TotDif As Double
  Dim ThisLPay As Double
  Dim ThisTPay As Double
  Dim RedCnt As Integer
  Dim NegBal As Double
  
  'in this function we are taking a look at the amounts entered
  'for payment distribution in the penalty, issuance fee and each of the
  'license categories...the program will not allow:
  '   1. penalty or issuance fees to be underpaid while a license fee is overpaid
  '   2. one license fee to be overpaid while another is underpaid
  '
  On Error GoTo ERRORSTUFF
  
  TotalsOK = True
  'category index starts at 0
  CodeCnt = 4 'NumOfCodes - 1
  
  For x = 0 To CodeCnt 'Total up amounts entered for payment
  'in the license amount fields
    ThisLPay = OldRound(ThisLPay + fpcurrLicAmt(x).DoubleValue)
    If fpcurrLicBal(x).DoubleValue < 0 Then
      NegBal = NegBal + fpcurrLicBal(x).DoubleValue
    End If
  Next x

  ThisLPay = OldRound(ThisLPay + NegBal)
  
  ThisTPay = OldRound(ThisLPay + fpcurrPenAmtTot.DoubleValue + fpcurrIssAmt.DoubleValue)
  
  If OldRound(ThisTPay) > OldRound(fpCurrTotRecd.DoubleValue) Then
    fpcurrLicTotPay.BackColor = &H8080FF
    fpcurrPenAmtTot.BackColor = &H8080FF
    fpcurrIssAmt.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Error: Amount distributed (red) is greater than the amount received (yellow)."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    TotalsOK = False
    fpcurrLicTotPay.BackColor = &HFFFFFF
    fpcurrPenAmtTot.BackColor = &HFFFFFF
    fpcurrIssAmt.BackColor = &HFFFFFF
    fpCurrTotRecd.BackColor = &HFFFFFF
    If fpCurrTotRecd.Enabled = True Then
      fpCurrTotRecd.SetFocus
    ElseIf fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
    Exit Function
  End If
  
  ReDim ThisDif(0 To CodeCnt) As Double
  TotDif = 0
  For x = 0 To CodeCnt
    ThisDif(x) = 0
  Next x
  
  If OldRound(fpcurrPenTotDue.DoubleValue - fpcurrPenAmtTot.DoubleValue) > 0 Then
    If OldRound(fpCurrTotRecd.DoubleValue > fpcurrPenTotDue.DoubleValue) Then
      frmBLMessageBoxJr.Label1.Caption = "Please be sure that any penalty amounts due are paid in full first."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      If fpcurrPenAmtTot.Enabled = True Then
        fpcurrPenAmtTot.SetFocus
      End If
      TotalsOK = False
      Exit Function
    End If
  End If
    
  For x = 0 To CodeCnt
    If OldRound(fpcurrLicAmt(x).DoubleValue - fpcurrLicBal(x).DoubleValue) > 0 Then 'overpaid category found
      ThisDif(x) = Abs(fpcurrLicAmt(x).DoubleValue - fpcurrLicBal(x).DoubleValue)
      TotDif = TotDif + ThisDif(x)
    End If
  Next x
    
  RedCnt = 0
  
  If TotDif > 0 Then 'at least one of the categories is overpaid
    If OldRound(fpcurrPenTotDue.DoubleValue - fpcurrPenAmtTot.DoubleValue) > 0 Then
      fpcurrPenAmtTot.BackColor = &H8080FF
      For x = 0 To CodeCnt
        If ThisDif(x) > 0 Then
          fpcurrLicAmt(x).BackColor = &H80FFFF
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "There are Business License categories with overpaid balances totaling " + QPTrim$(Using("$##,###0.00", TotDif)) + "(yellow) while at the same time the Penalty Due (red) amount is underpaid. Please make sure the Penalty is paid in full first."
      frmBLMessageBoxJr.Label1.Height = 1500
      frmBLMessageBoxJr.Label1.Top = 580
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False
      fpcurrPenAmtTot.BackColor = &H80000005
      For x = 0 To CodeCnt
        fpcurrLicAmt(x).BackColor = &H80000005
      Next x
      If fpcurrPenAmtTot.Enabled = True Then
        fpcurrPenAmtTot.SetFocus
      End If
    ElseIf OldRound(fpcurrIssDue.DoubleValue - fpcurrIssAmt.DoubleValue) > 0 Then
      fpcurrIssAmt.BackColor = &H8080FF
      For x = 0 To CodeCnt
        If ThisDif(x) > 0 Then
          fpcurrLicAmt(x).BackColor = &H80FFFF
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "There are Business License categories with overpaid balances totaling " + QPTrim$(Using("$##,###0.00", TotDif)) + " (yellow) while at the same time the Issuance Fee Due (red) amount is underpaid. Please make sure the Issuance Fee is paid in full first."
      frmBLMessageBoxJr.Label1.Height = 1500
      frmBLMessageBoxJr.Label1.Top = 580
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False
      fpcurrIssAmt.BackColor = &H80000005
      For x = 0 To CodeCnt
        fpcurrLicAmt(x).BackColor = &H80000005
      Next x
      If fpcurrIssAmt.Enabled = True Then
        fpcurrIssAmt.SetFocus
      End If
    Else
      If CodeCnt = 0 Then GoTo CodeCntIsZero 'only one category so
      'no need to check any further...looking to see if a license category
      'is overpaid while a different license category is underpaid
      For x = 0 To CodeCnt
        If OldRound(fpcurrLicBal(x).DoubleValue - fpcurrLicAmt(x).DoubleValue) > 0 Then 'found one with a positive balance
          fpcurrLicAmt(x).BackColor = &H8080FF
          RedCnt = RedCnt + 1
        ElseIf ThisDif(x) > 0 Then
          fpcurrLicAmt(x).BackColor = &H80FFFF
        End If
      Next x
      If RedCnt > 0 Then 'as long as all balances are paid then
      'its OK to allow overpayments
        frmBLMessageBoxJr.Label1.Caption = "There are still outstanding Business License category balances (red) while other balances are overpaid (yellow). Category overpayments are allowed only if all categories have been paid in full first."
        frmBLMessageBoxJr.Label1.Height = 1500
        frmBLMessageBoxJr.Label1.Top = 680
        frmBLMessageBoxJr.Show vbModal
        TotalsOK = False
      End If
      For x = 0 To CodeCnt
        fpcurrLicAmt(x).BackColor = &H80000005
      Next x
      If fpcmbType.Enabled = True Then
        fpcmbType.SetFocus
      End If
    End If
  End If
CodeCntIsZero:
  If TotalsOK = False Then Exit Function
  TotDif = 0
  
  For x = 0 To CodeCnt
    ThisDif(x) = 0
  Next x
  
  If OldRound(fpcurrPenTotDue.DoubleValue - fpcurrPenAmtTot.DoubleValue) < 0 Then 'overpaid penalty
    For x = 0 To CodeCnt
      If OldRound(fpcurrLicBal(x).DoubleValue - fpcurrLicAmt(x).DoubleValue) > 0 Then
         ThisDif(x) = OldRound(fpcurrLicBal(x).DoubleValue - fpcurrLicAmt(x).DoubleValue)
         TotDif = TotDif + ThisDif(x)
      End If
    Next x
    If TotDif > 0 Then
      For x = 0 To CodeCnt
        If ThisDif(x) > 0 Then
          fpcurrLicAmt(x).BackColor = &H8080FF
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The penalty amount due has been overpaid while there are license amounts due that are underpaid (red). After the penalty balance is paid in full please make sure any amounts left over get distributed to underpaid license balances."""
      frmBLMessageBoxJr.Label1.Height = 1500
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False
      If fpcurrPenAmtTot.Enabled = True Then
        fpcurrPenAmtTot.SetFocus
      End If
      For x = 0 To CodeCnt
        fpcurrLicAmt(x).BackColor = &H80000005
      Next x
    End If
  End If
  
  If OldRound(fpcurrIssDue.DoubleValue - fpcurrIssAmt.DoubleValue) < 0 Then 'overpaid penalty
    For x = 0 To CodeCnt
      If OldRound(fpcurrLicBal(x).DoubleValue - fpcurrLicAmt(x).DoubleValue) > 0 Then
         ThisDif(x) = OldRound(fpcurrLicBal(x).DoubleValue - fpcurrLicAmt(x).DoubleValue)
         TotDif = TotDif + ThisDif(x)
      End If
    Next x
    If TotDif > 0 Then
      For x = 0 To CodeCnt
        If ThisDif(x) > 0 Then
          fpcurrLicAmt(x).BackColor = &H8080FF
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The issue fee amount due has been overpaid while there are license amounts due that are underpaid (red). After the penalty balance is paid in full please make sure any amounts left over get distributed to underpaid license balances."
      frmBLMessageBoxJr.Label1.Height = 1500
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False
      If fpcurrIssAmt.Enabled = True Then
        fpcurrIssAmt.SetFocus
      End If
      For x = 0 To CodeCnt
        fpcurrLicAmt(x).BackColor = &H80000005
      Next x
    End If
  End If
  
  Exit Function
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "TotalsOK", Erl)
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
 '   CitiTerminate

End Function

Private Sub MakeChange()
  Dim ThisNeg As Double
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbType.Text = "Charge" Then
    fpcurrChange = 0
    Exit Sub
  End If
  fpcurrChange.BackColor = &H80000005
  If OldRound(fpcurrRevGTPay.DoubleValue) = 0 Then
'    fpcurrChange = OldRound(fpcurrTotRecd.DoubleValue - fpcurrAmtDue1.DoubleValue)
'    If fpcurrChange.DoubleValue < 0 Then fpcurrChange = 0
    fpcurrChange = fpCurrTotRecd.DoubleValue
  Else
    If OldRound(fpcurrRevGTPay.DoubleValue) <= OldRound(fpcurrAmtDue1.DoubleValue) Then
      fpcurrChange = OldRound(fpCurrTotRecd.DoubleValue - fpcurrRevGTPay.DoubleValue)
      If fpcurrChange.DoubleValue < 0 Then fpcurrChange = 0
    Else
      If OldRound(fpcurrRevGTPay.DoubleValue) > OldRound(fpcurrAmtDue1.DoubleValue) Then
        If OldRound(fpCurrTotRecd.DoubleValue) >= OldRound(fpcurrRevGTPay.DoubleValue) Then
          fpcurrChange = OldRound(fpCurrTotRecd.DoubleValue - fpcurrRevGTPay.DoubleValue)
        Else
          ThisNeg = OldRound(fpCurrTotRecd.DoubleValue - fpcurrRevGTPay.DoubleValue)
          fpcurrChange.BackColor = &H80FFFF
          fpcurrChange = ThisNeg
        End If
      End If
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransEntry", "MakeChange", Erl)
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
 '   CitiTerminate
  
End Sub

Private Function CompareAcctNumWData() As Boolean
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  
  'A user can bring up a business's data and then change the account number
  'without bringing up the new account number's data...this function is
  'designed to trap for this situation...the data would still be saved for the
  'business whose address appears and not for the business associated with
  'the new account number
  On Error Resume Next
  CompareAcctNumWData = True
  OpenBLCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  For x = 1 To NumOfCustRecs
    Get CHandle, x, CustRec
    'in testing it was discovered that sometimes a deleted customer
    'had the same customer name as a customer that wasn't deleted and
    'it would kick out the customer we wanted because the customer
    'numbers didn't match
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then GoTo DeletedAcct
    If QPTrim$(CustRec.CustName) = QPTrim$(fptxtName.Text) Then
      If QPTrim$(CustRec.CUSTNUMB) = QPTrim$(fptxtAccount.Text) Then
        Exit For
      Else
        CompareAcctNumWData = False
        frmBLMessageBoxJr.Label1.Caption = "The account number entered does not match the other data shown for this business. Please check the customer list for the correct data."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        Exit For
      End If
    End If
DeletedAcct:
  Next x
  Close CHandle
End Function

Private Sub fptxtAccount_LostFocus()
  If NotFirstLoad = True Then
    Exit Sub
  Else
    Call LostFocusCheck
  End If
End Sub

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpcurrLicAmt(0).Enabled = True Then
      fpcurrLicAmt(0).SetFocus
    ElseIf fpcurrLicAmt(1).Enabled = True Then
      fpcurrLicAmt(1).SetFocus
    ElseIf fpcurrLicAmt(2).Enabled = True Then
      fpcurrLicAmt(2).SetFocus
    ElseIf fpcurrLicAmt(3).Enabled = True Then
      fpcurrLicAmt(3).SetFocus
    ElseIf fpcurrLicAmt(4).Enabled = True Then
      fpcurrLicAmt(4).SetFocus
    Else
      If fpcurrPenAmtTot.Enabled = True Then
        fpcurrPenAmtTot.SetFocus
      End If
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcmbSetFlag.Enabled = True Then
      fpcmbSetFlag.SetFocus
    End If
  End If
      
End Sub

Private Function Check4ValidCustNum(ThisCust As String) As Boolean
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean

  Check4ValidCustNum = True
  
  If QPTrim$(fptxtAccount.Text) = "" Then
    Check4ValidCustNum = False
    Exit Function
  End If
  
  OpenBLCustFile CHandle
  TotalAccts = LOF(CHandle) \ Len(CustRec)

  If TotalAccts = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There alre no business customers saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  
  For x = 1 To TotalAccts
    Get CHandle, x, CustRec
    If ThisCust$ = QPTrim$(CustRec.CUSTNUMB) Then 'match the selected
    'row with the right code
      If CustRec.Deleted = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then
        Check4ValidCustNum = False
      End If
      Exit For
    End If
  Next x

  Close CHandle

  If x > TotalAccts Then
    Call Clearscreen
    Check4ValidCustNum = False
  End If
  
End Function

Private Sub GetInfo()

 If Exist("custinfomodal.dat") Then
   Exit Sub
 End If
 
 If QPTrim$(fptxtAccount.Text) <> "" Then
   If Check4ValidCustNum(QPTrim$(fptxtAccount.Text)) = True Then
     ThisCustXNum = CInt(fptxtAccount.Text)
     Load frmBLCustInfoTrans
     DoEvents
     frmBLCustInfoTrans.Show vbModal
     DoEvents
     Me.Hide
     DoEvents
   Else
     frmBLMessageBoxJr.Label1.Caption = "The customer number entered is not valid. Please enter a valid customer number."
     frmBLMessageBoxJr.Label1.Top = 800
     frmBLMessageBoxJr.Show vbModal
     'Clearscreen
     If fpcmbType.Enabled = True Then
       fpcmbType.SetFocus
     End If
   End If
 Else
   frmBLMessageBoxJr.Label1.Caption = "Please enter a valid customer number in the 'Account #' field."
   frmBLMessageBoxJr.Label1.Top = 800
   frmBLMessageBoxJr.Show vbModal
  ' Clearscreen
   If fptxtAccount.Enabled = True Then
     fptxtAccount.SetFocus
   End If
 End If

End Sub

Private Sub LostFocusCheck()
  If QPTrim$(fptxtAccount.Text) = "" Then
    Call Clearscreen
    Exit Sub
  End If
  
  If frmBLMessageBoxJr.Visible = True Then Exit Sub
  If Check4ValidCustNum(QPTrim$(fptxtAccount.Text)) = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer number is not valid. Please enter a valid customer number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Call Clearscreen
    If fpcmbType.Enabled = True Then
      fpcmbType.SetFocus
    End If
    Exit Sub
  End If

  Call GetCust
  
  If fpcmbType.Enabled = True Then
    fpcmbType.SetFocus
    DoEvents
  End If
  

End Sub
Private Sub PrintReceipt()
  Dim ListFile As Integer, PayFileName As String, PayRecLen As Integer
  Dim RecptNum As Long, RHandle As Integer, PayRecpName As String
  Dim CutPaper As String, PostDate As String, RevCnt As Integer
  Dim NumofRevs As Integer, RecpRev As String, RptHandle2 As Integer
  Dim RHandle2 As Integer, PayRecpName2 As String
  Dim PayRec(1) As AREditPaymentRecType
  RecpRev$ = Space$(15)
  CutPaper$ = Chr$(29) + Chr$(86) + Chr$(66) + Chr$(64)
   If InStr(TownName$, "Dobson") > 0 Then CutPaper$ = Chr$(27) + Chr$(100)
  PayRecLen = Len(PayRec(1))
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  PayRecpName$ = "c:\CPWork\CMRCP" + Oper$ + ".RPT"
  PayRecpName2$ = "C:\CPWork\CMVLD" + Oper$ + ".Rpt"
  PostDate$ = fptxtTDate.Text
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = PayRecLen
  'RecptNum& = LOF(ListFile) / UBPayRecLen
  Get #ListFile, 1, PayRec(1)
  Close
  If PrnRecp = False And PrnVali = True Then GoTo Validationthing
  NumofRevs = MaxRevsCnt
  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  If CntrlDef = 1 Then
    Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    Print #RHandle, Chr$(7)
  End If
  Print #RHandle, TownName$
  Print #RHandle, "CM BUSINESS LICENSE PAYMENT"
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle, "Time: "; Time
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, "--------------------------------"
  Print #RHandle, PayRec(1).CustName
  Print #RHandle, PayRec(1).Add1
  Print #RHandle, PayRec(1).Desc
  Print #RHandle, "Acct. No. "; PayRec(1).CustNumber
  Print #RHandle,
  Print #RHandle, QPTrim$(fpcmbType.Text)
  Print #RHandle,
  Print #RHandle, "       Cash: "; Using("$##,###,###.##", PayRec(1).CashAmt)
  If QPTrim$(PayRec(1).CASHCHK) <> "Charge" Then
    Print #RHandle, "      Check: "; Using("$##,###,###.##", PayRec(1).ChkAmt)
    Print #RHandle, "     Charge: "; Using("$##,###,###.##", 0)
  Else
    Print #RHandle, "      Check: "; Using("$##,###,###.##", 0)
    Print #RHandle, "     Charge: "; Using("$##,###,###.##", PayRec(1).ChkAmt)
  End If
  Print #RHandle, " Total Owed: "; Using("$##,###,###.##", PayRec(1).TotDue)
  Print #RHandle, " Total Paid: "; Using("$##,###,###.##", PayRec(1).AmtPaid)
  Print #RHandle, " Change Due: "; Using("$##,###,###.##", PayRec(1).Change)
  Print #RHandle, "Amt Applied: "; Using("$##,###,###.##", PayRec(1).TotPaid)
  Print #RHandle, "    Balance: "; Using("$##,###,###.##", PayRec(1).TotDue - PayRec(1).TotPaid)

  
'  Print #RHandle, "Total Paid: "; Using("$##,###,###.##", PayRec(1).AMTPAID)
'  Print #RHandle, "Change Due: "; Using("$##,###,###.##", PayRec(1).CHANGE)
'  Print #RHandle, "   Balance: "; Using("$##,###,###.##", PayRec(1).TOTDUE - PayRec(1).TotPaid)
  Print #RHandle,
    If PayRec(1).LICPAID1 <> 0 Then
      Print #RHandle, Mid$(PayRec(1).CatDesc1, 1, 10); " " + Using("$########.##", PayRec(1).LICPAID1)
      'Print #RHandle, Tab(25);
    End If
    If PayRec(1).LICPAID2 <> 0 Then
      Print #RHandle, Mid$(PayRec(1).CatDesc2, 1, 10); " " + Using("$########.##", PayRec(1).LICPAID2)
      'Print #RHandle, Tab(25);
    End If
    If PayRec(1).LICPAID3 <> 0 Then
      Print #RHandle, Mid$(PayRec(1).CatDesc3, 1, 10); " " + Using("$########.##", PayRec(1).LICPAID3)
      'Print #RHandle, Tab(25);
    End If
    If PayRec(1).LICPAID4 <> 0 Then
      Print #RHandle, Mid$(PayRec(1).CatDesc4, 1, 10); " " + Using("$########.##", PayRec(1).LICPAID4)
      'Print #RHandle, Tab(25);
    End If
    If PayRec(1).LICPAID5 <> 0 Then
      Print #RHandle, Mid$(PayRec(1).CatDesc5, 1, 10); " " + Using("$########.##", PayRec(1).LICPAID5)
      'Print #RHandle, Tab(25);
    End If
    If fpcurrPenAmtTot <> 0 Then
      Print #RHandle, " Penalty Fee"; " " + Using("$########.##", fpcurrPenAmtTot)
    End If
    If fpcurrIssAmt <> 0 Then
      Print #RHandle, "Issuance Fee"; " " + Using("$########.##", fpcurrIssAmt)
    End If
  Print #RHandle,
  Print #RHandle, "Set Renewal Flag - " + fpcmbSetFlag.Text
  Print #RHandle,
  Print #RHandle, "Operator: "; OperNum
  Print #RHandle, "Receipt#: "; Using("######", CmNum&)
  Print #RHandle,
  Print #RHandle, "       T H A N K   Y O U !"
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  If CntrlDef = 1 Then
    Print #RHandle, CutPaper$
  Else
    Print #RHandle,
    Print #RHandle,
    Print #RHandle,
  End If
  Close RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
If CntrlDef = 1 Then
  fpcmdDrawer_Click
End If
  'PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
10:
  DefPrinter = RecpPort '"LPT" + QPTrim$(Str$(RecpPort)) + ":"
20:
 ' MsgBox "Printer -" + DefPrinter, vbOKOnly
  
  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
        
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
50:
        Exit Do
        'Printer.EndDoc
      End If
    Loop Until eof(RptHandle)
60:
    Close RptHandle
62:
    Close LPTHandle
65:
    Next CopyLoop
68:
 Printer.EndDoc
69:
Validationthing:
   If QPTrim$(fpcmbType.Text) = "Check" Or QPTrim$(fpcmbType.Text) = "Cash & Check" Then
   If RctValidate And PrnVali = True Then
     DefPrinter = RecpPort
     RHandle2 = FreeFile
     Open PayRecpName2$ For Output As RHandle2
     Print #RHandle2, Chr$(27); Chr$(&H63); Chr$(&H30); Chr$(&H4)
     Print #RHandle2, Chr$(13); Chr$(10)
     Print #RHandle2, Tab(12); TownName$
     Print #RHandle2, Tab(12); "Bank- "; BnkAcctNum$
     Print #RHandle2, Tab(12); "FOR DEPOSIT ONLY"
     Print #RHandle2, Tab(12); "Acct. No. "; PayRec(1).CustNumber
     Print #RHandle2, Tab(12); "Date: "; PostDate$
     Print #RHandle2, Tab(12); "Time: "; Time
     Print #RHandle2,
     Print #RHandle2, Chr$(12)
     Close RHandle2
     LPTHandle = FreeFile

     Open DefPrinter For Output As LPTHandle
     RptHandle2 = FreeFile
     Open PayRecpName2$ For Input As RptHandle2
     Do
       If frmPrint.cmdCancel = False Then
         Line Input #RptHandle2, ToPrint$
         ToPrint$ = RTrim$(ToPrint$)
         Print #LPTHandle, ToPrint$
       Else
         Exit Do
       End If
     Loop Until eof(RptHandle2)
     Close RptHandle2
     Close LPTHandle
    Printer.EndDoc
    BLLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(PayRec(1).CustNumber)
    CMLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(PayRec(1).CustNumber)
  End If
 End If

70:
If PrnRecp = True Then
 BLLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(PayRec(1).CustNumber)
 CMLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(PayRec(1).CustNumber)
 KillFile PayRecpName$
 KillFile PayFileName$
End If
80:
  Exit Sub
Cancel:
  If Err > 0 Then
    CMLog "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

  
End Sub
