VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAItemLookUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Asset Item LookUp"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmItemLookUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpListSearch 
      Height          =   2790
      Left            =   435
      TabIndex        =   9
      Top             =   5370
      Width           =   10815
      _Version        =   196608
      _ExtentX        =   19076
      _ExtentY        =   4921
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      BorderWidth     =   2
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
      ColDesigner     =   "frmItemLookUp.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   690
      Left            =   6090
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list (below) of all items that fit within the parameters entered (above)."
      Top             =   3885
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmItemLookUp.frx":0CBE
   End
   Begin EditLib.fpText fptxtTagNumber 
      Height          =   396
      Left            =   4248
      TabIndex        =   0
      ToolTipText     =   $"frmItemLookUp.frx":0E9C
      Top             =   1356
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      CharValidationText=   ""
      MaxLength       =   150
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Left            =   4248
      TabIndex        =   1
      ToolTipText     =   $"frmItemLookUp.frx":0F54
      Top             =   1824
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      CharValidationText=   ""
      MaxLength       =   150
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
   Begin EditLib.fpText fpTxtSerialNum 
      Height          =   396
      Left            =   4248
      TabIndex        =   2
      ToolTipText     =   $"frmItemLookUp.frx":0FF1
      Top             =   2304
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      CharValidationText=   ""
      MaxLength       =   150
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
   Begin EditLib.fpText fptxtPONum 
      Height          =   396
      Left            =   4248
      TabIndex        =   3
      ToolTipText     =   $"frmItemLookUp.frx":10AB
      Top             =   2784
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      CharValidationText=   ""
      MaxLength       =   150
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
   Begin EditLib.fpText fptxtChkNum 
      Height          =   396
      Left            =   4248
      TabIndex        =   4
      ToolTipText     =   $"frmItemLookUp.frx":1165
      Top             =   3264
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      CharValidationText=   ""
      MaxLength       =   150
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
      Height          =   684
      Left            =   3450
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to create the desired report."
      Top             =   3888
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmItemLookUp.frx":1210
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "*D* = Disposed Of"
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
      Left            =   528
      TabIndex        =   13
      Top             =   5088
      Width           =   2172
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "*P* = Disposal Pending"
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
      Left            =   528
      TabIndex        =   12
      Top             =   4800
      Width           =   2172
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number:"
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
      Left            =   1716
      TabIndex        =   11
      Top             =   3360
      Width           =   2136
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "P.O. Number:"
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
      Left            =   1716
      TabIndex        =   10
      Top             =   2880
      Width           =   2136
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:"
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
      Left            =   1716
      TabIndex        =   8
      Top             =   2400
      Width           =   2136
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   1716
      TabIndex        =   7
      Top             =   1920
      Width           =   2136
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Number:"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   1548
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3492
      Left            =   888
      Top             =   1224
      Width           =   9816
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1380
      Top             =   300
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets Item LookUp"
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
      Left            =   2820
      TabIndex        =   5
      Top             =   432
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1380
      Top             =   240
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAItemLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFAItemMaintMenu.Show
  Close
  DoEvents
  Unload frmFAItemLookUp
End Sub

Public Sub cmdSearch_Click()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim Found As Boolean
  Dim TagFlag As Boolean
  Dim DescFlag As Boolean
  Dim SerialFlag As Boolean
  Dim PONumFlag As Boolean
  Dim ChkNumFlag As Boolean
  Dim TempTag$
  Dim TempDesc$
  Dim TempSerial$
  Dim TempPO$
  Dim TempChk$
  Dim FoundCnt As Integer
  Dim MatchCnt As Integer
  Dim PrintDesc$
  Dim OnlyOneFound$
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim Dspl$
  
  On Error GoTo ERRORSTUFF
  'this sub captures values from the screen the user entered to
  'be used as parameters and then looks for a fixed asset that
  'fits within all the parameters...if only one asset matches then
  'the edit screen automatically opens with that asset's data...if
  'multiple assets match then they are listed on the list so the
  'user can select from them
  fpListSearch.Clear
  
  TagFlag = False
  DescFlag = False
  SerialFlag = False
  PONumFlag = False
  ChkNumFlag = False
  
  If QPTrim$(fptxtTagNumber.Text) <> "" Then
    TagFlag = True
    TempTag = QPTrim$(fptxtTagNumber.Text)
  End If
  If QPTrim$(fptxtDesc.Text) <> "" Then
    DescFlag = True
    TempDesc = QPTrim$(fptxtDesc.Text)
  End If
  If QPTrim$(fptxtSerialNum.Text) <> "" Then
    SerialFlag = True
    TempSerial = QPTrim$(fptxtSerialNum.Text)
  End If
  If QPTrim$(fptxtPONum.Text) <> "" Then
    PONumFlag = True
    TempPO = QPTrim$(fptxtPONum.Text)
  End If
  If QPTrim$(fptxtChkNum.Text) <> "" Then
    ChkNumFlag = True
    TempChk = QPTrim$(fptxtChkNum.Text)
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfRecs = LOF(TagIdxHandle) \ Len(TagIdx)
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
  OpenFAItemFile FAHandle
  
  For x = 1 To NumOfRecs
    Get FAHandle, TagIdxRecs(x), FAItemRec
    Found = True
    If TagFlag = True Then
      If InStr(UCase$(FAItemRec.ItemTag), TempTag) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    If DescFlag = True Then
      If InStr(UCase$(FAItemRec.IDESC1), TempDesc) > 0 Or InStr(UCase$(FAItemRec.IDESC2), TempDesc) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    If SerialFlag = True Then
      If InStr(UCase$(FAItemRec.SERIALNO), TempSerial) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    If PONumFlag = True Then
      If InStr(UCase$(FAItemRec.PONum), TempPO) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    If ChkNumFlag = True Then
      If InStr(UCase$(FAItemRec.CheckNum), TempChk) > 0 Then
        Found = True
      Else
        Found = False
        GoTo NotAMatch
      End If
    End If
    If Found Then
      FoundCnt = FoundCnt + 1
      fpListSearch.Row = -1
      MatchCnt = MatchCnt + 1
      GRecNum = x
      If QPTrim$(FAItemRec.IDESC1) <> "" Then
        PrintDesc$ = QPTrim$(FAItemRec.IDESC1)
      Else
        PrintDesc$ = QPTrim$(FAItemRec.IDESC2)
      End If
      If FAItemRec.DsplFlag = 2 Then
        Dspl = "*D*"
      ElseIf FAItemRec.DsplFlag = 1 Then
        Dspl = "*P*"
      Else
        Dspl = ""
      End If
      fpListSearch.InsertRow = Dspl + QPTrim$(FAItemRec.ItemTag) & Chr$(9) & " " & PrintDesc$ & " " & Chr$(9) & " " & QPTrim$(FAItemRec.SERIALNO) & Chr$(9) & "  " & QPTrim$(FAItemRec.PONum) & Chr$(9) & " " & QPTrim$(FAItemRec.CheckNum)
      'only used if no more than one found
      OnlyOneFound = QPTrim$(FAItemRec.ItemTag)
    End If
NotAMatch:
  Next x
  
  'this screen is opening back up from a return from edit screen
  If Exist("edititemopen.dat") Then
    If FoundCnt >= 1 Then
      fpListSearch.SearchText = ThisTag
      fpListSearch.Action = ActionSearch
      If fpListSearch.SearchIndex <> -1 Then
        fpListSearch.ListIndex = fpListSearch.SearchIndex
        fpListSearch.SetFocus
      End If
    Else
      MsgBox "No match found."
    End If
    Close 'added 03/15/2004
    Exit Sub
  End If
  
  If MatchCnt <= 0 Then
    MsgBox "No match found"
    If TagFlag = True Then
      fptxtTagNumber.SetFocus
      Exit Sub
      Close
    ElseIf DescFlag = True Then
      fptxtDesc.SetFocus
      Exit Sub
      Close
    ElseIf SerialFlag = True Then
      fptxtSerialNum.SetFocus
      Exit Sub
      Close
    ElseIf PONumFlag = True Then
      fptxtPONum.SetFocus
      Exit Sub
      Close
    ElseIf ChkNumFlag = True Then
      fptxtChkNum.SetFocus
      Exit Sub
      Close
    End If
  End If
   
  'if only one match is found than the global GRecNum is assigned
  'the value of it's record and then the edit screen is loaded up
  'with the matching asset's data
  If FoundCnt = 1 Then
    For x = 1 To NumOfRecs
      Get FAHandle, x, FAItemRec
        If OnlyOneFound = QPTrim$(FAItemRec.ItemTag) Then
          GRecNum = x
          Exit For
        Else
          Found = False
          GoTo NotThisTime
        End If
NotThisTime:
    Next x
    
'    fptxtTagNumber.Text = ""
'    fptxtDesc.Text = ""
'    fptxtSerialNum.Text = ""
'    fpListSearch.Clear
    FoundCnt = 0
    frmFAEditItemWTabs.Caption = "Fixed Asset Edit Item"
    frmFAEditItemWTabs.Label2 = "Fixed Asset Edit Item"
    frmFAEditItemWTabs.Show
    DoEvents
    frmFAItemLookUp.Hide
  End If
  Close FAHandle
  
  If FoundCnt > 1 Then
    fpListSearch.SetFocus
  End If
  
  fpListSearch.ListIndex = 0
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemLookUp", "cmdSearch_Click", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpListSearch.ListIndex <> -1 Then GoTo EmpAlreadySelected '8/6
    If Len(fptxtTagNumber.Text) > 0 Or Len(fptxtDesc.Text) > 0 Or Len(fptxtSerialNum.Text) > 0 Then
      Call cmdSearch_Click
      KeyCode = 0
      Exit Sub
    End If
EmpAlreadySelected:
    fpListSearch.Col = 1
    If QPTrim$(fpListSearch.ColText) = "" Then
      MsgBox "No item has been selected"
      Exit Sub
    Else
      Call fpListSearch_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSearch_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAItemLookUp.")
      Call Terminate
      End
    End If
  End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpListSearch_DblClick()
  Dim FAHandle As Integer
  Dim NumOfRecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim x As Long
  Dim TagNum$
  Dim Desc$
  Dim SerialNum$
  Dim PONum$
  Dim ChkNum$
  Dim PrintDesc$
  Dim Found As Boolean
  
  On Error GoTo ERRORSTUFF
  fpListSearch.Col = 0
  'trap for double clicking on nothing
  If QPTrim$(fpListSearch.ColText) = "" Then
    MsgBox "No item has been selected"
    Exit Sub
  End If
  TagNum$ = QPTrim$(fpListSearch.ColText)
  
  If Mid(TagNum$, 1, 3) = "*D*" Or Mid(TagNum$, 1, 3) = "*P*" Then
    TagNum = Mid(TagNum$, 4)
  End If
  
  fpListSearch.Col = 1
  Desc$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 2
  SerialNum$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 3
  PONum$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 4
  ChkNum$ = QPTrim$(fpListSearch.ColText)
  
  OpenFAItemFile FAHandle
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  For x = 1 To NumOfRecs
    Get FAHandle, x, FAItemRec
    If QPTrim$(FAItemRec.IDESC1) <> "" Then
      PrintDesc$ = QPTrim$(FAItemRec.IDESC1)
    Else
      PrintDesc$ = QPTrim$(FAItemRec.IDESC2)
    End If
  
    If InStr(UCase$(FAItemRec.ItemTag), TagNum$) > 0 And InStr(UCase$(PrintDesc$), Desc$) > 0 And InStr(FAItemRec.SERIALNO, SerialNum$) >= 0 _
    And Len(QPTrim$(FAItemRec.ItemTag)) = Len(QPTrim$(TagNum$)) And InStr(FAItemRec.PONum, PONum) > 0 And InStr(FAItemRec.CheckNum, ChkNum$) > 0 Then '8/7 added Len = Len because
    'if two people had the same name and the emp number of one had a number that
    'included the other's (ie. 123 vs 1234) then then smaller number would not be accessed ever
      Found = True
      fpListSearch.Row = -1
      GRecNum = x
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close FAHandle
  
  frmFAEditItemWTabs.Show
  frmFAEditItemWTabs.Caption = "Fixed Asset Edit Item"
  frmFAEditItemWTabs.Label2 = "Fixed Asset Edit Item"
  DoEvents
  frmFAItemLookUp.Hide
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemLookUp", "fpListSearch_DblClick", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

