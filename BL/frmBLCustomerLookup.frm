VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustomerLookup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customer Lookup"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCustomerLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbCategory 
      Height          =   405
      Left            =   3120
      TabIndex        =   6
      Tag             =   "Select from one of the categories to narrow your search down to a particular category."
      Top             =   4200
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   714
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
      Object.TabStop         =   -1  'True
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmBLCustomerLookup.frx":08CA
   End
   Begin LpLib.fpList fpListSearch 
      Height          =   3360
      Left            =   930
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   $"frmBLCustomerLookup.frx":0C51
      Top             =   4950
      Width           =   9810
      _Version        =   196608
      _ExtentX        =   17304
      _ExtentY        =   5927
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
      Columns         =   3
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
      ColDesigner     =   "frmBLCustomerLookup.frx":0CEF
   End
   Begin EditLib.fpText fptxtCustNum 
      Height          =   390
      Left            =   3135
      TabIndex        =   0
      Tag             =   $"frmBLCustomerLookup.frx":1083
      Top             =   1590
      Width           =   1305
      _Version        =   196608
      _ExtentX        =   2307
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
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
   Begin EditLib.fpText fptxtBusName 
      Height          =   390
      Left            =   3120
      TabIndex        =   2
      Tag             =   $"frmBLCustomerLookup.frx":1261
      Top             =   2112
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
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
   Begin EditLib.fpText fptxtBillName 
      Height          =   390
      Left            =   3120
      TabIndex        =   3
      Tag             =   $"frmBLCustomerLookup.frx":1451
      Top             =   2634
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
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
   Begin EditLib.fpText fptxtSearchName 
      Height          =   390
      Left            =   6495
      TabIndex        =   1
      Tag             =   $"frmBLCustomerLookup.frx":1611
      Top             =   1590
      Width           =   2505
      _Version        =   196608
      _ExtentX        =   4419
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
   Begin EditLib.fpText fptxtContact 
      Height          =   390
      Left            =   3120
      TabIndex        =   4
      Tag             =   $"frmBLCustomerLookup.frx":17FE
      Top             =   3156
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
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
      Height          =   540
      Left            =   8160
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "Press the 'Exit' button to leave this screen and return to the main Customer Maintenance menu."
      Top             =   2400
      Width           =   2250
      _Version        =   131072
      _ExtentX        =   3969
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustomerLookup.frx":1A1D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   540
      Left            =   8160
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   $"frmBLCustomerLookup.frx":1BFB
      Top             =   3135
      Width           =   2250
      _Version        =   131072
      _ExtentX        =   3969
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustomerLookup.frx":1DCB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   8160
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   $"frmBLCustomerLookup.frx":1FA9
      Top             =   3885
      Width           =   2250
      _Version        =   131072
      _ExtentX        =   3969
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustomerLookup.frx":203A
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   345
      Left            =   480
      TabIndex        =   17
      Top             =   600
      Width           =   795
      _Version        =   131072
      _ExtentX        =   1402
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   3000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin EditLib.fpText fptxtServAdd 
      Height          =   390
      Left            =   3120
      TabIndex        =   5
      Tag             =   $"frmBLCustomerLookup.frx":221D
      Top             =   3678
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      Left            =   1800
      TabIndex        =   20
      Top             =   4320
      Width           =   1170
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
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
      Left            =   1125
      TabIndex        =   19
      Top             =   3800
      Width           =   1890
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8235
      TabIndex        =   18
      Top             =   4470
      Width           =   2100
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer LookUp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2819
      TabIndex        =   13
      Top             =   633
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1386
      Top             =   498
      Width           =   8655
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3465
      Left            =   735
      Top             =   1395
      Width           =   10200
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number:"
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
      Left            =   795
      TabIndex        =   12
      Top             =   1710
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Name:"
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
      Left            =   1260
      TabIndex        =   11
      Top             =   2190
      Width           =   1755
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Name:"
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
      Left            =   1455
      TabIndex        =   10
      Top             =   2700
      Width           =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   4770
      TabIndex        =   9
      Top             =   1710
      Width           =   1605
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name:"
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
      Left            =   1365
      TabIndex        =   8
      Top             =   3240
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1386
      Top             =   438
      Width           =   8655
   End
End
Attribute VB_Name = "frmBLCustomerLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  KillFile "custlookup.dat"
  frmBLCustMaintMenu.Show
  DoEvents
  Unload frmBLCustomerLookup
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtCustNum.ToolTipText = ""
    fptxtBusName.ToolTipText = ""
    fptxtBillName.ToolTipText = ""
    fptxtSearchName.ToolTipText = ""
    fptxtContact.ToolTipText = ""
    fptxtServAdd.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSearch.ToolTipText = ""
    fpListSearch.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtCustNum.ToolTipText = "Enter an entire customer number to search for a specific customer. Or enter a partial number to search for all customers with that partial number in their customer numbers."
'    fptxtBusName.ToolTipText = "Enter an entire business name to search for a specific customer. Or enter a partial business name to search for all customers with that partial name in their customer names."
'    fptxtBillName.ToolTipText = "Enter an entire billing name to search for a specific customer. Or enter a partial billing name to search for all customers with that partial name in their billing name."
'    fptxtSearchName.ToolTipText = "Enter an entire customer search name to search for a specific customer. Or enter a partial search name to search for all customers with that partial name in their search name."
'    fptxtContact.ToolTipText = "Enter an entire customer contact name to search for a specific customer. Or enter a partial contact name to search for all customers with that partial name in their contact name."
'    fptxtServAdd.ToolTipText = "Enter an entire service address to search for a specific customer. Otherwise, enter a partial service address to search for all customers with that partial service address in their service address."
'    cmdHelp.ToolTipText = "Press to bring up a brief help screen."
'    cmdExit.ToolTipText = "Press to exit this screen."
'    cmdSearch.ToolTipText = "Press to activate search procedure."
'    fpListSearch.ToolTipText = "The program lists businesses matching the criteria entered above."
  End If
End Sub

Private Sub Form_Load()
  Dim One As Integer
  Dim DHandle As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim CodeIdxRec As CatCodeIdxType
  Dim IdxHandle As Integer
  Dim x As Integer
  Dim NumOfCodes As Integer
  
  If Not Exist("arcode.dat") Then 'no file there
    fpcmbCategory.Visible = False
    Label7.Visible = False
    GoTo NoCode
  End If
  
  OpenCatCodeIdxFile IdxHandle
  NumOfCodes = LOF(IdxHandle) / Len(CodeIdxRec)
  ReDim IdxRec(1 To NumOfCodes) As Integer
  For x = 1 To NumOfCodes
    Get IdxHandle, x, CodeIdxRec
      IdxRec(x) = CodeIdxRec.CatCodeRec
  Next x
  Close IdxHandle
  
  fpcmbCategory.Clear
  OpenCatCodeFile CodeHandle
  
  For x = 1 To NumOfCodes
    Get CodeHandle, IdxRec(x), CodeRec
      fpcmbCategory.AddItem CodeRec.CatCode + Chr(9) + CodeRec.CODEDESC
  Next x
  Close CodeHandle
  
NoCode:
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FromCustEdit = False
  lblBalloon.Visible = False
'  fptxtCustNum.ToolTipText = "Enter an entire customer number to search for a specific customer. Or enter a partial number to search for all customers with that partial number in their customer numbers."
'  fptxtBusName.ToolTipText = "Enter an entire business name to search for a specific customer. Or enter a partial business name to search for all customers with that partial name in their customer names."
'  fptxtBillName.ToolTipText = "Enter an entire billing name to search for a specific customer. Or enter a partial billing name to search for all customers with that partial name in their billing name."
'  fptxtSearchName.ToolTipText = "Enter an entire customer search name to search for a specific customer. Or enter a partial search name to search for all customers with that partial name in their search name."
'  fptxtContact.ToolTipText = "Enter an entire customer contact name to search for a specific customer. Or enter a partial contact name to search for all customers with that partial name in their contact name."
'  fptxtServAdd.ToolTipText = "Enter an entire service address to search for a specific customer. Otherwise, enter a partial service address to search for all customers with that partial service address in their service address."
'  cmdHelp.ToolTipText = "Press to bring up a brief help screen."
'  cmdExit.ToolTipText = "Press to exit this screen."
'  cmdSearch.ToolTipText = "Press to activate search procedure."
'  fpListSearch.ToolTipText = "The program lists businesses matching the criteria entered above."
  One = 1
  DHandle = FreeFile
  Open "custlookup.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpListSearch.ListIndex <> -1 Then
      GoTo CustAlreadySelected
    ElseIf QPTrim$(fptxtCustNum.Text) <> "" Or QPTrim$(fptxtBusName.Text) <> "" Or QPTrim$(fptxtBillName.Text) <> "" _
    Or QPTrim$(fptxtSearchName.Text) <> "" Or QPTrim$(fptxtContact.Text) <> "" Or QPTrim$(fptxtServAdd.Text) <> "" Then
      Call cmdSearch_Click
      KeyCode = 0
      Exit Sub
    Else
      SendKeys "{Tab}"
      KeyCode = 0
      Exit Sub
    End If
CustAlreadySelected:
    fpListSearch.Col = 1
    If QPTrim$(fpListSearch.ColText) = "" Then
      frmBLMessageBoxJr.Label1.Caption = "No Customer has been selected"
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    Else
      Call fpListSearch_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown:
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
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
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
      KillFile "custrptsmenu.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustomerLookup.")
      Call Terminate
      End
    End If
  End If
End Sub

Public Sub cmdSearch_Click()
   Dim CustRec As ARCustRecType
   Dim CustIdxRec As CustNameIdxType
   Dim CustIdxHandle As Integer
   Dim CustIdxRecNum As Integer
   Dim CHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim cnt As Integer
   Dim CustCnt As Integer
   Dim TempCustNum$
   Dim TempBusName$
   Dim TempBillName$
   Dim TempSearchName$
   Dim TempContactName$
   Dim TempServAdd$
   Dim TempCatCode$
   Dim NumFlag As Boolean
   Dim BusNameFlag As Boolean
   Dim BillNameFlag As Boolean
   Dim SearchFlag As Boolean
   Dim ContactFlag As Boolean
   Dim ServiceFlag As Boolean
   Dim CatFlag As Boolean
   Dim Found As Boolean
   Dim MatchCnt As Integer
   Dim FoundCnt As Integer
   Dim OnlyOneFound$
   Dim HoldGCustNum As Integer
   Dim ThisRow As Integer
   
   On Error GoTo ERRORSTUFF
   
   If GCustNum > 0 Then
     HoldGCustNum = GCustNum
   End If
   
   fpListSearch.Clear
   
   NumFlag = False
   BusNameFlag = False
   BillNameFlag = False
   SearchFlag = False
   ContactFlag = False
   ServiceFlag = False
   CatFlag = False
   
   If Not Exist("arcustnameidx.dat") Then 'no file there
     frmBLMessageBoxJr.Label1.Caption = "No Customer Name Index has been saved."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Exit Sub
   End If
   
   OpenCustNameIdxFile CustIdxHandle
   CustIdxRecNum = LOF(CustIdxHandle) \ Len(CustIdxRec)
   If CustIdxRecNum = 0 Then 'file is there but there is nothing in it
     frmBLMessageBoxJr.Label1.Caption = "No Customers in index."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Close
     Exit Sub
   End If
   
   ReDim CustIdx(1 To CustIdxRecNum) As Integer
   For x = 1 To CustIdxRecNum
     Get CustIdxHandle, x, CustIdxRec
     CustIdx(x) = CustIdxRec.CustRec 'load array with record pointers
   Next x
   Close CustIdxHandle
   
   If Not Exist("ARCUST.DAT") Then
     frmBLMessageBoxJr.Label1.Caption = "Path to ARCUST.DAT could not be found"
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Exit Sub
   End If
   
   If QPTrim$(fptxtCustNum.Text) <> "" Then
     TempCustNum$ = QPTrim$(fptxtCustNum.Text)
     NumFlag = True
   End If
   
   If QPTrim$(fptxtBusName.Text) <> "" Then
     TempBusName = QPTrim$(fptxtBusName.Text)
     BusNameFlag = True
   End If
   
   If QPTrim$(fptxtBillName.Text) <> "" Then
     TempBillName = QPTrim$(fptxtBillName.Text)
     BillNameFlag = True
   End If
   
   If QPTrim$(fptxtSearchName.Text) <> "" Then
     TempSearchName = QPTrim$(fptxtSearchName.Text)
     SearchFlag = True
   End If
   
   If QPTrim$(fptxtContact.Text) <> "" Then
     TempContactName = QPTrim$(fptxtContact.Text)
     ContactFlag = True
   End If
   
   If QPTrim$(fptxtServAdd.Text) <> "" Then
     TempServAdd = QPTrim$(fptxtServAdd.Text)
     ServiceFlag = True
   End If
   
   If QPTrim$(fpcmbCategory.Text) <> "" Then
     fpcmbCategory.Col = 0
     fpcmbCategory.Row = -1
     TempCatCode = QPTrim$(fpcmbCategory.ColText)
     CatFlag = True
   End If
   
   OpenCustFile CHandle
   CustCnt = LOF(CHandle) / Len(CustRec)
   
   If CustCnt = 0 Then
     frmBLMessageBoxJr.Label1.Caption = "No Customer data on file."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Close
     Exit Sub
   End If
   
   For x = 1 To CustIdxRecNum
     Get CHandle, CustIdx(x), CustRec
     If QPTrim$(CustRec.SortName) = "DELETED" Or QPTrim$(CustRec.Deleted) = "Y" Then GoTo NotAMatch
     Found = True
     If NumFlag = True Then
       If InStr(UCase$(CustRec.CustNumb), TempCustNum) > 0 Then ' And Len(QPTrim$(CustRec.CustNumb)) = Len(QPTrim$(TempCustNum)) Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If BusNameFlag = True Then
       If InStr(UCase$(CustRec.CustName), TempBusName) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If BillNameFlag = True Then
       If InStr(UCase$(CustRec.BillName), TempBillName) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If SearchFlag = True Then
       If InStr(UCase$(CustRec.SortName), TempSearchName) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If ContactFlag = True Then
       If InStr(UCase$(CustRec.Contact), TempContactName) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If ServiceFlag = True Then
       If InStr(UCase$(CustRec.ServAdd), TempServAdd) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If CatFlag = True Then
       If QPTrim(CustRec.BILLCAT1) = TempCatCode Then
         Found = True
       ElseIf QPTrim(CustRec.BILLCAT2) = TempCatCode Then
         Found = True
       ElseIf QPTrim(CustRec.BILLCAT3) = TempCatCode Then
         Found = True
       ElseIf QPTrim(CustRec.BILLCAT4) = TempCatCode Then
         Found = True
       ElseIf QPTrim(CustRec.BILLCAT5) = TempCatCode Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
    
    If Found Then
      FoundCnt = FoundCnt + 1
      fpListSearch.Row = -1
      If HoldGCustNum > 0 Then
        If CustIdx(x) = HoldGCustNum Then ThisRow = MatchCnt
      End If
      MatchCnt = MatchCnt + 1
      GCustNum = CustIdx(x)
      fpListSearch.InsertRow = QPTrim$(CustRec.BillName) & Chr$(9) & " " & QPTrim$(CustRec.City) & Chr$(9) & "  " & QPTrim$(CustRec.CustNumb)
      'only used if no more than one found
      OnlyOneFound = QPTrim$(CustRec.CustNumb)
    End If
NotAMatch:
  Next x
  
  If FoundCnt > 1 Then
    fpListSearch.SetFocus
  End If
   
  If MatchCnt <= 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No match found."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
    Close
  End If
  
  If HoldGCustNum > 0 Then
    fpListSearch.ListIndex = ThisRow
  ElseIf FoundCnt > 1 Then
    fpListSearch.ListIndex = 0
  End If
  
  If CustCnt = 1 Then 'the program would exit all the way back to the
  'desktop when after saving the first customer you went back to edit and save
  'this one and only customer
    If QPTrim$(fptxtBillName.Text) = "" And QPTrim$(fptxtBusName.Text) = "" And _
      QPTrim$(fptxtContact.Text) = "" And QPTrim$(fptxtCustNum.Text) = "" And _
      QPTrim$(fptxtSearchName.Text) = "" And QPTrim$(fptxtServAdd.Text) = "" Then
        Close
        Exit Sub
    End If
  End If
  
  If FoundCnt = 1 Then
    For x = 1 To CustCnt
      Get CHandle, CustIdx(x), CustRec
      If OnlyOneFound = QPTrim$(CustRec.CustNumb) Then
        GCustNum = CustIdx(x)
        Exit For
      Else
        Found = False
        GoTo NotThisTime
      End If
   
NotThisTime:
    Next x
    
    fptxtBillName.Text = ""
    fptxtBusName.Text = ""
    fptxtContact.Text = ""
    fptxtCustNum.Text = ""
    fptxtSearchName.Text = ""
    fptxtServAdd.Text = ""
    
    fpListSearch.Clear
    FoundCnt = 0
    
    frmBLCustEdit.Caption = "Business License Edit Item"
    frmBLCustEdit.Show
    DoEvents
    Unload frmBLCustomerLookup
  End If
  Close
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustomerLookup", "cmdSearch_Click", Erl)
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
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub


Private Sub fpcmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    fpcmbCategory.Text = ""
    Exit Sub
  End If
  
  If KeyCode = vbKeySpace Then
    fpcmbCategory.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCategory.ListIndex = -1
  End If
  If fpcmbCategory.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtCustNum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpListSearch_DblClick()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim cnt As Integer
  Dim CustCnt As Integer
  Dim City$
  Dim CustNum$
  Dim BillName$
  Dim Found As Boolean
  
  On Error GoTo ERRORSTUFF
  
  fpListSearch.Col = 0
  'trap for double clicking on nothing
  If QPTrim$(fpListSearch.ColText) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "No item has been selected"
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  BillName$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 1
  City$ = QPTrim$(fpListSearch.ColText)
  
  fpListSearch.Col = 2
  CustNum$ = QPTrim$(fpListSearch.ColText)
  
  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) \ Len(CustRec)
  For x = 1 To NumOfCustRecs
    Get CHandle, x, CustRec
    If InStr(UCase$(CustRec.BillName), UCase$(BillName$)) > 0 And InStr(UCase$(CustRec.City), UCase(City$)) > 0 And InStr(CustRec.CustNumb, CustNum$) > 0 Then
    'if two people had the same name and the emp number of one had a number that
    'included the other's (ie. 123 vs 1234) then then smaller number would not be accessed ever
      Found = True
      fpListSearch.Row = -1
      GCustNum = x
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close CHandle
  
  frmBLCustEdit.Show
  frmBLCustEdit.Caption = "Customer Edit"
  frmBLCustEdit.Label2 = "Customer Edit"
  DoEvents
  Me.Hide 'this code is different from the category code
  'lookup because of the transaction history module screen
  'available off the customer edit screen...if this screen
  'is visible then when the modal form closes it returns
  'automatically to this screen instead of to the customer
  'edit screen
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustomerLookup", "fpListSearch_DblClick", Erl)
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

Private Sub fptxtContact_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtCustNum.SetFocus
  End If
End Sub

Public Sub RefreshSearchList()
  Dim CHandle As Integer
  Dim CustRec As ARCustRecType
  Dim x As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim IdxRecNum As Integer
  Dim ThisRow As Integer
  
  fpListSearch.Action = ActionClear
  OpenCustNameIdxFile IdxHandle
  IdxRecNum = LOF(IdxHandle) \ Len(IdxRec)
  ReDim CustIdx(1 To IdxRecNum) As Integer
  For x = 1 To IdxRecNum
    Get IdxHandle, x, IdxRec
    CustIdx(x) = IdxRec.CustRec 'load array with record pointers
  Next x
  Close IdxHandle
  
  OpenCustFile CHandle
  For x = 1 To IdxRecNum
    Get CHandle, CustIdx(x), CustRec
    fpListSearch.InsertRow = QPTrim$(CustRec.BillName) & Chr$(9) & " " & QPTrim$(CustRec.City) & Chr$(9) & "  " & QPTrim$(CustRec.CustNumb)
    If CustIdx(x) = CInt(GCustNum) Then
      ThisRow = x
    End If
  Next x
  Close CHandle
  
  fpListSearch.ListIndex = ThisRow - 1
End Sub
