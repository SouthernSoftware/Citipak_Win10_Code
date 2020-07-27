VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFAEditItem 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmFAEditItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbStatus 
      Height          =   384
      Left            =   2256
      TabIndex        =   5
      ToolTipText     =   "Enter the active status of this fixed asset (required)."
      Top             =   3168
      Width           =   3132
      _Version        =   196608
      _ExtentX        =   5524
      _ExtentY        =   677
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
      ColDesigner     =   "frmFAEditItem.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbDepYN 
      Height          =   384
      Left            =   7632
      TabIndex        =   18
      ToolTipText     =   "Enter a Y if this fixed asset will be depreciated or N if this fixed asset will not be depreciated (required)."
      Top             =   2112
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
      _ExtentY        =   677
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
      ColDesigner     =   "frmFAEditItem.frx":0B8D
   End
   Begin VB.CommandButton cmdVehicle 
      Caption         =   "F6 &Vehicle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   5952
      TabIndex        =   58
      ToolTipText     =   "Click this button to bring up a screen on which you can enter data specific to a vehicle."
      Top             =   7776
      Width           =   1692
   End
   Begin VB.CommandButton cmdFundList 
      Caption         =   "F7 &Fund List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   4224
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all fund numbers."
      Top             =   2112
      Width           =   1548
   End
   Begin VB.CommandButton cmdDept 
      Caption         =   "F8 &Dept List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   4224
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all fixed assets."
      Top             =   4752
      Width           =   1548
   End
   Begin VB.CommandButton cmdTagList 
      Caption         =   "F9 I&tem List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   4224
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all fixed assets."
      Top             =   1584
      Width           =   1548
   End
   Begin VB.CommandButton cmdAssetList 
      Caption         =   "F11 Code &List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   4224
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all asset codes."
      Top             =   5280
      Width           =   1548
   End
   Begin EditLib.fpCurrency fptxtOriginalCost 
      Height          =   396
      Left            =   7632
      TabIndex        =   21
      ToolTipText     =   "Enter the amount paid for this fixed asset (required)."
      Top             =   2640
      Width           =   3276
      _Version        =   196608
      _ExtentX        =   5778
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10 &SAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   9696
      TabIndex        =   32
      ToolTipText     =   "Click this button to save all data entered above."
      Top             =   7776
      Width           =   1692
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   7824
      TabIndex        =   31
      ToolTipText     =   "Click this button to return to the Item Lookup screen."
      Top             =   7776
      Width           =   1692
   End
   Begin EditLib.fpText fptxtTagNumber 
      Height          =   396
      Left            =   2256
      TabIndex        =   0
      ToolTipText     =   "Enter the tag number here."
      Top             =   1584
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      CharValidationText=   "0 1 2 3 4 5 6 7 8 9 - "
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
   Begin EditLib.fpText fptxtDesc1 
      Height          =   396
      Left            =   2256
      TabIndex        =   6
      ToolTipText     =   "Enter a brief description of this fixed asset here (required)."
      Top             =   3696
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
   Begin EditLib.fpText fptxtDesc2 
      Height          =   396
      Left            =   2256
      TabIndex        =   7
      ToolTipText     =   "Enter a brief description for this fixed asset here (optional)."
      Top             =   4224
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
   Begin EditLib.fpText fptxtAssetLife 
      Height          =   396
      Left            =   9024
      TabIndex        =   19
      ToolTipText     =   "Enter the expected number of years this asset should be of value."
      Top             =   2100
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
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
      CharValidationText=   "0 , 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9"
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
   Begin EditLib.fpText fptxtGLNum 
      Height          =   396
      Left            =   2256
      TabIndex        =   4
      ToolTipText     =   "Enter the desired general ledger number here (optional)."
      Top             =   2640
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - "
      MaxLength       =   14
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
   Begin EditLib.fpText fptxtDeptNum 
      Height          =   396
      Left            =   2256
      TabIndex        =   8
      ToolTipText     =   "Enter a valid department number here (required)."
      Top             =   4752
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpText fptxtGroupCode 
      Height          =   396
      Left            =   2256
      TabIndex        =   10
      ToolTipText     =   "Enter a valid asset code here (required)."
      Top             =   5280
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpText fptxtVendorNum 
      Height          =   396
      Left            =   2256
      TabIndex        =   12
      ToolTipText     =   "Enter the vendor of this fixed asset here (optional)."
      Top             =   5808
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
   Begin EditLib.fpText fptxtSerialNum 
      Height          =   396
      Left            =   2256
      TabIndex        =   13
      ToolTipText     =   "Enter the serial number for this fixed asset here (optional)."
      Top             =   6336
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
   Begin EditLib.fpText fptxtMfg 
      Height          =   396
      Left            =   2256
      TabIndex        =   14
      ToolTipText     =   "Enter the manufacturer of this fixed asset here (optional)."
      Top             =   6864
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
   Begin EditLib.fpText fptxtContact 
      Height          =   396
      Left            =   2256
      TabIndex        =   15
      ToolTipText     =   "Enter the contact person for this fixed asset here (optional)."
      Top             =   7392
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
   Begin EditLib.fpText fptxtLocation 
      Height          =   396
      Left            =   7632
      TabIndex        =   30
      ToolTipText     =   "Enter the location where this fixed asset can be found (optional)."
      Top             =   6864
      Width           =   3276
      _Version        =   196608
      _ExtentX        =   5778
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
   Begin EditLib.fpDateTime fptxtAcquiredDate 
      Height          =   348
      Left            =   8160
      TabIndex        =   17
      ToolTipText     =   "Enter the date on which this  fixed asset was purchased (required)."
      Top             =   1584
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   614
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
   Begin EditLib.fpDateTime fptxtCurrDepDate 
      Height          =   348
      Left            =   7632
      TabIndex        =   26
      ToolTipText     =   "This date, the most current depreciation date, is automatically calculated and cannot be edited."
      Top             =   4752
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   614
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
   Begin EditLib.fpDateTime fptxtDisposalDate 
      Height          =   348
      Left            =   7632
      TabIndex        =   28
      ToolTipText     =   "This date indicates the date this item was disposed of and is figured in the disposal process. It is not editable here."
      Top             =   5808
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   614
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
      HideSelection   =   0   'False
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   "01/14/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19200101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
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
   Begin EditLib.fpDateTime fptxtEOLDate 
      Height          =   348
      Left            =   7632
      TabIndex        =   27
      ToolTipText     =   "This date, the End Of Life date, is automatically calculated and cannot be edited."
      Top             =   5280
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   614
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
   Begin EditLib.fpCurrency fptxtDep2Date 
      Height          =   396
      Left            =   7632
      TabIndex        =   24
      ToolTipText     =   "This field is automatically calculated when this fixed asset is depreciated. This value can be edited (but not recommended.)"
      Top             =   3696
      Width           =   3276
      _Version        =   196608
      _ExtentX        =   5778
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
   Begin EditLib.fpCurrency fptxtCurrVal 
      Height          =   396
      Left            =   7632
      TabIndex        =   25
      ToolTipText     =   "This value is automatically calculated. It cannot be edited."
      Top             =   4224
      Width           =   3276
      _Version        =   196608
      _ExtentX        =   5778
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
   Begin EditLib.fpCurrency fptxtDispPrice 
      Height          =   396
      Left            =   7632
      TabIndex        =   29
      ToolTipText     =   "This value is calculated in the disposal process and is not editable here."
      Top             =   6336
      Width           =   3276
      _Version        =   196608
      _ExtentX        =   5778
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
   Begin EditLib.fpText fptxtFundNum 
      Height          =   396
      Left            =   2256
      TabIndex        =   2
      ToolTipText     =   "Enter the General Ledger fund number here (required)."
      Top             =   2112
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      CharValidationText=   "0 , 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9"
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
   Begin EditLib.fpText fptxtLeft 
      Height          =   396
      Left            =   10320
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "This field displays the life remaining for this fixed asset."
      Top             =   2100
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "0 , 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9"
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
      Left            =   7632
      TabIndex        =   22
      ToolTipText     =   "Enter the purchase order number for this fixed asset here (optional)."
      Top             =   3168
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
   Begin EditLib.fpText fptxtChkNum 
      Height          =   396
      Left            =   9600
      TabIndex        =   23
      ToolTipText     =   "Enter the check number used to pay for this fixed asset here (optional)."
      Top             =   3168
      Width           =   1308
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
   Begin EditLib.fpDateTime fpDateWrntyX 
      Height          =   348
      Left            =   2256
      TabIndex        =   16
      ToolTipText     =   "Enter the expiration date for the warranty for this fixed asset (optional)."
      Top             =   7920
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   614
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
      ControlType     =   0
      Text            =   "03/19/2003"
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
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   288
      Y1              =   1356.211
      Y2              =   8277.564
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   312
      X2              =   11410
      Y1              =   1356.211
      Y2              =   1356.211
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wrnty Expires:"
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
      Left            =   432
      TabIndex        =   63
      Top             =   7968
      Width           =   1692
   End
   Begin VB.Label lblDspl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1500
      Left            =   9648
      TabIndex        =   62
      Top             =   4752
      Width           =   1260
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
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
      Left            =   9408
      TabIndex        =   61
      Top             =   3264
      Width           =   252
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PO/Chk Number:"
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
      Left            =   5424
      TabIndex        =   60
      Top             =   3264
      Width           =   2028
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Left:"
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
      Left            =   9600
      TabIndex        =   59
      Top             =   2196
      Width           =   684
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   11424
      X2              =   11424
      Y1              =   1356.211
      Y2              =   7295.48
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   5704
      Y1              =   8277.564
      Y2              =   8277.564
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   5700
      X2              =   11420
      Y1              =   7295.48
      Y2              =   7295.48
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Fund*:"
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
      Left            =   972
      TabIndex        =   57
      Top             =   2208
      Width           =   1116
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Disposal Price:"
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
      Left            =   5808
      TabIndex        =   56
      Top             =   6432
      Width           =   1644
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Required fields denoted with *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Left            =   240
      TabIndex        =   55
      Top             =   1056
      Width           =   2796
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
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
      Left            =   6384
      TabIndex        =   54
      Top             =   6960
      Width           =   1068
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
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
      Left            =   1056
      TabIndex        =   53
      Top             =   7488
      Width           =   1020
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer:"
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
      Left            =   528
      TabIndex        =   52
      Top             =   6960
      Width           =   1548
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Num:"
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
      Left            =   720
      TabIndex        =   51
      Top             =   6432
      Width           =   1356
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EOL Date:"
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
      Left            =   6288
      TabIndex        =   50
      Top             =   5328
      Width           =   1164
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor:"
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
      Left            =   864
      TabIndex        =   49
      Top             =   5904
      Width           =   1212
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Disposal Date:"
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
      Left            =   5856
      TabIndex        =   48
      Top             =   5856
      Width           =   1596
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Dpr Date:"
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
      Left            =   5808
      TabIndex        =   47
      Top             =   4800
      Width           =   1644
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Value:"
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
      Left            =   5760
      TabIndex        =   46
      Top             =   4320
      Width           =   1692
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deprec To Date:"
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
      Height          =   348
      Left            =   5568
      TabIndex        =   45
      Top             =   3792
      Width           =   1884
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code*:"
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
      Left            =   528
      TabIndex        =   44
      Top             =   5376
      Width           =   1548
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dept. Num*:"
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
      Left            =   576
      TabIndex        =   43
      Top             =   4848
      Width           =   1500
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "G/L Acct:"
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
      Left            =   816
      TabIndex        =   42
      Top             =   2736
      Width           =   1260
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Life*:"
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
      Left            =   8256
      TabIndex        =   41
      Top             =   2196
      Width           =   684
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Price*:"
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
      Left            =   5568
      TabIndex        =   40
      Top             =   2736
      Width           =   1884
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depreciate*?"
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
      Left            =   5952
      TabIndex        =   39
      Top             =   2196
      Width           =   1452
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status*:"
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
      Left            =   1152
      TabIndex        =   38
      Top             =   3264
      Width           =   924
   End
   Begin VB.Label Label4 
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
      Left            =   672
      TabIndex        =   37
      Top             =   4320
      Width           =   1404
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description*:"
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
      Left            =   624
      TabIndex        =   36
      Top             =   3792
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acquired On*:"
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
      Left            =   6432
      TabIndex        =   35
      Top             =   1632
      Width           =   1548
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Number*:"
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
      Left            =   492
      TabIndex        =   34
      Top             =   1692
      Width           =   1584
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   240
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   33
      Top             =   384
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   192
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   5700
      X2              =   5700
      Y1              =   8277.564
      Y2              =   7295.48
   End
End
Attribute VB_Name = "frmFAEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim BadGLNum As Boolean
  Dim BadAssetCodeNum As Boolean
  Dim BadTagNum As Boolean
  Dim FirstTagNum$
  Dim AssLife As Integer
  Dim AssLifeLeft As Integer
  Dim AcqDate As Integer
  Dim TempTagNumber$
  Dim TempItemTag$
  Dim TempISTATUS$
  Dim TempDEPYN$
  Dim TempAQURDATE As Integer
  Dim TempIDESC1$
  Dim TempIDESC2$
  Dim TempGLACCT$
  Dim TempIDEPT    As Integer
  Dim TempASSETCode$
  Dim TempILIFE    As Double
  Dim TempORGCOST  As Double
  Dim TempDEP2DATE As Double
  Dim TempCURRVAL  As Double
  Dim TempCDEPDATE As Integer
  Dim TempDispDate As Integer
  Dim TempVENDOR$
  Dim TempSERIALNO$
  Dim TempITEMMFG$
  Dim TempCONTACT$
  Dim TempITEMLOC$
  Dim TempEOLDATE As Integer
  Dim TempVHCLMAKE$
  Dim TempVHCLMODL$
  Dim TempVHCLVIN$
  Dim TempVHCLTAG$
  Dim TempVHCLCOLR$
  Dim TempWARRXDAT As Integer
  Dim TempFundNum As Integer
  Dim TempDisposAmt As Double
  Dim TempLastDprRec As Long
  Dim TempLifeLeft As Integer
  Dim TempPONum$
  Dim TempCheckNum$
  Dim TempDsplFlag$
  Dim TempDsplMethod$
  Dim FirstTime As Boolean
  Dim EditFlag As Boolean
  Dim TempGRecNum As Integer
  

Private Sub cmdAssetList_Click()
  frmFAAssetCodeList.Show vbModal
End Sub

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal
End Sub

Public Sub cmdExit_Click()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  Dim TVHandle As Integer
  Dim TVRec As TempVHCLDataType
  Dim TVCnt As Integer
  Dim VChangeFlag As Boolean
  
  ItemChangeFlag = False
  If VhclTempDsplFlag = True Then GoTo RecNumIsZero 'this means
  'that the fixed asset data on the screen represents a disposed
  'fixed asset...checking changes is not necessary (fields are
  'disabled for disposed of items)
  
  ChangeFlag = False
  VChangeFlag = False
  
  If GRecNum = 0 Then  'user is exiting
    If MsgBox("Are you sure you want to proceed without saving any changes?", vbYesNo) = vbYes Then
  'without saving new record entries...also skips the change
  'check feature and if taglist is open then the number
  'double clicked will be brought up to this screen
      GoTo RecNumIsZero
    Else
      Close
      fptxtTagNumber.SetFocus
      Exit Sub
    End If
  End If
  
  OpenFAItemFile FAHandle
  Get FAHandle, GRecNum, FAItemRec
  Close FAHandle
  
  OpenTempVhclFile TVHandle
  TVCnt = LOF(TVHandle) / Len(TVRec)

  If TVCnt = 0 Then 'check to see if there is anything
  'valid in the temporary vehicle data file and if there
  'is nothing there then skip the next section
    Close TVHandle
    GoTo NOTVRecs
  End If
  
  Get TVHandle, 1, TVRec 'otherwise get the temporary
  'vehicle data and look for unsaved changes
  Close TVHandle
  
  If QPTrim$(FAItemRec.VHCLMAKE) <> QPTrim$(TVRec.VHCLMAKE) Then
    VChangeFlag = True
    FocusOn = 1
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.VHCLMODL) <> QPTrim$(TVRec.VHCLMODL) Then
    VChangeFlag = True
    FocusOn = 2
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.VHCLVIN) <> QPTrim$(TVRec.VHCLVIN) Then
    VChangeFlag = True
    FocusOn = 3
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.VHCLTAG) <> QPTrim$(TVRec.VHCLTAG) Then
    VChangeFlag = True
    FocusOn = 4
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.VHCLCOLR) <> QPTrim$(TVRec.VHCLCOLR) Then
    VChangeFlag = True
    FocusOn = 5
    GoTo ChangeFound
  End If

NOTVRecs:
  
  If QPTrim$(fpDateWrntyX.Text) <> "NOT SAVED" Then
    If FAItemRec.WARRXDAT <> Date2Num(fpDateWrntyX) Then
      ChangeFlag = True
      fpDateWrntyX.SetFocus
      GoTo ChangeFound
    End If
  End If
  

  If QPTrim$(FAItemRec.DEPYN) <> QPTrim$(fpcmbDepYN.Text) Then
    ChangeFlag = True
    fpcmbDepYN.SetFocus
    GoTo ChangeFound
  End If
  
  If Mid(fpcmbStatus.Text, 1, 1) <> QPTrim$(FAItemRec.ISTATUS) Then
    ChangeFlag = True
    fpcmbStatus.SetFocus
    GoTo ChangeFound
  End If
  
  If FAItemRec.AQURDATE <> Date2Num(fptxtAcquiredDate) Then
    ChangeFlag = True
    fptxtAcquiredDate.SetFocus
    GoTo ChangeFound
  End If
    
  If FAItemRec.ILIFE <> Val(fptxtAssetLife.Text) Then
    ChangeFlag = True
    fptxtAssetLife.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(FAItemRec.CONTACT) <> QPTrim$(fptxtContact) Then
    ChangeFlag = True
    fptxtContact.SetFocus
    GoTo ChangeFound
  End If
    
  If OldRound(FAItemRec.DEP2DATE) <> OldRound(fptxtDep2Date.DoubleValue) Then
    ChangeFlag = True
    fptxtDep2Date.SetFocus
    GoTo ChangeFound
  End If
    
  If FAItemRec.IDEPT <> fptxtDeptNum Then
    ChangeFlag = True
    fptxtDeptNum.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(FAItemRec.IDESC1) <> QPTrim$(fptxtDesc1) Then
    ChangeFlag = True
    fptxtDesc1.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.IDESC2) <> QPTrim$(fptxtDesc2) Then
    ChangeFlag = True
    fptxtDesc2.SetFocus
    GoTo ChangeFound
  End If
  
  If fptxtDisposalDate <> "NOT SAVED" Then
    If FAItemRec.DispDate <> Date2Num(fptxtDisposalDate) Then
      ChangeFlag = True
      fptxtDisposalDate.SetFocus
      GoTo ChangeFound
    End If
  End If
  
  If FAItemRec.EOLDATE <> Date2Num(fptxtEOLDate) Then
    ChangeFlag = True
    fptxtEOLDate.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.GLACCT) <> QPTrim$(fptxtGLNum) Then
    ChangeFlag = True
    fptxtGLNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ASSETCODE) <> QPTrim$(fptxtGroupCode) Then
    ChangeFlag = True
    fptxtGroupCode.SetFocus
    GoTo ChangeFound
  End If
  
  If FAItemRec.FundNum <> Val(QPTrim$(fptxtFundNum)) Then
    ChangeFlag = True
    fptxtFundNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ITEMLOC) <> QPTrim$(fptxtLocation) Then
    ChangeFlag = True
    fptxtLocation.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ITEMMFG) <> QPTrim$(fptxtMfg) Then
    ChangeFlag = True
    fptxtMfg.SetFocus
    GoTo ChangeFound
  End If
  
  If FAItemRec.ORGCOST <> fptxtOriginalCost Then
    ChangeFlag = True
    fptxtOriginalCost.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.SERIALNO) <> QPTrim$(fptxtSerialNum) Then
    ChangeFlag = True
    fptxtSerialNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ItemTag) <> QPTrim$(fptxtTagNumber) Then
    ChangeFlag = True
    fptxtTagNumber.SetFocus
    GoTo ChangeFound
  End If
  
  If FAItemRec.DisposAmt <> fptxtDispPrice Then
    ChangeFlag = True
    fptxtOriginalCost.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.VENDOR) <> QPTrim$(fptxtVendorNum) Then
    ChangeFlag = True
    fptxtVendorNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.CheckNum) <> QPTrim$(fptxtChkNum.Text) Then
    ChangeFlag = True
    fptxtChkNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.PONum) <> QPTrim$(fptxtPONum.Text) Then
    ChangeFlag = True
    fptxtPONum.SetFocus
  End If
  
ChangeFound:
  'ItemChangeFlag..This flag is read by the tag list form to
  'know what to do with the decision made by the user below
  '-----------------------------------------------------------
  'The tag list can be double clicked to change the data on this screen. However, if
  'a change has been made before a new tag selection has been made and the user didn't
  'save it then he might lose data he thought was saved. So the ChangeFound routine looks
  'to see if the tag list is open ("taglistopen.dat") and if it is then we know it came
  'from the double click sub on that form where the .dat file is created. If the user wants
  'to abandon the changed data then the routine goes ahead and pops the screen with the
  'new tag data. If the user wants to save the changed data then the routine checks for
  'any save traps and if an error is found then the routine discards the new tag data and
  'returns the user to the screen to correct the error. If the user wants to save the changes
  'and there are no errors then the routine saves the data and pops the screen with the
  'new tag data. If the user wants to review any change then the routine discards the new
  'tag data and returns the user to the screen. The .dat file is always deleted so the tag
  'list can be reopened from scratch.
  If VChangeFlag = True Then 'Vehicle change
    VChangeFlag = False
    ItemChangeFlag = True
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
      Exit Sub
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      If Exist("taglistopen.dat") Then
        Unload frmFATagList
        KillFile ("taglistopen.dat")
      End If
      frmFAVehicleSpecific.Show vbModal
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      If Exist("taglistopen.dat") Then 'if taglist was open then
      'the purpose was to switch to another item so we don't want
      'to exit from the edit screen but we do want everything reset
      'to the new data
        ItemChangeFlag = False
        KillFile ("taglistopen.dat")
        Unload frmFAVehicleSpecific
        KillFile ("FATMPVHC.DAT")
        Exit Sub
      End If
      Unload frmFAVehicleSpecific
      frmFAItemLookUp.Show
      DoEvents
      KillFile ("edititemopen.dat")
      KillFile ("FATMPVHC.DAT")
      Unload frmFAEditItem
      AddItemFlag = False
      Exit Sub
    Case Else:
    End Select
    
  ElseIf ChangeFlag = True Then
    ChangeFlag = False
    ItemChangeFlag = True
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
      Exit Sub 'don't exit
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      If Exist("taglistopen.dat") Then 'if the tag list is still open then
      'unload it and return to the screen with the focus on tag number (set above)
        Unload frmFATagList
        KillFile ("taglistopen.dat")
      End If
      Exit Sub 'go back to the screen
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      If Exist("taglistopen.dat") Then
        Unload frmFAVehicleSpecific
        KillFile ("FATMPVHC.DAT")
        ItemChangeFlag = False 'tell the tag list that it's OK to continue
        'with changing the data on this screen with the new tag number entered
        KillFile ("taglistopen.dat")
        Exit Sub
      End If
      GRecNum = 0
      KillFile ("FATMPVHC.DAT") 'a user might have opened this
      'file and not made any changes...if it is not deleted then
      'it will appear with the next item's data
      Unload frmFAVehicleSpecific
      AddItemFlag = False
      frmFAItemLookUp.Show
      DoEvents
      KillFile ("edititemopen.dat")
      Unload frmFAEditItem
      Exit Sub
    Case Else:
    'Do nothing because we don't know about any options except
    'save, review or abandon...used as a placeholder for adding
    'other options at a later date
    End Select
  End If
RecNumIsZero:
  If Exist("taglistopen.dat") Then 'without this the program would
  'exit out to the main menu when tag list was double clicked
    KillFile ("taglistopen.dat")
    Exit Sub
  End If
  
  Unload frmFAVehicleSpecific
  If AddItemFlag = False Then
    frmFAItemLookUp.Show
  Else
    AddItemFlag = False
    frmFAItemMaintMenu.Show
  End If
  Close
  DoEvents
  KillFile ("edititemopen.dat")
  KillFile ("FATMPVHC.DAT")
  Unload frmFAEditItem
End Sub

Private Sub cmdFundList_Click()
  frmFAFundList.Show vbModal
End Sub

Private Sub cmdSave_Click()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfRecs As Long
  Dim IdxFlag As Boolean
  Dim TVHandle As Integer
  Dim TVRec As TempVHCLDataType
  Dim TVCnt As Integer
  Dim TagChangeIndexFlag As Boolean
  
  If VhclTempDsplFlag = True Then GoTo NoAsset 'this item has been disposed
  TagChangeIndexFlag = False
  IdxFlag = False
  'check for duplicate tag numbers only if the current tag entered
  'has been changed from the original number
  
  If QPTrim$(fptxtTagNumber.Text) <> FirstTagNum Then
    CheckForValidTAGNum
    If BadTagNum = True Then
      fptxtTagNumber.SetFocus
      BadTagNum = False
      Exit Sub
    Else
      TagChangeIndexFlag = True 'we're changing a tag number so it will need indexing
    End If
  End If
  
  CheckForValidAssetCodeNum
  'when taglistopen.dat exists then we know the user has accessed it
  'to get another tag number
  If BadAssetCodeNum = True Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    BadAssetCodeNum = False
    Exit Sub
  End If
  
  If Check4ValidDept = False Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    Exit Sub
  End If
  
  If Check4ValidFund = False Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    Exit Sub
  End If
  
  OpenFAItemFile FAHandle
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  
  If GRecNum = 0 Then 'new item here
    IdxFlag = True 'tells the program that tags need to be
    'reindexed to include this new item
    GRecNum = NumOfRecs + 1 'set record number where this data will
    'be saved
  Else
    Get FAHandle, GRecNum, FAItemRec 'otherwise pull up data for this asset
  End If
  
  If Len(QPTrim$(fpcmbDepYN.Text)) = 0 Then 'user forgot to fill in
  'the DepYN field
    If Exist("taglistopen.dat") Then 'OK...so if the tag list has been
    'accessed and is still open then the temp .dat file for that form should be deleted.
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no Y or N saved for depreciation. Please enter a Y or N for depreciation."
    Close FAHandle 'go back to the screen so the user can enter a value for DepYN
    If IdxFlag = True Then GRecNum = 0 'if this is a new record then reset the global to zero
    fpcmbDepYN.SetFocus
    Exit Sub
  Else
    FAItemRec.DEPYN = QPTrim$(fpcmbDepYN.Text)
  End If
  
  If Len(QPTrim$(fpcmbStatus.Text)) = 0 Then 'nothing saved for this
  'required field so close temp files and go back to screen
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no date saved for this item's status. Please enter a status value."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fpcmbStatus.SetFocus
    Exit Sub
  ElseIf fpcmbStatus.Text = "Active" Then 'field OK so save value
    FAItemRec.ISTATUS = "A"
  Else
    FAItemRec.ISTATUS = "I"
  End If
  
  If Len(QPTrim(fptxtAcquiredDate)) = 0 Then 'nothing saved for this
  'required field so close temp files down and go back to screen
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no date saved for this item's acquired date. Please enter an acquired date."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtAcquiredDate.SetFocus
    Exit Sub
  Else 'else save this value
    FAItemRec.AQURDATE = Date2Num(fptxtAcquiredDate)
  End If
  
  If QPTrim$(fpDateWrntyX.Text) <> "NOT SAVED" Then
    'user entered a warranty date but it comes before when the
    'asset was purchased
    If Date2Num(fpDateWrntyX) < Date2Num(fptxtAcquiredDate) Then
      If MsgBox("The warranty date entered comes before the acquire date. Do you wish to continue anyway?", vbYesNo) = vbNo Then
        If Exist("taglistopen.dat") Then
          Unload frmFATagList
          KillFile ("taglistopen.dat")
        End If
        Close
        fpDateWrntyX.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  If QPTrim$(fpDateWrntyX.Text) = "NOT SAVED" Then
    FAItemRec.WARRXDAT = 0
  Else
    FAItemRec.WARRXDAT = Date2Num(fpDateWrntyX)
  End If
  
  If fptxtAssetLife = 0 Then 'nothing saved for this required field
  'so close temp files and go back to screen
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's asset life. Please enter a value for asset life."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtAssetLife.SetFocus
    Exit Sub
  Else
    FAItemRec.ILIFE = Val(fptxtAssetLife) 'save this valid data
  End If
  
  FAItemRec.LifeLeft = Val(fptxtLeft.Text) 'can be edited but
  'is also figured automatically
  
  FAItemRec.CONTACT = QPTrim$(fptxtContact) 'not required
  
  If fptxtCurrDepDate = "NOT SAVED" Then
    FAItemRec.CDEPDATE = -11001 'value represents an invalid date...
    'program validates any date over -11000 (and under the year 2100)
  Else
    FAItemRec.CDEPDATE = Date2Num(fptxtCurrDepDate) 'save valid date
  End If
  
  FAItemRec.CURRVAL = fptxtCurrVal 'locked and automatically figured
  FAItemRec.DEP2DATE = fptxtDep2Date 'locked and automatically figured

  If Len(QPTrim(fptxtDeptNum)) = 0 Then 'required field with no valid value
    If Exist("taglistopen.dat") Then 'close down temp files and return to screen
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's department number. Please enter a value for department number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtDeptNum.SetFocus
    Exit Sub
  Else
    FAItemRec.IDEPT = Val(fptxtDeptNum) 'valid value so save it
  End If

  'no description entered for this item so close down temp files
  'and return to screen for correction
  If QPTrim$(fptxtDesc1) = "" And QPTrim$(fptxtDesc2) = "" Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "No description has been entered for this item. Please enter a description."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtDesc1.SetFocus
    Exit Sub
  Else 'otherwise save descriptions as entered
    FAItemRec.IDESC1 = QPTrim$(fptxtDesc1)
    FAItemRec.IDESC2 = QPTrim$(fptxtDesc2)
  End If
  
  'disposal data handling
  If fptxtDisposalDate = "NOT SAVED" Then
    FAItemRec.DsplFlag = 0
    FAItemRec.DispDate = 0
    GoTo NotDisposed
  ElseIf FAItemRec.DsplFlag = 1 Then
    GoTo NotDisposed
  Else
    FAItemRec.DispDate = Date2Num(fptxtDisposalDate) 'only if we allow disposal data to be changed here and not in the disposal routine
    FAItemRec.DsplFlag = 2
  End If
NotDisposed:
  FAItemRec.EOLDATE = Date2Num(fptxtEOLDate) 'a date will always be in here
  
  If Len(QPTrim(fptxtFundNum)) = 0 Then 'fund number is a required field
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's Fund number. Please enter a value for Fund number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtFundNum.SetFocus
    Exit Sub
  Else
    FAItemRec.FundNum = Val(QPTrim(fptxtFundNum)) 'save this valid fund number...
    'we know it's a valid fund number because it was checked earlier
  End If
  
  FAItemRec.GLACCT = QPTrim$(fptxtGLNum) 'not required
  
  If Len(QPTrim(fptxtGroupCode)) = 0 Then 'asset group code entry isn't valid
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's group code number. Please enter a value for group code number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtGroupCode.SetFocus
    Exit Sub
  Else
    FAItemRec.ASSETCODE = QPTrim$(fptxtGroupCode) 'value is good, it's been
    'checked in Check4ValidAssetCode...save it
  End If
  
  FAItemRec.ITEMLOC = QPTrim$(fptxtLocation) 'not required
  
  FAItemRec.ITEMMFG = QPTrim$(fptxtMfg) 'not required
  
  If fptxtOriginalCost = 0 Then 'purchase price value is not valid
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's purchase price. Please enter a value for purchase price."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtOriginalCost.SetFocus
    Exit Sub
  Else
    FAItemRec.ORGCOST = fptxtOriginalCost 'purchase price is valid so save it
  End If
  
  FAItemRec.SERIALNO = QPTrim$(fptxtSerialNum) 'not required
  
  If Len(QPTrim(fptxtTagNumber)) = 0 Then 'essential tag number not valid
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's tag number. Please enter a value for tag number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    fptxtTagNumber.SetFocus
    Exit Sub
  Else
    FAItemRec.ItemTag = QPTrim$(fptxtTagNumber) 'checked in Check4ValidTagNum and
    'is valid so save
  End If
  
  FAItemRec.VENDOR = QPTrim$(fptxtVendorNum) 'not required
  
  FAItemRec.CheckNum = QPTrim$(fptxtChkNum.Text) 'not required
  FAItemRec.PONum = QPTrim$(fptxtPONum.Text) 'not required
  
  OpenTempVhclFile TVHandle
  TVCnt = LOF(TVHandle) / Len(TVRec)
  If TVCnt = 0 And IdxFlag = True Then 'opened a new item but did not use Vehicle
    FAItemRec.VHCLMAKE = ""
    FAItemRec.VHCLMODL = ""
    FAItemRec.VHCLVIN = ""
    FAItemRec.VHCLTAG = ""
    FAItemRec.VHCLCOLR = ""
    Close TVHandle
  ElseIf TVCnt = 0 And IdxFlag = False Then 'opened an existing file and did not use Vehicle data
    Close TVHandle
  Else
    Get TVHandle, 1, TVRec
    Close TVHandle
    FAItemRec.VHCLMAKE = QPTrim$(TVRec.VHCLMAKE)
    FAItemRec.VHCLMODL = QPTrim$(TVRec.VHCLMODL)
    FAItemRec.VHCLVIN = QPTrim$(TVRec.VHCLVIN)
    FAItemRec.VHCLTAG = QPTrim$(TVRec.VHCLTAG)
    FAItemRec.VHCLCOLR = QPTrim$(TVRec.VHCLCOLR)
  End If
  
  If IdxFlag = True Then 'save values as empty just to hold space
    FAItemRec.Fill1 = ""
    FAItemRec.LastDprRec = 0
    FAItemRec.DsplMethod = ""
  End If
  
  KillFile ("FATMPVHC.DAT") 'this file is used only to assign values to FAItemRec
  
  Put FAHandle, GRecNum, FAItemRec 'save it to disk
  Close FAHandle
  
  If IdxFlag = True Or TagChangeIndexFlag = True Then 'this is a new asset so work it into
  'the tag index
    Call CreateTagIdx
    MainLog ("Item number " + QPTrim$(FAItemRec.ItemTag) + " data saved in frmFAEditItem.")
    IdxFlag = False
    TagChangeIndexFlag = False
  Else
    Call LogSaves 'records any changes made in this save with existing items
  End If
  
  MsgBox "Item data for " + QPTrim$(FAItemRec.ItemTag) + " has been saved."
  'If this save request was initiated by a double click from tag list then return
  'control back to that form

NoAsset:
  If Exist("taglistopen.dat") Then 'If a user has made a change and then
  'double clicked the tag list but did not save his change...he was alerted and
  'decide to save the change...this if statement sends the new tag data
  '(just double clicked) to the screen instead of exiting to the main menu
    KillFile ("taglistopen.dat")
    ItemChangeFlag = False
    fptxtTagNumber.SetFocus
    Exit Sub
  End If
  
  If AddItemFlag = True Then 'entering a list of several items is tedious
  'if after each save the program returns to the menu so this feature allows
  'the user to speed up the entry process
    If MsgBox("Do you wish to add another new item?", vbYesNo) = vbYes Then
      GRecNum = 0
      Unload frmFAVehicleSpecific
      fptxtTagNumber.SetFocus
      Call LoadMe
      Exit Sub
    Else
      Unload frmFAVehicleSpecific
      frmFAItemMaintMenu.Show
      DoEvents
      KillFile ("edititemopen.dat")
      Unload frmFAEditItem
      Exit Sub
    End If
  Else 'just editing an existing item...sends user back to menu upon
  'completion
    Unload frmFAVehicleSpecific
    GRecNum = 0
    frmFAItemLookUp.Show
    DoEvents
    KillFile ("edititemopen.dat")
    Unload frmFAEditItem
  End If
End Sub

Private Sub cmdTagList_Click()
  frmFATagList.Show vbModal
End Sub

Private Sub cmdVehicle_Click()
  frmFAVehicleSpecific.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FirstTime = True
  EditFlag = False
  TempTagNumber = ""
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%V"
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%F"
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%T"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%L"
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
      Unload frmFAVehicleSpecific
      KillFile ("edititemopen.dat")
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAEditItem.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Public Sub LoadMe()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim Today As String * 10
  Dim One As Integer
  Dim FileHandle As Integer
  
  VhclTempDsplFlag = False 'global telling the vehicle specific
  'form if it loads with fields enabled or not...at this point
  'they would be enabled because we're not sure if this is an
  'item disposed of or not
  lblDspl.Visible = False
  One = 1
  FileHandle = FreeFile
  Open "edititemopen.dat" For Output As FileHandle Len = 2
  
  Print #FileHandle, One
  Close FileHandle
  
  Date$ = FormatDateTime(Date, vbShortDate)
  
  Today = Date$
  
  fpcmbDepYN.Clear
  fpcmbDepYN.AddItem "Y"
  fpcmbDepYN.AddItem "N"
  
  fpcmbStatus.Clear
  fpcmbStatus.AddItem "Active"
  fpcmbStatus.AddItem "Inactive"
  
  If GRecNum = 0 Then 'load procedure for adding a new item
'    With AddItemFlag = True setting ...this global alerts the program
    'to use the "add another item" feature allowing the user to
    'speed up the entry process of a list of items to ass...keeps the
    'program from exiting out to the menu after each save
    fpcmbDepYN.Text = "N" 'defaults to this
    fpcmbStatus.Text = "Active" 'defaults to this
    fptxtAcquiredDate = Today 'defaults to this
    AcqDate = Date2Num(fptxtAcquiredDate)
    fptxtAssetLife = "0" 'defaults to this
    AssLife = 0 'defaults to this
    fptxtContact = "" 'defaults to this
    fptxtCurrDepDate = "NOT SAVED" 'defaults to this
    fptxtCurrVal = "0.00" 'defaults to this
    fptxtDep2Date = "0.00" 'defaults to this
    fptxtDeptNum = ""
    fptxtDesc1 = ""
    fptxtDesc2 = ""
    fptxtDisposalDate = "12/31/1979" 'means nothing is saved
    fptxtEOLDate = Today 'defaults to this
    fptxtGLNum = ""
    fptxtGroupCode = ""
    fptxtLocation = ""
    fptxtMfg = ""
    fptxtOriginalCost = "0.00" 'defaults to this
    fptxtSerialNum = "" 'etc
    fptxtFundNum = ""
    fptxtTagNumber = "" 'etc
    FirstTagNum = ""
    fptxtVendorNum = ""
    fptxtDispPrice = 0
    fptxtLeft.Text = 0
    fpDateWrntyX = "NOT SAVED" 'defaults to this
    fptxtPONum.Text = ""
    fptxtChkNum.Text = ""
  Else 'load procedure for an existing item
    OpenFAItemFile FAHandle
    Get FAHandle, GRecNum, FAItemRec
    If FAItemRec.DsplFlag > 0 Then 'this item is either disposed of or in the process
    'of being disposed of
      lblDspl.Visible = True 'show label telling of the disposal date
      If FAItemRec.DsplFlag = 2 Then 'if it's disposed of
        lblDspl.Caption = "This item was disposed of on " + MakeRegDate(FAItemRec.DispDate) + "."
      ElseIf FAItemRec.DsplFlag = 1 Then 'or it's in the disposal process
        lblDspl.Caption = "This item is scheduled for disposal on " + MakeRegDate(FAItemRec.DispDate) + "."
      End If
      VhclTempDsplFlag = True 'if vehicle specific form is accessed
      'then all fields will be disabled when that form reads this value
      
      'since this item is disposed of then disable the following
      cmdAssetList.Enabled = False
      cmdDept.Enabled = False
      cmdSave.Enabled = False
      cmdFundList.Enabled = False
'      cmdTagList.Enabled = False
      fpcmbDepYN.Enabled = False
      fpcmbStatus.Enabled = False
      fptxtAcquiredDate.Enabled = False
      fptxtAssetLife.Enabled = False
      fptxtContact.Enabled = False
      fptxtDep2Date.Enabled = False
      fptxtDeptNum.Enabled = False
      fptxtDesc1.Enabled = False
      fptxtDesc2.Enabled = False
      fpDateWrntyX.Enabled = False
      fptxtPONum.Enabled = False
      fptxtChkNum.Enabled = False
      fptxtGLNum.Enabled = False
      fptxtFundNum.Enabled = False
      fptxtGroupCode.Enabled = False
      fptxtOriginalCost.Enabled = False
      fptxtSerialNum.Enabled = False
      fptxtTagNumber.Enabled = False
      fptxtLocation.Enabled = False
      fptxtMfg.Enabled = False
      fptxtVendorNum.Enabled = False
      fptxtLeft.Enabled = False
    Else 'this item is not disabled so activate the following
      cmdAssetList.Enabled = True
      cmdDept.Enabled = True
      cmdSave.Enabled = True
      cmdFundList.Enabled = True
      lblDspl.Visible = False
      fpcmbDepYN.Enabled = True
      fpcmbStatus.Enabled = True
      fptxtAcquiredDate.Enabled = True
      fptxtAssetLife.Enabled = True
      fptxtContact.Enabled = True
      fptxtDep2Date.Enabled = True
      fptxtDeptNum.Enabled = True
      fptxtDesc1.Enabled = True
      fptxtDesc2.Enabled = True
      fpDateWrntyX.Enabled = True
      fptxtPONum.Enabled = True
      fptxtChkNum.Enabled = True
      fptxtGLNum.Enabled = True
      fptxtFundNum.Enabled = True
      fptxtGroupCode.Enabled = True
      fptxtOriginalCost.Enabled = True
      fptxtSerialNum.Enabled = True
      fptxtTagNumber.Enabled = True
      fptxtLocation.Enabled = True
      fptxtMfg.Enabled = True
      fptxtVendorNum.Enabled = True
      fptxtLeft.Enabled = True
    End If
    'now populate fields regardless if they are enabled or disabled
    '...globals are used when save routine occurs to record any changes
    'to main log
    FirstTagNum = QPTrim$(FAItemRec.ItemTag)
    TempItemTag$ = QPTrim$(FAItemRec.ItemTag) 'global
    fpcmbDepYN.Text = FAItemRec.DEPYN
    TempDEPYN$ = FAItemRec.DEPYN 'global
    If QPTrim$(FAItemRec.ISTATUS) = "A" Then
      fpcmbStatus.Text = "Active"
    Else
      fpcmbStatus.Text = "Inactive"
    End If
    TempISTATUS$ = QPTrim$(FAItemRec.ISTATUS) 'global
    fptxtAcquiredDate = MakeRegDate(FAItemRec.AQURDATE)
    AcqDate = FAItemRec.AQURDATE
    TempAQURDATE = FAItemRec.AQURDATE 'global
    If FAItemRec.ILIFE < 0 Then
      fptxtAssetLife = "0"
      AssLife = 0
    Else
      fptxtAssetLife = FAItemRec.ILIFE
      AssLife = FAItemRec.ILIFE
    End If
    TempILIFE = FAItemRec.ILIFE 'global
    fptxtContact = QPTrim$(FAItemRec.CONTACT)
    TempCONTACT$ = QPTrim$(FAItemRec.CONTACT) 'global
    If FAItemRec.CDEPDATE < -11000 Then 'roughly 1950
      FAItemRec.CDEPDATE = 0
      fptxtCurrDepDate = "NOT SAVED"
    Else
      fptxtCurrDepDate = MakeRegDate(FAItemRec.CDEPDATE)
    End If
    TempCDEPDATE = FAItemRec.CDEPDATE 'global
    fptxtCurrVal = FAItemRec.CURRVAL '
    TempCURRVAL = FAItemRec.CURRVAL 'global
    fptxtDep2Date = FAItemRec.DEP2DATE
    TempDEP2DATE = FAItemRec.DEP2DATE 'global
    fptxtDeptNum = FAItemRec.IDEPT
    TempIDEPT = FAItemRec.IDEPT 'global
    fptxtDesc1 = QPTrim$(FAItemRec.IDESC1)
    TempIDESC1$ = QPTrim$(FAItemRec.IDESC1) ' global
    fptxtDesc2 = QPTrim$(FAItemRec.IDESC2)
    TempIDESC2$ = QPTrim$(FAItemRec.IDESC2) 'global
    fptxtDisposalDate = MakeRegDate(FAItemRec.DispDate)
    TempDispDate = FAItemRec.DispDate 'global
    If CheckValDate(fptxtDisposalDate) = False Then
      fptxtDisposalDate.Text = "NOT SAVED"
    ElseIf FAItemRec.DispDate = 0 Then
      fptxtDisposalDate.Text = "NOT SAVED"
    End If
    
    fpDateWrntyX = MakeRegDate(FAItemRec.WARRXDAT)
    
    If CheckValDate(fpDateWrntyX.Text) = False Then
      fpDateWrntyX.Text = "NOT SAVED"
    ElseIf FAItemRec.WARRXDAT = 0 Then
      fpDateWrntyX.Text = "NOT SAVED"
    End If
    TempWARRXDAT = FAItemRec.WARRXDAT 'global
    fptxtPONum.Text = QPTrim$(FAItemRec.PONum)
    TempPONum$ = QPTrim$(FAItemRec.PONum) 'global
    fptxtChkNum.Text = QPTrim$(FAItemRec.CheckNum)
    TempCheckNum$ = QPTrim$(FAItemRec.CheckNum) 'global
    fptxtEOLDate = MakeRegDate(FAItemRec.EOLDATE)
    TempEOLDATE = FAItemRec.EOLDATE 'global
    fptxtGLNum = QPTrim$(FAItemRec.GLACCT)
    TempGLACCT$ = QPTrim$(FAItemRec.GLACCT) 'global
    fptxtFundNum = FAItemRec.FundNum
    TempFundNum = FAItemRec.FundNum 'global
    fptxtGroupCode = QPTrim$(FAItemRec.ASSETCODE)
    TempASSETCode$ = QPTrim$(FAItemRec.ASSETCODE) 'global
    fptxtLocation = QPTrim$(FAItemRec.ITEMLOC)
    TempITEMLOC$ = QPTrim$(FAItemRec.ITEMLOC) 'global
    fptxtMfg = QPTrim$(FAItemRec.ITEMMFG)
    TempITEMMFG$ = QPTrim$(FAItemRec.ITEMMFG) 'global
    fptxtOriginalCost = FAItemRec.ORGCOST
    TempORGCOST = FAItemRec.ORGCOST 'global
    fptxtSerialNum = QPTrim$(FAItemRec.SERIALNO)
    TempSERIALNO$ = QPTrim$(FAItemRec.SERIALNO) 'global
    fptxtTagNumber = QPTrim$(FAItemRec.ItemTag)
    fptxtVendorNum = QPTrim$(FAItemRec.VENDOR)
    TempVENDOR$ = QPTrim$(FAItemRec.VENDOR) 'global
    fptxtDispPrice = FAItemRec.DisposAmt
    TempDisposAmt = FAItemRec.DisposAmt 'global
    fptxtLeft.Text = FAItemRec.LifeLeft
    TempLifeLeft = FAItemRec.LifeLeft 'global
    AssLifeLeft = FAItemRec.LifeLeft
    Close FAHandle
  End If
  TempVHCLMAKE$ = QPTrim$(FAItemRec.VHCLMAKE) 'global
  TempVHCLMODL$ = QPTrim$(FAItemRec.VHCLMODL) 'global
  TempVHCLVIN$ = QPTrim$(FAItemRec.VHCLVIN) 'global
  TempVHCLTAG$ = QPTrim$(FAItemRec.VHCLTAG) 'global
  TempVHCLCOLR$ = QPTrim$(FAItemRec.VHCLCOLR) 'global
  If PWcnt = 0 Then 'this is a special case if sosoft needs
  'to access any field because of some kind of unexpected
  'problem...entering fixed assets with the sosoft code allows this
    fpcmbDepYN.Enabled = True
    fpcmbStatus.Enabled = True
    fptxtAcquiredDate.Enabled = True
    fptxtAssetLife.Enabled = True
    fptxtContact.Enabled = True
    fptxtDep2Date.Enabled = True
    fptxtDeptNum.Enabled = True
    fptxtDesc1.Enabled = True
    fptxtDesc2.Enabled = True
    fpDateWrntyX.Enabled = True
    fptxtPONum.Enabled = True
    fptxtChkNum.Enabled = True
    fptxtGLNum.Enabled = True
    fptxtFundNum.Enabled = True
    fptxtGroupCode.Enabled = True
    fptxtOriginalCost.Enabled = True
    fptxtSerialNum.Enabled = True
    fptxtTagNumber.Enabled = True
    fptxtLocation.Enabled = True
    fptxtMfg.Enabled = True
    fptxtLeft.Enabled = True
    fptxtLeft.ControlType = ControlTypeNormal
    fptxtCurrDepDate.ControlType = ControlTypeNormal
    fptxtCurrVal.ControlType = ControlTypeNormal
    fptxtDep2Date.ControlType = ControlTypeNormal
    fptxtDisposalDate.ControlType = ControlTypeNormal
    fptxtDispPrice.ControlType = ControlTypeNormal
    fptxtEOLDate.ControlType = ControlTypeNormal
  End If
End Sub
Private Sub fpcmbDepYN_KeyDown(KeyCode As Integer, Shift As Integer)
  'This routine is designed to allow the user to scroll through the
  'form without inadvertently changing data in this combo box
  If KeyCode = vbKeySpace Then
    fpcmbDepYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDepYN.ListIndex = -1
  End If
  If fpcmbDepYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  'This routine is designed to allow the user to scroll through the
  'form without inadvertently changing data in this combo box
  If KeyCode = vbKeySpace Then
    fpcmbStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStatus.ListIndex = -1
  End If
  If fpcmbStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtAcquiredDate_LostFocus()
  Dim AcqYear As Integer
  Dim NewEOL As Integer
  'this routine automatically updates the EOL field if
  'the user changes the acquire date
  If Date2Num(fptxtAcquiredDate) <> AcqDate Then 'program sees that this field has changed
  'from the currently saved global acquire date
    AcqYear = CInt(Mid(fptxtAcquiredDate, 7, 4)) 'find the new acquire year
    NewEOL = AcqYear + CInt(fptxtAssetLife) 'now determine the new EOL Year
    fptxtEOLDate = Mid(fptxtAcquiredDate, 1, 6) + CStr(NewEOL) 'now assign entire date to EOL
    AcqDate = Date2Num(fptxtAcquiredDate) 'reassign global
  End If

End Sub

Private Sub fptxtAssetLife_Change()
  'goes ahead and figures the life left based on the
  'assigned life for this new asset
  If GRecNum = 0 Then
    fptxtLeft.Text = fptxtAssetLife.Text
  End If
  
End Sub

'Private Sub CheckForValidWHNum()
'   Dim JGLIdxRec(1) As JGLAcctIdxType
'   Dim GLIdxNum$
'   Dim GLDHandle As Integer
'   Dim GLIdxRecLen As Integer
'   Dim GLDescRecLen As Integer
'   Dim TotalAccts As Integer
'   Dim x As Integer
'   Dim GLIDATDesc$
'   Dim GLDesc(1) As GLAcctRecType
'   Dim GLIdxHandle As Integer
'   Dim DoWhatFlag As BadGLNUMOption
'   Dim n As Integer
'   Dim FundLength As Integer
'   Dim AcctLength As Integer
'   Dim DetLength As Integer
'   Dim Nextx As Integer
'   Dim Y As Integer
'   Dim ThisText$
'
'   On Error GoTo ERRORSTUFF
'   Call GetAcctStruct(FundLength, AcctLength, DetLength)
'   If FundLength = 0 And AcctLength = 0 And DetLength = 0 Then
'     Exit Sub
'   End If
'
'   BadGLNum = False
'
'   If Exist("GLACCT.IDX") Then
'     GLIdxNum$ = "GLACCT.IDX"
'   Else
'     MsgBox "No G/L account number validation possible...GLACCT.IDX could not be found."
'     Exit Sub
'   End If
'
'   If Exist("GLACCT.DAT") Then
'     GLIDATDesc$ = "GLACCT.DAT"
'   Else
'     MsgBox "No G/L account number validation possible...GLACCT.DAT could not be found."
'     Exit Sub
'   End If
'
'   GLIdxRecLen = Len(JGLIdxRec(1))
'   GLDescRecLen = Len(GLDesc(1))
'   TotalAccts = FileSize(GLIDATDesc$) \ GLDescRecLen
'
'   If TotalAccts = 0 Then Exit Sub
'   ReDim DescBuff(1 To TotalAccts)
'   GLIdxHandle = FreeFile
'   Open GLIdxNum$ For Random As GLIdxHandle Len = GLIdxRecLen
'   For x = 1 To TotalAccts
'     Get GLIdxHandle, x, JGLIdxRec(1)
'     DescBuff(x) = JGLIdxRec(1).RecNo
'   Next x
'   Close GLIdxHandle
'   GLDHandle = FreeFile
'   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
'
'   'go thru each number one at a time and compare against all gl nums
'   ThisText$ = QPTrim$(ReplaceString(fptxtGLNum, "-", ""))
'   If ThisText$ = "" Then GoTo ZeroText
'   For x = 1 To TotalAccts
'   If DescBuff(x) = 0 Then GoTo DescBuffIsZero
'     Get GLDHandle, DescBuff(x), GLDesc(1)
'        If ThisText = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
'          Exit For
'        End If
'DescBuffIsZero:
'     If x = TotalAccts Then
'       DoWhatFlag = PromptBadGLNum(Me)
'       Select Case DoWhatFlag
'       Case BadGLNUMOption.badglReturn
'         Close
'         fptxtGLNum.SetFocus
'         BadGLNum = True
'         Exit Sub
'       Case BadGLNUMOption.badglSave
'         Close
'         Exit Sub
'       Case BadGLNUMOption.badglExit
'         frmFAItemLookUp.Show
'         DoEvents
'         Unload frmFAEditItem
'         BadGLNum = True
'         Close
'         Exit Sub
'       Case Else:
'          'Do nothing because we don't know about any options except
'          'save, review or abandon...used as a placeholder for adding
'          'other options at a later date
'       End Select
'       Close GLDHandle
'       Exit Sub
'     End If
'  Next x
'ZeroText:
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmSPRTThisEmp", "CheckForValidWHNum", Erl)
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
'End Sub

Private Sub fptxtAssetLife_LostFocus()
  Dim AcqYear As Integer
  Dim NewEOL As Integer
  Dim LifeDif As Integer
  
  'asset life affects when the assets EOL will be so this routine
  'determines EOL and asset life left
  If Val(fptxtAssetLife) <= 0 Then 'this may change but for now all
  'assets must have a life of at least 1 year
    MsgBox "Each fixed asset must have a life of at least 1 year."
    fptxtAssetLife = 1
  End If
  
  If GRecNum = 0 Then 'this is a new asset being added
    fptxtLeft.Text = fptxtAssetLife.Text 'life and life left are the
    'same for a new asset
    AcqYear = CInt(Mid(fptxtAcquiredDate, 7, 4)) 'assign global
    NewEOL = AcqYear + CInt(fptxtAssetLife) 'now figure EOL
    fptxtEOLDate = Mid(fptxtAcquiredDate, 1, 6) + CStr(NewEOL)
    If fptxtAssetLife <> AssLife Then 'update global if necessary
      AssLife = fptxtAssetLife
    End If
    Exit Sub
  End If
  
  If fptxtAssetLife = "" Then 'reassign global if field is blank
    fptxtAssetLife = AssLife
  End If
  
  If fptxtAssetLife <> AssLife Then 'change detected
    AcqYear = CInt(Mid(fptxtAcquiredDate, 7, 4)) 'start changing the EOL Date
    NewEOL = AcqYear + CInt(fptxtAssetLife) 'get new EOL Year
    fptxtEOLDate = Mid(fptxtAcquiredDate, 1, 6) + CStr(NewEOL) 'assign entire EOL Date
    If AssLife > CInt(fptxtAssetLife.Text) Then 'old asset life was more than the new one
      LifeDif = AssLife - CInt(fptxtAssetLife.Text) 'get difference between old and new
      fptxtLeft.Text = CInt(fptxtLeft.Text) - LifeDif 'new life left value is reduced to this value
      If Val(fptxtLeft.Text) < 0 Then fptxtLeft.Text = 0 'if the new life left ends up being
      'less than 0 then make it 0
    ElseIf AssLife < CInt(fptxtAssetLife.Text) Then 'old asset life is less than the new asset life
      LifeDif = CInt(fptxtAssetLife.Text) - AssLife 'get difference between new and old
      fptxtLeft.Text = CInt(fptxtLeft.Text) + LifeDif 'increase life left by the difference
      If Val(fptxtLeft.Text) < 0 Then fptxtLeft.Text = 0 'this should never happen
    End If
    AssLife = CInt(fptxtAssetLife) 'reassign global
    AssLifeLeft = CInt(fptxtLeft.Text)
  End If
  
End Sub

Private Sub fptxtDep2Date_LostFocus()
  Dim ORCost As Double
  Dim DepToDate As Double
  
  'this procedure updates the fixed asset's current value
  'if the user changes the depreciation amount to date manually
  ORCost = fptxtOriginalCost
  DepToDate = fptxtDep2Date
  fptxtCurrVal = ORCost - DepToDate

End Sub

Private Sub fptxtDisposalDate_Change()

  If fptxtDisposalDate.Text = "12/31/1979" Or QPTrim$(fptxtDisposalDate) = "" Then
    fptxtDisposalDate.Text = "NOT SAVED"
    Exit Sub
  End If
  
  If fptxtDisposalDate.Text <> "NOT SAVED" And fptxtDispPrice > 0 Then
    fptxtCurrVal = 0 'if a disposal price and date have been saved then this
    'asset is no longer owned and the current value has to be zero
  End If
  
End Sub

Private Sub fptxtDispPrice_Change()
  If fptxtDisposalDate.Text <> "NOT SAVED" And fptxtDispPrice > 0 Then
    fptxtCurrVal = 0 'if the disposal price is more than 0 and the date is valid then the item is assumed
    'to be disposed of and has no more value
  End If

End Sub

Private Sub CheckForValidAssetCodeNum()
   Dim CodeRec As FAAssetCodeRecType
   Dim ACHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim ThisText$
   
   'this routine is designed to make sure that the asset code entered
   'by the user is actually one of the codes saved
   On Error GoTo ERRORSTUFF
   
   ThisText$ = QPTrim$(fptxtGroupCode) 'user entered no value
   If ThisText$ = "" Then GoTo ZeroText 'exit sub...this is a required
   'field so if the user tries to save an empty filed the program
   'will force him to enter a valid number
   
   BadAssetCodeNum = False 'so far this asset code is OK
   
   If Not Exist("FACODES.DAT") Then 'nothing saved for system asset codes yet
     MsgBox "No Fixed Asset code number validation possible...FACODES.DAT could not be found."
     Exit Sub 'can't validate
   End If
   
   OpenFACodeNameFile ACHandle
   TotalAccts = LOF(ACHandle) \ Len(CodeRec)
   
   If TotalAccts = 0 Then
     MsgBox "No Fixed Asset code numbers on file." 'screens if the file
     'exists but contains no data
     Exit Sub
   End If
   
   'go thru each number one at a time and compare against all asset code nums
   For x = 1 To TotalAccts
     Get ACHandle, x, CodeRec
       If ThisText = QPTrim$(CodeRec.ASSETCODE) Then 'found a match...this number is OK
         Exit For 'no reason to continue matching
       End If
  Next x
  
  If x = TotalAccts + 1 Then 'been thru all depts and found nothing to match
  'what the user entered
    MsgBox "The asset code number entered is not valid. Check the asset code list for valid asset code numbers."
    BadAssetCodeNum = True
    fptxtGroupCode.SetFocus
  End If
  
  Close ACHandle
ZeroText:
   
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmSPRTThisEmp", "CheckForValidWHNum", Erl)
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
    Close
End Sub

Private Sub CheckForValidTAGNum()
   Dim TagRec As FAItemRecType
   Dim THandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim ThisText$
   
   On Error GoTo ERRORSTUFF
   
   ThisText$ = QPTrim$(fptxtTagNumber)
   If ThisText$ = "" Then Exit Sub
   
   BadTagNum = False
   
   OpenFAItemFile THandle
   TotalAccts = LOF(THandle) \ Len(TagRec)
   If TotalAccts = 0 Then
     Close
     Exit Sub
   End If
   'go thru each number one at a time and compare against all tag numbers
   For x = 1 To TotalAccts
     Get THandle, x, TagRec
       If ThisText = QPTrim$(TagRec.ItemTag) Then
         MsgBox "This Tag Number is already in use. Please select a new one."
         BadTagNum = True
         Exit For
       End If
   Next x
   Close THandle
    
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmSPRTThisEmp", "CheckForValidWHNum", Erl)
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
    Close
End Sub

Private Function Check4ValidDept() As Boolean
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim x As Integer
  Dim NumOfDepts As Integer
  Dim CompareThis As Integer
  
  Check4ValidDept = True 'so far the number is fine
  
  OpenFADeptCodeFile DHandle
  NumOfDepts = LOF(DHandle) \ Len(DeptRec)
  
  If NumOfDepts = 0 Then 'no depts saved
    Close
    If MsgBox("There are no department numbers saved. It is recommended that department numbers be saved before continuing. Do you wish to continue anyway?", vbYesNo) = vbNo Then
      Check4ValidDept = False 'user warned and he elected to not save this data
    Else 'user warned and he elected to save anyway and it was recorded in the log
      MainLog ("User warned that no department numbers have been saved. The user elected to save item data anyway for tag number in frmFAEditItem" + fptxtTagNumber.Text + ". ")
    End If
    Exit Function 'no departments saved so no need to try to match anything
  End If
  
  CompareThis = Val(fptxtDeptNum.Text)
  For x = 1 To NumOfDepts
    Get DHandle, x, DeptRec 'start looking for a department match
    If CompareThis = Val(DeptRec.DeptNum) Then
      Exit For 'found it so we're finished
    End If
  Next x
  
  If x = NumOfDepts + 1 Then 'been thru all depts and found nothing to match
  'waht the user entered
    MsgBox "The department number entered is not valid. Check the department list for valid department numbers."
    Check4ValidDept = False
    fptxtDeptNum.SetFocus
  End If
End Function

Private Function Check4ValidFund() As Boolean
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim x As Integer
  Dim NumOfFunds As Integer
  Dim CompareThis As Integer
  
  Check4ValidFund = True
  
  OpenFAFundCodeFile FHandle
  NumOfFunds = LOF(FHandle) \ Len(FundRec)
  If NumOfFunds = 0 Then
    Close
    If MsgBox("There are no fund numbers available from which to choose. It is recommended that fund numbers be saved before continuing. Do you wish to continue anyway?", vbYesNo) = vbNo Then
      Check4ValidFund = False 'user warned and he elected to not save this data
    Else 'user warned and he elected to save anyway and it was recorded in the log
      MainLog ("User warned that no fund numbers have been saved. The user elected to save item data anyway for tag number in frmFAEditItem " + fptxtTagNumber.Text + ".")
    End If
    Exit Function 'no departments saved so no need to try to match anything
  End If
  
  CompareThis = Val(fptxtFundNum.Text)
  
  For x = 1 To NumOfFunds
    Get FHandle, x, FundRec
    If CompareThis = Val(FundRec.FundNum) Then
      Exit For
    End If
  Next x
  
  If x = NumOfFunds + 1 Then
    MsgBox "The fund number entered is not valid. Check the fund list for valid fund numbers."
    Check4ValidFund = False
    fptxtFundNum.SetFocus
  End If
  
End Function

Private Sub fptxtLeft_Change()
'  If QPTrim$(fptxtLeft.Text) = "" Then
'    fptxtLeft.Text = AssLifeLeft
'  ElseIf CInt(fptxtLeft.Text) <> AssLifeLeft Then
'    If MsgBox("The asset life left has been edited and may not be accurate. If you are not sure of the accuracy of this change then select No. If you want to continue with this value anyway then select Yes.", vbYesNo) = vbNo Then
'      fptxtLeft.Text = CStr(AssLifeLeft)
'    End If
'  End If
End Sub

Private Sub fptxtLeft_LostFocus()
  'This routine tries to protect the integrity of the life left value
  If QPTrim$(fptxtLeft.Text) = "" Then
    fptxtLeft.Text = AssLifeLeft
  ElseIf CInt(fptxtLeft.Text) <> AssLifeLeft Then 'AssLifeLeft should be the correct value
  'figured by the program...if the user changes this number to an inaccurate number and
  'then changes it again to the old correct value then the program will alert him again as
  'if the current change may be wrong
    If MsgBox("The asset life left has been edited and may not be accurate. If you are NOT sure of the accuracy of this change then select No. If you want to continue with this value anyway then select Yes.", vbYesNo) = vbNo Then
      fptxtLeft.Text = CStr(AssLifeLeft)
      fptxtLeft.SetFocus
    Else 'record this warning
      MainLog ("The user changed the life left value for this asset. A warning was issued stating that the new asset life left (" + fptxtLeft.Text + " years) may not be accurate. The current asset life is " + CStr(AssLifeLeft) + " years. The user elected to save anyway in frmFAEditItem.")
    End If
  End If

End Sub

Private Sub fptxtOriginalCost_LostFocus()
  Dim ORCost As Double
  Dim DepToDate As Double
  
  'update the current value based on a change in the
  'original cost
  ORCost = fptxtOriginalCost
  DepToDate = fptxtDep2Date
  fptxtCurrVal = ORCost - DepToDate

End Sub

Private Sub LogSaves()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  
  'save to the log any kind of change saved
  OpenFAItemFile FAHandle
  Get FAHandle, GRecNum, FAItemRec
  Close FAHandle
  
  If QPTrim$(TempItemTag$) <> QPTrim$(FAItemRec.ItemTag) Then
    MainLog ("Item Tag Number, " + QPTrim$(TempItemTag$) + ", changed and saved as " + QPTrim$(FAItemRec.ItemTag) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempISTATUS$) <> QPTrim$(FAItemRec.ISTATUS) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item Status, " + QPTrim$(TempISTATUS$) + ", changed and saved as " + QPTrim$(FAItemRec.ISTATUS) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempDEPYN$) <> QPTrim$(FAItemRec.DEPYN) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Depreciate Y/N?, " + QPTrim$(TempDEPYN$) + ", changed and saved as " + QPTrim$(FAItemRec.DEPYN) + " in frmFAEditItem.")
  End If
  
  If TempAQURDATE <> FAItemRec.AQURDATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item acquire date, " + MakeRegDate(TempAQURDATE) + ", changed and saved as " + MakeRegDate(FAItemRec.AQURDATE) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempIDESC1$) <> QPTrim$(FAItemRec.IDESC1) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item description 1,  " + QPTrim$(TempIDESC1$) + ", changed and saved as " + QPTrim$(FAItemRec.IDESC1) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempIDESC2$) <> QPTrim$(FAItemRec.IDESC2) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item description 2, " + QPTrim$(TempIDESC2$) + ", changed and saved as " + QPTrim$(FAItemRec.IDESC2) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempGLACCT$) <> QPTrim$(FAItemRec.GLACCT) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item GL acct number, " + QPTrim$(TempGLACCT$) + ", changed and saved as " + QPTrim$(FAItemRec.GLACCT) + " in frmFAEditItem.")
  End If
  
  If TempIDEPT <> FAItemRec.IDEPT Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item department number, " + CStr(TempIDEPT) + ", changed and saved as " + CStr(FAItemRec.IDEPT) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempASSETCode$) <> QPTrim$(FAItemRec.ASSETCODE) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item asset code number, " + QPTrim$(TempASSETCode$) + ", changed and saved as " + QPTrim$(FAItemRec.ASSETCODE) + " in frmFAEditItem.")
  End If
  
  If TempILIFE <> FAItemRec.ILIFE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item life, " + CStr(TempILIFE) + ", changed and saved as " + CStr(FAItemRec.ILIFE) + " in frmFAEditItem.")
  End If
  
  If TempORGCOST <> FAItemRec.ORGCOST Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item purchase price, " + CStr(TempORGCOST) + ", changed and saved as " + CStr(FAItemRec.ORGCOST) + " in frmFAEditItem.")
  End If
  
  If TempDEP2DATE <> FAItemRec.DEP2DATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item depreciation to date,  " + CStr(TempDEP2DATE) + ", changed and saved as " + CStr(FAItemRec.DEP2DATE) + " in frmFAEditItem.")
  End If
  
  If TempCURRVAL <> FAItemRec.CURRVAL Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item current value, " + CStr(TempCURRVAL) + ", changed and saved as " + CStr(FAItemRec.CURRVAL) + " in frmFAEditItem.")
  End If
  
  If FAItemRec.CDEPDATE < -11000 Then FAItemRec.CDEPDATE = 0
  If TempCDEPDATE <> FAItemRec.CDEPDATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item last depreciation date, " + MakeRegDate(TempCDEPDATE) + ", changed and saved as " + MakeRegDate(FAItemRec.CDEPDATE) + " in frmFAEditItem.")
  End If
  
  If TempDispDate <> FAItemRec.DispDate Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item disposal date, " + MakeRegDate(TempDispDate) + ", changed and saved as " + MakeRegDate(FAItemRec.DispDate) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempVENDOR$) <> QPTrim$(FAItemRec.VENDOR) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vendor, " + QPTrim$(TempVENDOR$) + ", changed and saved as " + QPTrim$(FAItemRec.VENDOR) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempSERIALNO$) <> QPTrim$(FAItemRec.SERIALNO) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item serial number, " + QPTrim$(TempSERIALNO$) + ", changed and saved as " + QPTrim$(FAItemRec.SERIALNO) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempITEMMFG$) <> QPTrim$(FAItemRec.ITEMMFG) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item manufacturer, " + QPTrim$(TempITEMMFG$) + ", changed and saved as " + QPTrim$(FAItemRec.ITEMMFG) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempCONTACT$) <> QPTrim$(FAItemRec.CONTACT) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item contact, " + QPTrim$(TempCONTACT$) + ", changed and saved as " + QPTrim$(FAItemRec.CONTACT) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempITEMLOC$) <> QPTrim$(FAItemRec.ITEMLOC) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item location, " + QPTrim$(TempITEMLOC$) + ", changed and saved as " + QPTrim$(FAItemRec.ITEMLOC) + " in frmFAEditItem.")
  End If
  
  If TempEOLDATE <> FAItemRec.EOLDATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item end of life, " + MakeRegDate(TempEOLDATE) + ", changed and saved as " + MakeRegDate(FAItemRec.EOLDATE) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempVHCLMAKE$) <> QPTrim$(FAItemRec.VHCLMAKE) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle make, " + QPTrim$(TempVHCLMAKE$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLMAKE) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempVHCLMODL$) <> QPTrim$(FAItemRec.VHCLMODL) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle model, " + QPTrim$(TempVHCLMODL$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLMODL) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempVHCLVIN$) <> QPTrim$(FAItemRec.VHCLVIN) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle ID number, " + QPTrim$(TempVHCLVIN$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLVIN) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempVHCLTAG$) <> QPTrim$(FAItemRec.VHCLTAG) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle license tag number, " + QPTrim$(TempVHCLTAG$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLTAG) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempVHCLCOLR$) <> QPTrim$(FAItemRec.VHCLCOLR) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle color," + QPTrim$(TempVHCLCOLR$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLCOLR) + " in frmFAEditItem.")
  End If
  
  If TempWARRXDAT <> FAItemRec.WARRXDAT Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item warranty expiration date, " + MakeRegDate(TempWARRXDAT) + ", changed and saved as " + MakeRegDate(FAItemRec.WARRXDAT) + " in frmFAEditItem.")
  End If
  
  If TempFundNum <> FAItemRec.FundNum Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item fund number, " + CStr(TempFundNum) + ", changed and saved as " + CStr(FAItemRec.FundNum) + " in frmFAEditItem.")
  End If
  
  If TempDisposAmt <> FAItemRec.DisposAmt Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item disposal amount, " + CStr(TempDisposAmt) + ", changed and saved as " + CStr(FAItemRec.DisposAmt) + " in frmFAEditItem.")
  End If
  
  If TempLifeLeft <> FAItemRec.LifeLeft Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item life left, " + CStr(TempLifeLeft) + ", changed and saved as " + CStr(FAItemRec.LifeLeft) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempPONum$) <> QPTrim$(FAItemRec.PONum) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item purchase order number, " + QPTrim$(TempPONum$) + ", changed and saved as " + QPTrim$(FAItemRec.PONum) + " in frmFAEditItem.")
  End If
  
  If QPTrim$(TempCheckNum$) <> QPTrim$(FAItemRec.CheckNum) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item check number, " + QPTrim$(TempCheckNum$) + ", changed and saved as " + QPTrim$(FAItemRec.CheckNum) + " in frmFAEditItem.")
  End If

End Sub

