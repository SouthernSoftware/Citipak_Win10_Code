VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxPenaltySetUp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Penalty Setup"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   Icon            =   "frmTaxPenaltySetUp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpCombo fpcmbMethod 
      Height          =   372
      Left            =   4080
      TabIndex        =   2
      Tag             =   $"frmTaxPenaltySetUp.frx":08CA
      Top             =   1560
      Width           =   3540
      _Version        =   196608
      _ExtentX        =   6244
      _ExtentY        =   656
      Text            =   ""
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
      ColDesigner     =   "frmTaxPenaltySetUp.frx":0A32
   End
   Begin LpLib.fpCombo fpcmbCombine 
      Height          =   372
      Left            =   4080
      TabIndex        =   3
      Tag             =   $"frmTaxPenaltySetUp.frx":0E0D
      Top             =   2040
      Width           =   3060
      _Version        =   196608
      _ExtentX        =   5397
      _ExtentY        =   656
      Text            =   ""
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
      ColDesigner     =   "frmTaxPenaltySetUp.frx":0F75
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   4800
      Width           =   8055
      _Version        =   196609
      _ExtentX        =   14208
      _ExtentY        =   873
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      Caption         =   ""
      Picture         =   "frmTaxPenaltySetUp.frx":1350
      Begin VB.OptionButton OptPenOpt3 
         BackColor       =   &H008F8265&
         Caption         =   "Penalty Rev Opt 3"
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
         Left            =   5520
         TabIndex        =   14
         Top             =   120
         Width           =   2175
      End
      Begin VB.OptionButton OptPenOpt2 
         BackColor       =   &H008F8265&
         Caption         =   "Penalty Rev Opt 2"
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
         TabIndex        =   13
         Top             =   120
         Width           =   2175
      End
      Begin VB.OptionButton OptPenOpt1 
         BackColor       =   &H008F8265&
         Caption         =   "Penalty Rev Opt 1"
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
         TabIndex        =   12
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.CheckBox chkOptRev3 
      BackColor       =   &H008F8265&
      Caption         =   "Optional Revenue #3"
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
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CheckBox chkOptRev2 
      BackColor       =   &H008F8265&
      Caption         =   "Optional Revenue #2"
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
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CheckBox chkOptRev1 
      BackColor       =   &H008F8265&
      Caption         =   "Optional Revenue #1"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CheckBox chkLateList 
      BackColor       =   &H008F8265&
      Caption         =   "Late Listing"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox chkAdvertising 
      BackColor       =   &H008F8265&
      Caption         =   "Advertising/Collection"
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
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox chkInterest 
      BackColor       =   &H008F8265&
      Caption         =   "Interest"
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
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox chkRealEstate 
      BackColor       =   &H008F8265&
      Caption         =   "Real Estate"
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
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox chkPersProp 
      BackColor       =   &H008F8265&
      Caption         =   "Personal Property"
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
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin EditLib.fpDoubleSingle fpdblPct 
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegFormat       =   1
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
   Begin EditLib.fpCurrency fpCurrFlatRate 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
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
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   492
      Left            =   1128
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1932
      _Version        =   131072
      _ExtentX        =   3408
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
      ButtonDesigner  =   "frmTaxPenaltySetUp.frx":136C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   492
      Left            =   5688
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1932
      _Version        =   131072
      _ExtentX        =   3408
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
      ButtonDesigner  =   "frmTaxPenaltySetUp.frx":1548
   End
   Begin EditLib.fpText fptxtPenDesc 
      Height          =   390
      Left            =   3600
      TabIndex        =   15
      Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
      Top             =   5400
      Width           =   3735
      _Version        =   196608
      _ExtentX        =   6588
      _ExtentY        =   688
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   492
      Left            =   3408
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1932
      _Version        =   131072
      _ExtentX        =   3408
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
      ButtonDesigner  =   "frmTaxPenaltySetUp.frx":1723
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   5775
      Left            =   120
      Top             =   120
      Width           =   8535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Description:"
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
      Left            =   1200
      TabIndex        =   25
      Top             =   5400
      Width           =   2190
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Penalty Revenue Select:"
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
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Apply Penalty To:"
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
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Combination Type:"
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
      Left            =   1800
      TabIndex        =   22
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Application Method:"
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
      Left            =   720
      TabIndex        =   21
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Pct Rate:"
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
      Left            =   4680
      TabIndex        =   20
      Top             =   1140
      Width           =   2070
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Flat Rate:"
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
      Left            =   480
      TabIndex        =   19
      Top             =   1140
      Width           =   2070
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   510
      Left            =   2370
      Top             =   330
      Width           =   4050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Penalty Setup"
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
      Height          =   330
      Left            =   2625
      TabIndex        =   18
      Top             =   405
      Width           =   3510
   End
End
Attribute VB_Name = "frmTaxPenaltySetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ThisPenIdx As Integer
  Dim TempPenIdx   As Integer
  Dim TempPenDesc  As String
  Dim TempPenPct   As Double
  Dim TempPenFlat  As Double
  Dim TempUsePct   As String
  Dim TempUseFlat  As String
  Dim TempUseBoth  As String
  Dim TempUseHigh  As String
  Dim TempUseLow   As String
  Dim TempAppToRev1 As String
  Dim TempAppToRev2 As String
  Dim TempAppToRev3 As String
  Dim TempAppToRev4 As String
  Dim TempAppToRev5 As String
  Dim TempAppToRev6 As String
  Dim TempRev6Name  As String
  Dim TempAppToRev7 As String
  Dim TempRev7Name  As String
  Dim TempAppToRev8 As String
  Dim TempRev8Name  As String
  
Private Sub chkOptRev3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If OptPenOpt1.Enabled = True Then
      OptPenOpt1.SetFocus
    ElseIf OptPenOpt2.Enabled = True Then
      OptPenOpt2.SetFocus
    ElseIf OptPenOpt3.Enabled = True Then
      OptPenOpt3.SetFocus
    Else
      fptxtPenDesc.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    chkOptRev2.SetFocus
  End If
End Sub

Private Sub cmdClose_Click()
  If Check4Changes() = True Then
    Close
    Exit Sub
  End If
  
  Unload Me
  DoEvents
End Sub

Private Sub cmdDelete_Click()
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  Dim SetUpRec As TaxMasterType
  Dim MHandle As Integer
  
  If Not Exist(TaxPenHandling) Then
    frmTaxMsg.Label1.Caption = "There are no penalty setup records saved."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  frmTaxMsgWOpts.Label1.Caption = "Are you sure you wish to delete the current penalty setup records? To continue to delete then press F10. Otherwise, press ESC to abort the deletion procedure."
  frmTaxMsgWOpts.Label1.Top = 700
  frmTaxMsgWOpts.cmdCont.Text = "F10 Continue to Delete"
  frmTaxMsgWOpts.cmdExit.Text = "ESC Abort Delete"
  frmTaxMsgWOpts.Show vbModal
  If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
    Unload frmTaxMsgWOpts
    Close
    Exit Sub
  Else
    Unload frmTaxMsgWOpts
    MainLog ("DELETE: User elected to delete the current penalty setup file after being warned.")
  End If
  
  KillFile TaxPenHandling
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, SetUpRec
  Select Case ThisPenIdx
    Case 6
      SetUpRec.OptRev1 = ""
      frmTaxSystemSetup.vaSpread1.Col = 1
      frmTaxSystemSetup.vaSpread1.Row = 6
      frmTaxSystemSetup.vaSpread1.Text = ""
      frmTaxSystemSetup.vaSpread1.Lock = False
'      frmTaxSystemSetup.vaSpread1.Col = 2
'      frmTaxSystemSetup.vaSpread1.Row = 6
'      frmTaxSystemSetup.vaSpread1.Text = "No"
    Case 7
      SetUpRec.OptRev2 = ""
      frmTaxSystemSetup.vaSpread1.Col = 1
      frmTaxSystemSetup.vaSpread1.Row = 7
      frmTaxSystemSetup.vaSpread1.Text = ""
      frmTaxSystemSetup.vaSpread1.Lock = False
'      frmTaxSystemSetup.vaSpread1.Col = 2
'      frmTaxSystemSetup.vaSpread1.Row = 7
'      frmTaxSystemSetup.vaSpread1.Text = "No"
    Case 8
      SetUpRec.OptRev3 = ""
      frmTaxSystemSetup.vaSpread1.Col = 1
      frmTaxSystemSetup.vaSpread1.Row = 8
      frmTaxSystemSetup.vaSpread1.Text = ""
      frmTaxSystemSetup.vaSpread1.Lock = False
'      frmTaxSystemSetup.vaSpread1.Col = 2
'      frmTaxSystemSetup.vaSpread1.Row = 8
'      frmTaxSystemSetup.vaSpread1.Text = "No"
  End Select
  Put MHandle, 1, SetUpRec
  Close MHandle
'  frmTaxSystemSetup.PenIdx = 0
  frmTaxMsg.Label1.Caption = "Deletion has completed successfully."
  frmTaxMsg.Label1.Top = 900
  frmTaxMsg.Show vbModal
  Unload Me
  DoEvents
End Sub

Private Sub cmdSave_Click()
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  Dim SetUpRec As TaxMasterType
  Dim MHandle As Integer
  Dim RevDesc1$
  Dim RevDesc2$
  Dim RevDesc3$
  Dim RevDesc4$
  Dim RevDesc5$
  Dim RevDesc6$
  Dim RevDesc7$
  Dim RevDesc8$
  
  If chkOptRev1.Caption = "Not In Use" Then
    If chkOptRev1.Value = 1 Then
      frmTaxMsg.Label1.Caption = "Optional Revenue #1 is not in use. Applying a penalty is not allowed. Saving this value as true is aborted."
      frmTaxMsg.Label1.Top = 700
      frmTaxMsg.Show vbModal
      chkOptRev1.Value = 0
    End If
  End If
  
  If chkOptRev2.Caption = "Not In Use" Then
    If chkOptRev2.Value = 1 Then
      frmTaxMsg.Label1.Caption = "Optional Revenue #2 is not in use. Applying a penalty is not allowed. Saving this value as true is aborted."
      frmTaxMsg.Label1.Top = 700
      frmTaxMsg.Show vbModal
      chkOptRev2.Value = 0
    End If
  End If
  
  If chkOptRev3.Caption = "Not In Use" Then
    If chkOptRev3.Value = 1 Then
      frmTaxMsg.Label1.Caption = "Optional Revenue #3 is not in use. Applying a penalty is not allowed. Saving this value as true is aborted."
      frmTaxMsg.Label1.Top = 700
      frmTaxMsg.Show vbModal
      chkOptRev3.Value = 0
    End If
  End If
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, SetUpRec
  
  If OptPenOpt1.Value = True Or OptPenOpt2.Value = True Or OptPenOpt3.Value = True Then
    If QPTrim$(fptxtPenDesc.Text) = "" Then
      frmTaxMsg.Label1.Caption = "You must supply a penalty description since you have elected to include penalty as a revenue source."
      frmTaxMsg.Label1.Top = 800
      frmTaxMsg.Show vbModal
      fptxtPenDesc.SetFocus
      Close MHandle
      Exit Sub
    End If
  End If
  
  If chkPersProp.Value = 0 And chkRealEstate.Value = 0 And chkInterest.Value = 0 And chkAdvertising.Value = 0 _
    And chkLateList.Value = 0 And chkOptRev1.Value = 0 And chkOptRev2.Value = 0 And chkOptRev3 = 0 Then
      frmTaxMsgWOpts.Label1.Caption = "You have elected to include penalty as a revenue source but no sources of revenue have been designated as applicable to a penalty charge. If you wish to save anyway press F10. Otherwise, press ESC to review and edit."
      frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
      frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
      frmTaxMsgWOpts.Show vbModal
      If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        chkPersProp.SetFocus
        Close MHandle
        Exit Sub
      Else
        MainLog ("ERROR: User warned that they are saving the penalty setup without designating any revenues to have penalties applied to. The user elected to save anyway.")
        Unload frmTaxMsgWOpts
      End If
  End If
  
  If fpCurrFlatRate.Value = 0 And fpdblPct.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "You are attempting to save the penalty setup data without designating an amount for either the flat rate or a percentage. If you wish to save anyway then press F10 to continue. Otherwise, press ESC to review and edit."
    frmTaxMsgWOpts.Label1.Top = 650
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpCurrFlatRate.SetFocus
      Close MHandle
      Exit Sub
    Else
      MainLog ("ERROR: User warned that they are saving the penalty setup with a zero value for both the flat rate and the perentage rate. The user elected to save anyway.")
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If fpCurrFlatRate.Value = 0 And fpcmbMethod.Text = "Use flat rate and percentage." Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use both the flat rate and percentage rate for penalty applications but no value has been entered for flat rate. If you wish to save anyway then press F10. If you wish to review and edit then press ESC."
    frmTaxMsgWOpts.Label1.Top = 650
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpCurrFlatRate.SetFocus
      Close MHandle
      Exit Sub
    Else
      MainLog ("ERROR: User warned that they are saving the penalty setup using the flat rate and percentage method but no flat rate amount has been entered. The user elected to save anyway.")
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If fpdblPct.Value = 0 And fpcmbMethod.Text = "Use flat rate and percentage." Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use both the flat rate and percentage rate for penalty applications but no value has been entered for percentage rate. If you wish to save anyway then press F10. If you wish to review and edit then press ESC."
    frmTaxMsgWOpts.Label1.Top = 650
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpdblPct.SetFocus
      Close MHandle
      Exit Sub
    Else
      MainLog ("ERROR: User warned that they are saving the penalty setup using the flat rate and percentage method but no percentage rate amount has been entered. The user elected to save anyway.")
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If fpCurrFlatRate.Value = 0 And fpcmbMethod.Text = "Use flat rate only." Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use only the flat rate for penalty applications but no value has been entered for flat rate. If you wish to save anyway then press F10. If you wish to review and edit then press ESC."
    frmTaxMsgWOpts.Label1.Top = 650
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpCurrFlatRate.SetFocus
      Close MHandle
      Exit Sub
    Else
      MainLog ("ERROR: User warned that they are saving the penalty setup using the flat rate method only but no flat rate amount has been entered. The user elected to save anyway.")
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If fpdblPct.Value = 0 And fpcmbMethod.Text = "Use percentage only." Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use only the percentage rate for penalty applications but no value has been entered for the percentage rate. If you wish to save anyway then press F10. If you wish to review and edit then press ESC."
    frmTaxMsgWOpts.Label1.Top = 650
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpdblPct.SetFocus
      Close MHandle
      Exit Sub
    Else
      MainLog ("ERROR: User warned that they are saving the penalty setup using the percentage rate method only but no percentage rate amount has been entered. The user elected to save anyway.")
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If OptPenOpt1.Value = True Then
    If ThisPenIdx <> 6 And ThisPenIdx <> 0 Then
      frmTaxMsgWOpts.Label1.Caption = "You are changing the penalty option from " + CStr(ThisPenIdx) + " to 6. Press F10 to save and the program will update the revenue spreadsheet. Otherwise, press ESC to abort the save procedure."
      frmTaxMsgWOpts.Label1.Top = 700
      frmTaxMsgWOpts.Show vbModal
      If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
        Close MHandle
        Unload frmTaxMsgWOpts
        If OptPenOpt1.Enabled = True Then
          OptPenOpt1.SetFocus
        End If
        Exit Sub
      Else
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = ThisPenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        If ThisPenIdx = 7 Then
          SetUpRec.OptRev2 = ""
          chkOptRev2.Value = 0
        ElseIf ThisPenIdx = 8 Then
          SetUpRec.OptRev3 = ""
          chkOptRev3.Value = 0
        End If
      End If
    End If
    chkOptRev1.Value = 0
    SetUpRec.OptRev1 = QPTrim$(fptxtPenDesc.Text)
    frmTaxSystemSetup.vaSpread1.Col = 1
    frmTaxSystemSetup.vaSpread1.Row = 6
    frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
    frmTaxSystemSetup.vaSpread1.Lock = True
'    frmTaxSystemSetup.vaSpread1.Col = 2
'    frmTaxSystemSetup.vaSpread1.Text = "No"
'    frmTaxSystemSetup.PenIdx = 6
  ElseIf OptPenOpt2.Value = True Then
    If ThisPenIdx <> 7 And ThisPenIdx <> 0 Then
      frmTaxMsgWOpts.Label1.Caption = "You are changing the penalty option from " + CStr(ThisPenIdx) + " to 7. Press F10 to save and the program will update the revenue spreadsheet. Otherwise, press ESC to abort the save procedure."
      frmTaxMsgWOpts.Label1.Top = 700
      frmTaxMsgWOpts.Show vbModal
      If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
        Close MHandle
        Unload frmTaxMsgWOpts
        If OptPenOpt2.Enabled = True Then
          OptPenOpt2.SetFocus
        End If
        Exit Sub
      Else
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = ThisPenIdx
        frmTaxSystemSetup.vaSpread1.Lock = False
        frmTaxSystemSetup.vaSpread1.Text = ""
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        If ThisPenIdx = 6 Then
          SetUpRec.OptRev1 = ""
          chkOptRev1.Value = 0
        ElseIf ThisPenIdx = 8 Then
          SetUpRec.OptRev3 = ""
          chkOptRev3.Value = 0
        End If
      End If
    End If
    chkOptRev2.Value = 0
    SetUpRec.OptRev2 = QPTrim$(fptxtPenDesc.Text)
    frmTaxSystemSetup.vaSpread1.Col = 1
    frmTaxSystemSetup.vaSpread1.Row = 7
    frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
    frmTaxSystemSetup.vaSpread1.Lock = True
'    frmTaxSystemSetup.vaSpread1.Col = 2
'    frmTaxSystemSetup.vaSpread1.Text = "No"
'    frmTaxSystemSetup.PenIdx = 7
  ElseIf OptPenOpt3.Value = True Then
    If ThisPenIdx <> 8 And ThisPenIdx <> 0 Then
      frmTaxMsgWOpts.Label1.Caption = "You are changing the penalty option from " + CStr(ThisPenIdx) + " to 8. Press F10 to save and the program will update the revenue spreadsheet. Otherwise, press ESC to abort the save procedure."
      frmTaxMsgWOpts.Label1.Top = 700
      frmTaxMsgWOpts.Show vbModal
      If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
        Close MHandle
        Unload frmTaxMsgWOpts
        If OptPenOpt3.Enabled = True Then
          OptPenOpt3.SetFocus
        End If
        Exit Sub
      Else
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = ThisPenIdx
        frmTaxSystemSetup.vaSpread1.Lock = False
        frmTaxSystemSetup.vaSpread1.Text = ""
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        If ThisPenIdx = 6 Then
          SetUpRec.OptRev1 = ""
          chkOptRev1.Value = 0
        ElseIf ThisPenIdx = 7 Then
          SetUpRec.OptRev2 = ""
          chkOptRev2.Value = 0
        End If
      End If
    End If
    chkOptRev3.Value = 0
    SetUpRec.OptRev3 = QPTrim$(fptxtPenDesc.Text)
    frmTaxSystemSetup.vaSpread1.Col = 1
    frmTaxSystemSetup.vaSpread1.Row = 8
    frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
    frmTaxSystemSetup.vaSpread1.Lock = True
'    frmTaxSystemSetup.vaSpread1.Col = 2
'    frmTaxSystemSetup.vaSpread1.Text = "No"
'    frmTaxSystemSetup.PenIdx = 8
  End If
  Put MHandle, 1, SetUpRec
  Close MHandle
  
  RevDesc1$ = "Property Tax"
  RevDesc2$ = "Real Estate Tax"
  RevDesc3$ = "Interest"
  RevDesc4$ = "Advertising/Col"
  RevDesc5$ = "Late Listing"
  RevDesc6$ = QPTrim$(SetUpRec.OptRev1)
  RevDesc7$ = QPTrim$(SetUpRec.OptRev2)
  RevDesc8$ = QPTrim$(SetUpRec.OptRev3)
  
  If Exist(TaxPenHandling) Then
    OpenTaxPenFile PHandle
    Get PHandle, 1, PenRec
    PenRec.PenFlat = fpCurrFlatRate.Value
    PenRec.PenPct = fpdblPct.Value
    If fpcmbMethod.Text = "Use flat rate and percentage." Then
      PenRec.UseBoth = "Y"
      PenRec.UseFlat = "N"
      PenRec.UsePct = "N"
      If fpcmbCombine.Text = "Whichever is least." Then
        PenRec.UseLow = "Y"
        PenRec.UseHigh = "N"
      ElseIf fpcmbCombine.Text = "Whichever is most." Then
        PenRec.UseHigh = "Y"
        PenRec.UseLow = "N"
      End If
    ElseIf fpcmbMethod.Text = "Use flat rate only." Then
      PenRec.UseBoth = "N"
      PenRec.UseFlat = "Y"
      PenRec.UsePct = "N"
      PenRec.UseHigh = "N"
      PenRec.UseLow = "N"
    ElseIf fpcmbMethod.Text = "Use percentage only." Then
      PenRec.UseBoth = "N"
      PenRec.UseFlat = "N"
      PenRec.UsePct = "Y"
      PenRec.UseHigh = "N"
      PenRec.UseLow = "N"
    End If
    PenRec.Rev1Name = RevDesc1$
    If chkPersProp.Value = 1 Then
      PenRec.AppToRev1 = "Y"
    Else
      PenRec.AppToRev1 = "N"
    End If
    PenRec.Rev2Name = RevDesc2$
    If chkRealEstate.Value = 1 Then
      PenRec.AppToRev2 = "Y"
    Else
      PenRec.AppToRev2 = "N"
    End If
    PenRec.Rev3Name = RevDesc3$
    If chkInterest.Value = 1 Then
      PenRec.AppToRev3 = "Y"
    Else
      PenRec.AppToRev3 = "N"
    End If
    PenRec.Rev4Name = RevDesc4$
    If chkAdvertising.Value = 1 Then
      PenRec.AppToRev4 = "Y"
    Else
      PenRec.AppToRev4 = "N"
    End If
    PenRec.Rev5Name = RevDesc5$
    If chkLateList.Value = 1 Then
      PenRec.AppToRev5 = "Y"
    Else
      PenRec.AppToRev5 = "N"
    End If
    PenRec.Rev6Name = RevDesc6$
    If chkOptRev1.Value = 1 Then
      PenRec.AppToRev6 = "Y"
    Else
      PenRec.AppToRev6 = "N"
    End If
    PenRec.Rev7Name = RevDesc7$
    If chkOptRev2.Value = 1 Then
      PenRec.AppToRev7 = "Y"
    Else
      PenRec.AppToRev7 = "N"
    End If
    PenRec.Rev8Name = RevDesc8$
    If chkOptRev3.Value = 1 Then
      PenRec.AppToRev8 = "Y"
    Else
      PenRec.AppToRev8 = "N"
    End If
    PenRec.PenDesc = QPTrim$(fptxtPenDesc.Text)
    If OptPenOpt1.Value = True Then
      PenRec.PenIdx = 6
    ElseIf OptPenOpt2.Value = True Then
      PenRec.PenIdx = 7
    ElseIf OptPenOpt3.Value = True Then
      PenRec.PenIdx = 8
    End If
  Else
    PenRec.PenFlat = fpCurrFlatRate.Value
    PenRec.PenPct = fpdblPct.Value
    If fpcmbMethod.Text = "Use flat rate and percentage." Then
      PenRec.UseBoth = "Y"
      PenRec.UseFlat = "N"
      PenRec.UsePct = "N"
      If fpcmbCombine.Text = "Whichever is least." Then
        PenRec.UseLow = "Y"
        PenRec.UseHigh = "N"
      ElseIf fpcmbCombine.Text = "Whichever is most." Then
        PenRec.UseHigh = "Y"
        PenRec.UseLow = "N"
      End If
    ElseIf fpcmbMethod.Text = "Use flat rate only." Then
      PenRec.UseBoth = "N"
      PenRec.UseFlat = "Y"
      PenRec.UsePct = "N"
      PenRec.UseHigh = "N"
      PenRec.UseLow = "N"
    ElseIf fpcmbMethod.Text = "Use percentage only." Then
      PenRec.UseBoth = "N"
      PenRec.UseFlat = "N"
      PenRec.UsePct = "Y"
      PenRec.UseHigh = "N"
      PenRec.UseLow = "N"
    End If
    PenRec.Rev1Name = RevDesc1$
    If chkPersProp.Value = 1 Then
      PenRec.AppToRev1 = "Y"
    Else
      PenRec.AppToRev1 = "N"
    End If
    PenRec.Rev2Name = RevDesc2$
    If chkRealEstate.Value = 1 Then
      PenRec.AppToRev2 = "Y"
    Else
      PenRec.AppToRev2 = "N"
    End If
    PenRec.Rev3Name = RevDesc3$
    If chkInterest.Value = 1 Then
      PenRec.AppToRev3 = "Y"
    Else
      PenRec.AppToRev3 = "N"
    End If
    PenRec.Rev4Name = RevDesc4$
    If chkAdvertising.Value = 1 Then
      PenRec.AppToRev4 = "Y"
    Else
      PenRec.AppToRev4 = "N"
    End If
    PenRec.Rev5Name = RevDesc5$
    If chkLateList.Value = 1 Then
      PenRec.AppToRev5 = "Y"
    Else
      PenRec.AppToRev5 = "N"
    End If
    PenRec.Rev6Name = RevDesc6$
    If chkOptRev1.Value = 1 Then
      PenRec.AppToRev6 = "Y"
    Else
      PenRec.AppToRev6 = "N"
    End If
    PenRec.Rev7Name = RevDesc7$
    If chkOptRev2.Value = 1 Then
      PenRec.AppToRev7 = "Y"
    Else
      PenRec.AppToRev7 = "N"
    End If
    PenRec.Rev8Name = RevDesc8$
    If chkOptRev3.Value = 1 Then
      PenRec.AppToRev8 = "Y"
    Else
      PenRec.AppToRev8 = "N"
    End If
    PenRec.PenDesc = QPTrim$(fptxtPenDesc.Text)
    If OptPenOpt1.Value = True Then
      PenRec.PenIdx = 6
    ElseIf OptPenOpt2.Value = True Then
      PenRec.PenIdx = 7
    ElseIf OptPenOpt3.Value = True Then
      PenRec.PenIdx = 8
    End If
    OpenTaxPenFile PHandle
  End If
  Put PHandle, 1, PenRec
  Close PHandle
  
  Call LogSaves
  Call Savemsg(900, "Your penalty data has been saved successfully.")
  Unload Me
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF5:
      Call cmdDelete_Click
      KeyCode = 0
    Case vbKeyF10:
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select
  
End Sub

Private Sub Form_Load()
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  Dim SetUpRec As TaxMasterType
  Dim MHandle As Integer
  Dim RevDesc1$
  Dim RevDesc2$
  Dim RevDesc3$
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, SetUpRec
  Close MHandle
  RevDesc1 = QPTrim$(SetUpRec.OptRev1)
  If RevDesc1 = "" Then RevDesc1 = "Not Used"
  RevDesc2 = QPTrim$(SetUpRec.OptRev2)
  If RevDesc2 = "" Then RevDesc2 = "Not Used"
  RevDesc3 = QPTrim$(SetUpRec.OptRev3)
  If RevDesc3 = "" Then RevDesc3 = "Not Used"
  
  fpcmbMethod.Text = "Use percentage only."
  fpcmbMethod.AddItem "Use flat rate only."
  fpcmbMethod.AddItem "Use percentage only."
  fpcmbMethod.AddItem "Use flat rate and percentage."
  fpcmbCombine.Text = "Whichever is most."
  fpcmbCombine.AddItem "Whichever is least."
  fpcmbCombine.AddItem "Whichever is most."
  fpcmbCombine.Enabled = False
  
  If Exist(TaxPenHandling) Then
    OpenTaxPenFile PHandle
    Get PHandle, 1, PenRec
    Close PHandle
    fpCurrFlatRate = PenRec.PenFlat
    TempPenFlat = PenRec.PenFlat
    fpdblPct = PenRec.PenPct
    TempPenPct = PenRec.PenPct
    If PenRec.UseBoth = "Y" Then
      fpcmbMethod.Text = "Use flat rate and percentage."
    ElseIf PenRec.UseFlat = "Y" Then
      fpcmbMethod.Text = "Use flat rate only."
    ElseIf PenRec.UsePct = "Y" Then
      fpcmbMethod.Text = "Use percentage only."
    End If
    If PenRec.UseBoth = "Y" Then
      fpcmbCombine.Enabled = True
      If PenRec.UseLow = "Y" Then
        fpcmbCombine.Text = "Whichever is least."
      ElseIf PenRec.UseHigh = "Y" Then
        fpcmbCombine.Text = "Whichever is most."
      End If
    End If
    TempUsePct = PenRec.UsePct
    TempUseFlat = PenRec.UseFlat
    TempUseBoth = PenRec.UseBoth
    TempUseHigh = PenRec.UseHigh
    TempUseLow = PenRec.UseLow
    chkPersProp.Caption = QPTrim$(PenRec.Rev1Name)
    If PenRec.AppToRev1 = "Y" Then
      chkPersProp.Value = 1
    Else
      chkPersProp.Value = 0
    End If
    TempAppToRev1 = PenRec.AppToRev1
    chkRealEstate.Caption = QPTrim$(PenRec.Rev2Name)
    If PenRec.AppToRev2 = "Y" Then
      chkRealEstate.Value = 1
    Else
      chkRealEstate.Value = 0
    End If
    TempAppToRev2 = PenRec.AppToRev2
    chkInterest.Caption = QPTrim$(PenRec.Rev3Name)
    If PenRec.AppToRev3 = "Y" Then
      chkInterest.Value = 1
    Else
      chkInterest.Value = 0
    End If
    TempAppToRev3 = PenRec.AppToRev3
    chkAdvertising.Caption = QPTrim$(PenRec.Rev4Name)
    If PenRec.AppToRev4 = "Y" Then
      chkAdvertising.Value = 1
    Else
      chkAdvertising.Value = 0
    End If
    TempAppToRev4 = PenRec.AppToRev4
    chkLateList.Caption = QPTrim$(PenRec.Rev5Name)
    If PenRec.AppToRev5 = "Y" Then
      chkLateList.Value = 1
    Else
      chkLateList.Value = 0
    End If
    TempAppToRev5 = PenRec.AppToRev5
    If QPTrim$(PenRec.Rev6Name) = "" Then
      chkOptRev1.Caption = "Not In Use"
    Else
      chkOptRev1.Caption = QPTrim$(PenRec.Rev6Name)
    End If
    If PenRec.AppToRev6 = "Y" Then
      chkOptRev1.Value = 1
    Else
      chkOptRev1.Value = 0
    End If
    TempAppToRev6 = PenRec.AppToRev6
    TempRev6Name = chkOptRev1.Caption
    If QPTrim$(PenRec.Rev7Name) = "" Then
      chkOptRev2.Caption = "Not In Use"
    Else
      chkOptRev2.Caption = QPTrim$(PenRec.Rev7Name)
    End If
    If PenRec.AppToRev7 = "Y" Then
      chkOptRev2.Value = 1
    Else
      chkOptRev2.Value = 0
    End If
    TempAppToRev7 = PenRec.AppToRev7
    TempRev7Name = chkOptRev2.Caption
    If QPTrim$(PenRec.Rev8Name) = "" Then
      chkOptRev3.Caption = "Not In Use"
    Else
      chkOptRev3.Caption = QPTrim$(PenRec.Rev8Name)
    End If
    If PenRec.AppToRev8 = "Y" Then
      chkOptRev3.Value = 1
    Else
      chkOptRev3.Value = 0
    End If
    TempAppToRev8 = PenRec.AppToRev8
    TempRev8Name = chkOptRev3.Caption
    If PenRec.PenIdx = 6 Then
      ThisPenIdx = 6
      OptPenOpt1.Caption = PenRec.PenDesc
      OptPenOpt1.Value = True
    Else
      If RevDesc1 = "" Or RevDesc1 = "Not Used" Then
        OptPenOpt1.Caption = "Available"
      Else
        OptPenOpt1.Caption = "Not Available"
        OptPenOpt1.Enabled = False
      End If
      OptPenOpt1.Value = False
    End If
    If PenRec.PenIdx = 7 Then
      ThisPenIdx = 7
      OptPenOpt2.Caption = PenRec.PenDesc
      OptPenOpt2.Value = True
    Else
      If RevDesc2 = "" Or RevDesc2 = "Not Used" Then
        OptPenOpt2.Caption = "Available"
      Else
        OptPenOpt2.Caption = "Not Available"
        OptPenOpt2.Enabled = False
      End If
      OptPenOpt2.Value = False
    End If
    If PenRec.PenIdx = 8 Then
      ThisPenIdx = 8
      OptPenOpt3.Caption = PenRec.PenDesc
      OptPenOpt3.Value = True
    Else
      If RevDesc3 = "" Or RevDesc3 = "Not Used" Then
        OptPenOpt3.Caption = "Available"
      Else
        OptPenOpt3.Caption = "Not Available"
        OptPenOpt3.Enabled = False
      End If
      OptPenOpt3.Value = False
    End If
    fptxtPenDesc.Text = QPTrim$(PenRec.PenDesc)
    TempPenDesc = QPTrim$(PenRec.PenDesc)
  Else
    If RevDesc1 = "" Or RevDesc1 = "Not Used" Then
      OptPenOpt1.Caption = "Available"
    Else
      OptPenOpt1.Caption = "Not Available"
      OptPenOpt1.Enabled = False
    End If
    If RevDesc2 = "" Or RevDesc2 = "Not Used" Then
      OptPenOpt2.Caption = "Available"
    Else
      OptPenOpt2.Caption = "Not Available"
      OptPenOpt2.Enabled = False
    End If
    If RevDesc3 = "" Or RevDesc3 = "Not Used" Then
      OptPenOpt3.Caption = "Available"
    Else
      OptPenOpt3.Caption = "Not Available"
      OptPenOpt3.Enabled = False
    End If
  TempPenIdx = ThisPenIdx
  End If
  
End Sub

Private Sub fpcmbCombine_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCombine.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCombine.ListIndex = -1
  End If
  If fpcmbCombine.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      chkPersProp.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbMethod_Change()
  If fpcmbMethod.Text = "Use flat rate and percentage." Then
    fpcmbCombine.Enabled = True
  Else
    fpcmbCombine.Enabled = False
  End If
End Sub

Private Function Check4Changes() As Boolean
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisDesc As String
  Dim ThisDbl As Double
  Dim OptStr As String
  Dim OptInt As Integer
  Dim SetUpRec As TaxMasterType
  Dim MHandle As Integer
  'on error goto ERRORSTUFF
  
  Check4Changes = False
  If Exist(TaxPenHandling) Then
    OpenTaxPenFile PHandle
    Get PHandle, 1, PenRec
  Else
    frmTaxMsgWOpts.Label1.Caption = "Do you wish to exit without saving any changes? Press F10 to Save. Otherwise, press ESC to exit without saving."
    frmTaxMsgWOpts.Label1.Top = 800
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      Call cmdSave_Click
      Exit Function
    Else
      Unload frmTaxMsgWOpts
      Exit Function
    End If
  End If
  
  Set ThisControl = fpCurrFlatRate
  ThisDbl = PenRec.PenFlat
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Flat Rate' field has been changed from " + Using$("$#,##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PenRec.PenFlat = CDbl(ThisControl.Text)
      Put PHandle, 1, PenRec
      Call Savemsg(900, "Penalty Flat Rate has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpdblPct
  ThisDbl = PenRec.PenPct
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Pct Rate' field has been changed from " + Using$("$#,##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      PenRec.PenPct = CDbl(ThisControl.Text)
      Put PHandle, 1, PenRec
      Call Savemsg(900, "Penalty Pct Rate has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbMethod
  If fpcmbMethod.Text = "Use flat rate and percentage." Then
    If PenRec.UseFlat = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Application Method' field has been changed from 'Use flat rate only.' to 'Use flat rate and percentage.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.UseFlat = "N"
        PenRec.UseBoth = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Application Method has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    ElseIf PenRec.UsePct = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Application Method' field has been changed from 'Use percentage only.' to 'Use flat rate and percentage.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.UsePct = "N"
        PenRec.UseBoth = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Application Method has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf fpcmbMethod.Text = "Use flat rate only." Then
    If PenRec.UseBoth = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Application Method' field has been changed from 'Use flat rate and percentage.' to 'Use flat rate only.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.UseFlat = "Y"
        PenRec.UseBoth = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Application Method has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    ElseIf PenRec.UsePct = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Application Method' field has been changed from 'Use percentage only.' to 'Use flat rate only.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.UsePct = "Y"
        PenRec.UseBoth = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Application Method has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf fpcmbMethod.Text = "Use percentage only." Then
    If PenRec.UseBoth = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Application Method' field has been changed from 'Use flat rate and percentage.' to 'Use percentage only.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.UsePct = "Y"
        PenRec.UseBoth = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Application Method has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    ElseIf PenRec.UseFlat = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Application Method' field has been changed from 'Use flat rate only.' to 'Use percentage only.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.UsePct = "Y"
        PenRec.UseBoth = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Application Method has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  Set ThisControl = fpcmbCombine
  If fpcmbMethod.Text = "Use flat rate and percentage." Then
    If fpcmbCombine.Text = "Whichever is least." Then
      If PenRec.UseHigh = "Y" Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Combination Type' field has been changed from 'Whichever is most.' to 'Whichever is least.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          PenRec.UseLow = "Y"
          PenRec.UseHigh = "N"
            Put PHandle, 1, PenRec
        Call Savemsg(900, "Combination Type has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    ElseIf fpcmbCombine.Text = "Whichever is most." Then
      If PenRec.UseLow = "Y" Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Combination Type' field has been changed from 'Whichever is least.' to 'Whichever is most.' Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          PenRec.UseLow = "N"
          PenRec.UseHigh = "Y"
            Put PHandle, 1, PenRec
        Call Savemsg(900, "Combination Type has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
  End If
    
  Set ThisControl = chkPersProp
  If chkPersProp.Value = 1 Then
    If PenRec.AppToRev1 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Personal Property' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev1 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Personal Property has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkPersProp.Value = 0 Then
    If PenRec.AppToRev1 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Personal Property' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev1 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Personal Property has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
        
  Set ThisControl = chkRealEstate
  If chkRealEstate = 1 Then
    If PenRec.AppToRev2 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Real Estate' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev2 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Real Estate has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkRealEstate.Value = 0 Then
    If PenRec.AppToRev2 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Real Estate' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev2 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Real Estate has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
        
  Set ThisControl = chkInterest
  If chkInterest = 1 Then
    If PenRec.AppToRev3 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Interest' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev3 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Interest has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkInterest.Value = 0 Then
    If PenRec.AppToRev3 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Interest' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev3 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Interest has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
   
  Set ThisControl = chkAdvertising
  If chkAdvertising = 1 Then
    If PenRec.AppToRev4 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Advertising/Collection' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev4 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Advertising/Collection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkAdvertising.Value = 0 Then
    If PenRec.AppToRev4 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Advertising/Collection' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev4 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Advertising/Collection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  Set ThisControl = chkLateList
  If chkLateList = 1 Then
    If PenRec.AppToRev5 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Late Listing' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev5 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Late Listing has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkLateList.Value = 0 Then
    If PenRec.AppToRev5 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Late Listing' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev5 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Late Listing has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  If chkOptRev1.Caption = "Not In Use" Then
    If chkOptRev1.Value = 1 Then
      frmTaxMsg.Label1.Caption = "Optional Revenue #1 is not in use. Applying a penalty is not allowed. "
      frmTaxMsg.Label1.Top = 800
      frmTaxMsg.Show vbModal
      PenRec.AppToRev6 = "N"
      Put PHandle, 1, PenRec
      GoTo Next1
    End If
  End If
  Set ThisControl = chkOptRev1
  If chkOptRev1 = 1 Then
    If PenRec.AppToRev6 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Optional Revenue #1' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev6 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Optional Revenue #1 has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkOptRev1.Value = 0 Then
    If PenRec.AppToRev6 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Optional Revenue #1' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev6 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Optional Revenue #1 has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
Next1:
  
  If chkOptRev2.Caption = "Not In Use" Then
    If chkOptRev2.Value = 1 Then
      frmTaxMsg.Label1.Caption = "Optional Revenue #2 is not in use. Applying a penalty is not allowed. "
      frmTaxMsg.Label1.Top = 800
      frmTaxMsg.Show vbModal
      PenRec.AppToRev7 = "N"
      Put PHandle, 1, PenRec
      GoTo Next2
    End If
  End If
  Set ThisControl = chkOptRev2
  If chkOptRev2 = 1 Then
    If PenRec.AppToRev7 = "N" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Optional Revenue #2' field has been changed from 'False' to 'True'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev7 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Optional Revenue #2 has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkOptRev2.Value = 0 Then
    If PenRec.AppToRev7 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Optional Revenue #2' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev7 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Optional Revenue #2 has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
Next2:

  If chkOptRev3.Caption = "Not In Use" Then
    If chkOptRev3.Value = 1 Then
      frmTaxMsg.Label1.Caption = "Optional Revenue #3 is not in use. Applying a penalty is not allowed. "
      frmTaxMsg.Label1.Top = 800
      frmTaxMsg.Show vbModal
      PenRec.AppToRev8 = "N"
      Put PHandle, 1, PenRec
      GoTo Next3
    End If
  End If
  Set ThisControl = chkOptRev3
  If chkOptRev3 = 1 Then
    If PenRec.AppToRev8 = "N" Then
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev8 = "Y"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Optional Revenue #3 has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf chkOptRev3.Value = 0 Then
    If PenRec.AppToRev8 = "Y" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Penalty to Optional Revenue #3' field has been changed from 'True' to 'False'. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        PenRec.AppToRev8 = "N"
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Apply Penalty to Optional Revenue #3 has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
Next3:

  Set ThisControl = fptxtPenDesc
  ThisDesc = QPTrim$(PenRec.PenDesc)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    If QPTrim$(ThisControl.Text) = "" Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Description' field has been changed from " + ThisDesc + " to BLANK. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    Else
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Description' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    End If
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      If OptPenOpt1.Value = True Or OptPenOpt2.Value = True Or OptPenOpt3.Value = True Then
        If QPTrim$(fptxtPenDesc.Text) = "" Then
          frmTaxMsg.Label1.Caption = "You must supply a penalty description since you have elected to include penalty as a revenue source."
          frmTaxMsg.Label1.Top = 800
          frmTaxMsg.Show vbModal
          fptxtPenDesc.SetFocus
          Close PHandle
          Check4Changes = True
          Exit Function
        End If
      End If
      PenRec.PenDesc = QPTrim$(ThisControl.Text)
      Put PHandle, 1, PenRec
      Call Savemsg(900, "Penalty Description has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, SetUpRec

  If OptPenOpt1 = True Then
    Set ThisControl = OptPenOpt1
    If PenRec.PenIdx = 7 Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Revenue Selection' field has been changed from 7 to 6. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = PenRec.PenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = 6
        frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
        frmTaxSystemSetup.vaSpread1.Lock = True
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        SetUpRec.OptRev1 = QPTrim$(fptxtPenDesc.Text)
        If PenRec.PenIdx = 7 Then
          SetUpRec.OptRev2 = ""
          PenRec.Rev7Name = ""
          chkOptRev2.Value = 0
        ElseIf PenRec.PenIdx = 8 Then
          SetUpRec.OptRev3 = ""
          PenRec.Rev8Name = ""
          chkOptRev3.Value = 0
        End If
        Put MHandle, 1, SetUpRec
        PenRec.PenIdx = 6
        PenRec.Rev6Name = QPTrim$(fptxtPenDesc.Text)
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Revenue Selection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    ElseIf PenRec.PenIdx = 8 Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Revenue Selection' field has been changed from 8 to 6. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = PenRec.PenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = 6
        frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
        frmTaxSystemSetup.vaSpread1.Lock = True
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        SetUpRec.OptRev1 = QPTrim$(fptxtPenDesc.Text)
        If PenRec.PenIdx = 7 Then
          SetUpRec.OptRev2 = ""
          chkOptRev2.Value = 0
          PenRec.Rev7Name = ""
        ElseIf PenRec.PenIdx = 8 Then
          SetUpRec.OptRev3 = ""
          chkOptRev3.Value = 0
          PenRec.Rev8Name = ""
        End If
        Put MHandle, 1, SetUpRec
        PenRec.PenIdx = 6
        PenRec.Rev6Name = QPTrim$(fptxtPenDesc.Text)
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Revenue Selection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf OptPenOpt2 = True Then
    Set ThisControl = OptPenOpt2
    If PenRec.PenIdx = 6 Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Revenue Selection' field has been changed from 6 to 7. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = PenRec.PenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = 7
        frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
        frmTaxSystemSetup.vaSpread1.Lock = True
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        SetUpRec.OptRev2 = QPTrim$(fptxtPenDesc.Text)
        If PenRec.PenIdx = 6 Then
          SetUpRec.OptRev1 = ""
          chkOptRev1.Value = 0
          PenRec.Rev6Name = ""
        ElseIf PenRec.PenIdx = 8 Then
          SetUpRec.OptRev3 = ""
          chkOptRev3.Value = 0
          PenRec.Rev8Name = ""
        End If
        Put MHandle, 1, SetUpRec
        PenRec.PenIdx = 7
        PenRec.Rev7Name = QPTrim$(fptxtPenDesc.Text)
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Revenue Selection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    ElseIf PenRec.PenIdx = 8 Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Revenue Selection' field has been changed from 8 to 7. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = PenRec.PenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = 7
        frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
        frmTaxSystemSetup.vaSpread1.Lock = True
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        SetUpRec.OptRev2 = QPTrim$(fptxtPenDesc.Text)
        If PenRec.PenIdx = 6 Then
          SetUpRec.OptRev1 = ""
          chkOptRev1.Value = 0
          PenRec.Rev6Name = ""
        ElseIf PenRec.PenIdx = 8 Then
          SetUpRec.OptRev3 = ""
          chkOptRev3.Value = 0
          PenRec.Rev8Name = ""
        End If
        Put MHandle, 1, SetUpRec
        PenRec.PenIdx = 7
        PenRec.Rev7Name = QPTrim$(fptxtPenDesc.Text)
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Revenue Selection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  ElseIf OptPenOpt3 = True Then
    Set ThisControl = OptPenOpt3
    If PenRec.PenIdx = 6 Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Revenue Selection' field has been changed from 6 to 8. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = PenRec.PenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = 8
        frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
        frmTaxSystemSetup.vaSpread1.Lock = True
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        SetUpRec.OptRev3 = QPTrim$(fptxtPenDesc.Text)
        If PenRec.PenIdx = 7 Then
          SetUpRec.OptRev2 = ""
          chkOptRev2.Value = 0
          PenRec.Rev7Name = ""
        ElseIf PenRec.PenIdx = 6 Then
          SetUpRec.OptRev1 = ""
          chkOptRev1.Value = 0
          PenRec.Rev6Name = ""
        End If
        Put MHandle, 1, SetUpRec
        PenRec.Rev8Name = QPTrim$(fptxtPenDesc.Text)
        PenRec.PenIdx = 8
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Revenue Selection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    ElseIf PenRec.PenIdx = 7 Then
      frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Revenue Selection' field has been changed from 7 to 8. Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = PenRec.PenIdx
        frmTaxSystemSetup.vaSpread1.Text = ""
        frmTaxSystemSetup.vaSpread1.Lock = False
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        frmTaxSystemSetup.vaSpread1.Col = 1
        frmTaxSystemSetup.vaSpread1.Row = 8
        frmTaxSystemSetup.vaSpread1.Text = QPTrim$(fptxtPenDesc.Text)
        frmTaxSystemSetup.vaSpread1.Lock = True
'        frmTaxSystemSetup.vaSpread1.Col = 2
'        frmTaxSystemSetup.vaSpread1.Text = "No"
        SetUpRec.OptRev3 = QPTrim$(fptxtPenDesc.Text)
        If PenRec.PenIdx = 7 Then
          SetUpRec.OptRev2 = ""
          chkOptRev2.Value = 0
          PenRec.Rev7Name = ""
        ElseIf PenRec.PenIdx = 6 Then
          SetUpRec.OptRev1 = ""
          chkOptRev1.Value = 0
          PenRec.Rev6Name = ""
        End If
        Put MHandle, 1, SetUpRec
        PenRec.PenIdx = 8
        PenRec.Rev8Name = QPTrim$(fptxtPenDesc.Text)
        Put PHandle, 1, PenRec
        Call Savemsg(900, "Penalty Revenue Selection has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  Close MHandle
  
  If fpcmbMethod.Text = "Use flat rate only." And fpCurrFlatRate.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use a flat rate only for penalty applications but a zero value has been saved for the flat rate. If you wish to save anyway then press F10. Otherwise, press ESC to review and edit."
    frmTaxMsgWOpts.Label1.Top = 700
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      Close PHandle
      Check4Changes = True
      fpCurrFlatRate.SetFocus
      Exit Function
    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MainLog ("ERROR: User warned that they were saving a flat rate only penalty application method but the flat rate entered is zero. They elected to save anyway.")
    End If
  End If
  
  If fpcmbMethod.Text = "Use percentage only." And fpdblPct.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use a percentage only for penalty applications but a zero value has been saved for the percentage rate. If you wish to save anyway then press F10. Otherwise, press ESC to review and edit."
    frmTaxMsgWOpts.Label1.Top = 700
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      Close PHandle
      Check4Changes = True
      fpdblPct.SetFocus
      Exit Function
    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MainLog ("ERROR: User warned that they were saving a percentage only penalty application method but the percentage rate entered is zero. They elected to save anyway.")
    End If
  End If
  
  If fpCurrFlatRate.Value = 0 And fpdblPct.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "You are attempting to save the penalty setup data without designating an amount for either the flat rate or a percentage. If you wish to save anyway then press F10 to continue. Otherwise, press ESC to review and edit."
    frmTaxMsgWOpts.Label1.Top = 650
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fpCurrFlatRate.SetFocus
      Close PHandle
      Check4Changes = True
      Exit Function
    Else
      MainLog ("ERROR: User warned that they are saving the penalty setup with a zero value for both the flat rate and the perentage rate. The user elected to save anyway.")
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If fpcmbMethod.Text = "Use flat rate and percentage." And fpCurrFlatRate.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use a flat rate and percentage for penalty applications but a zero value has been saved for the flat rate. If you wish to save anyway then press F10. Otherwise, press ESC to review and edit."
    frmTaxMsgWOpts.Label1.Top = 700
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      Close PHandle
      Check4Changes = True
      fpCurrFlatRate.SetFocus
      Exit Function
    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MainLog ("ERROR: User warned that they were saving a flat rate and percentage penalty application method but the flat rate entered is zero. They elected to save anyway.")
    End If
  End If
  
  If fpcmbMethod.Text = "Use flat rate and percentage." And fpdblPct.Value = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "You have elected to use a flat rate and percentage for penalty applications but a zero value has been saved for the percentage rate. If you wish to save anyway then press F10. Otherwise, press ESC to review and edit."
    frmTaxMsgWOpts.Label1.Top = 700
    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      Close PHandle
      Check4Changes = True
      fpdblPct.SetFocus
      Exit Function
    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MainLog ("ERROR: User warned that they were saving a flat rate and percentage penalty application method but the percentage rate entered is zero. They elected to save anyway.")
    End If
  End If
  
  If chkPersProp.Value = 0 And chkRealEstate.Value = 0 And chkInterest.Value = 0 And chkAdvertising.Value = 0 _
    And chkLateList.Value = 0 And chkOptRev1.Value = 0 And chkOptRev2.Value = 0 And chkOptRev3 = 0 Then
      frmTaxMsgWOpts.Label1.Caption = "You have elected to include penalty as a revenue source but no sources of revenue have been designated as applicable to a penalty charge. If you wish to save anyway press F10. Otherwise, press ESC to review and edit."
      frmTaxMsgWOpts.cmdCont.Text = "F10 Continue Anyway"
      frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
      frmTaxMsgWOpts.Show vbModal
      If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        chkPersProp.SetFocus
        Close PHandle
        Check4Changes = True
        Exit Function
      Else
        MainLog ("ERROR: User warned that they are saving the penalty setup without designating any revenues to have penalties applied to. The user elected to save anyway.")
        Unload frmTaxMsgWOpts
      End If
  End If
  
  Close PHandle
  Call LogSaves
  
  Exit Function
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close PHandle
        Close MHandle
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        Close PHandle
        Close MHandle
        Check4Changes = True
        ThisControl.SetFocus
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPenaltySetup", "Check4Changes", Erl)
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
  
Private Sub LogSaves()
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  
  OpenTaxPenFile PHandle
  Get PHandle, 1, PenRec
  Close PHandle
  
  If TempPenIdx <> PenRec.PenIdx Then
    MainLog ("The penalty index number was changed from " + CStr(TempPenIdx) + " to " + CStr(PenRec.PenIdx) + " and saved.")
  End If
  
  If QPTrim$(TempPenDesc) <> QPTrim$(PenRec.PenDesc) Then
    MainLog ("The penalty description was changed from " + QPTrim$(TempPenDesc) + " to " + QPTrim$(PenRec.PenDesc) + " and saved.")
  End If
  
  If TempPenPct <> PenRec.PenPct Then
    MainLog ("The penalty percentage rate was changed from " + QPTrim$(Using$("##0.00", TempPenPct)) + "% to " + QPTrim$(Using$("##0.00", PenRec.PenPct)) + "% and saved.")
  End If
  
  If TempPenFlat <> PenRec.PenFlat Then
    MainLog ("The penalty flat rate was changed from " + QPTrim$(Using$("$#,##0.00", TempPenFlat)) + " to " + QPTrim$(Using$("$#,##0.00", PenRec.PenFlat)) + " and saved.")
  End If
  
  If TempUseBoth <> PenRec.UseBoth Then
    MainLog ("The 'use both' method of penalty application was changed from " + TempUseBoth + " to " + PenRec.UseBoth + " and saved.")
  End If
  
  If TempUseFlat <> PenRec.UseFlat Then
    MainLog ("The 'use flat rate only' method of penalty application was changed from " + TempUseFlat + " to " + PenRec.UseFlat + " and saved.")
  End If
  
  If TempUsePct <> PenRec.UsePct Then
    MainLog ("The 'use percentage only' method of penalty application was changed from " + TempUsePct + " to " + PenRec.UsePct + " and saved.")
  End If
  
  If TempUseHigh <> PenRec.UseHigh Then
    MainLog ("The penalty combined selection 'use high' was changed from " + TempUseHigh + " to " + PenRec.UseHigh + " and saved.")
  End If
  
  If TempUseLow <> PenRec.UseLow Then
    MainLog ("The penalty combined selection 'use low' was changed from " + TempUseLow + " to " + PenRec.UseLow + " and saved.")
  End If
  
  If TempAppToRev1 <> PenRec.AppToRev1 Then
    MainLog ("Penalty application to revenue #1 has changed from " + TempAppToRev1 + " to " + PenRec.AppToRev1 + " and saved.")
  End If

  If TempAppToRev2 <> PenRec.AppToRev2 Then
    MainLog ("Penalty application to revenue #2 has changed from " + TempAppToRev2 + " to " + PenRec.AppToRev2 + " and saved.")
  End If

  If TempAppToRev3 <> PenRec.AppToRev3 Then
    MainLog ("Penalty application to revenue #3 has changed from " + TempAppToRev3 + " to " + PenRec.AppToRev3 + " and saved.")
  End If

  If TempAppToRev4 <> PenRec.AppToRev4 Then
    MainLog ("Penalty application to revenue #4 has changed from " + TempAppToRev4 + " to " + PenRec.AppToRev4 + " and saved.")
  End If

  If TempAppToRev5 <> PenRec.AppToRev5 Then
    MainLog ("Penalty application to revenue #5 has changed from " + TempAppToRev5 + " to " + PenRec.AppToRev5 + " and saved.")
  End If

  If TempAppToRev6 <> PenRec.AppToRev6 Then
    MainLog ("Penalty application to revenue #6 has changed from " + TempAppToRev6 + " to " + PenRec.AppToRev6 + " and saved.")
  End If

  If TempAppToRev7 <> PenRec.AppToRev7 Then
    MainLog ("Penalty application to revenue #7 has changed from " + TempAppToRev7 + " to " + PenRec.AppToRev7 + " and saved.")
  End If

  If TempAppToRev8 <> PenRec.AppToRev8 Then
    MainLog ("Penalty application to revenue #8 has changed from " + TempAppToRev8 + " to " + PenRec.AppToRev1 + " and saved.")
  End If

End Sub

Private Sub fpcmbMethod_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbMethod.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbMethod.ListIndex = -1
  End If
  If fpcmbMethod.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbCombine.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If


End Sub

Private Sub fpCurrFlatRate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpdblPct.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtPenDesc.SetFocus
  End If
End Sub

Private Sub fptxtPenDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpCurrFlatRate.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If OptPenOpt3.Enabled = True Then
      OptPenOpt3.SetFocus
    ElseIf OptPenOpt2.Enabled = True Then
      OptPenOpt2.SetFocus
    ElseIf OptPenOpt1.Enabled = True Then
      OptPenOpt1.SetFocus
    Else
      chkOptRev3.SetFocus
    End If
  End If
End Sub
