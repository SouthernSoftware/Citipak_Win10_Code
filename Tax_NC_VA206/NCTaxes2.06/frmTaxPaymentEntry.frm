VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxPaymentEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Entry"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbTenderType 
      Height          =   390
      Left            =   3120
      TabIndex        =   3
      Top             =   4875
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   688
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
      ColDesigner     =   "frmTaxPaymentEntry.frx":08CA
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin EditLib.fpLongInteger fpLongAcctNum 
      Height          =   375
      Left            =   2820
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "0"
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
   Begin EditLib.fpText fptxtRevOpt2 
      Height          =   375
      Left            =   5925
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4860
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrAmtOwed 
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483642
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483641
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCash 
      Height          =   492
      Left            =   4536
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1224
      _Version        =   131072
      _ExtentX        =   2159
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":0BF9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCheck 
      Height          =   492
      Left            =   5886
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1212
      _Version        =   131072
      _ExtentX        =   2138
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":0E0C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCharge 
      Height          =   495
      Left            =   7230
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1215
      _Version        =   131072
      _ExtentX        =   2143
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1020
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
      Height          =   375
      Left            =   4500
      TabIndex        =   16
      Top             =   1800
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1235
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDist 
      Height          =   495
      Left            =   8580
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":144F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   492
      Left            =   10050
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1344
      _Version        =   131072
      _ExtentX        =   2371
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1662
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBills 
      Height          =   375
      Left            =   6420
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1876
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   246
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1344
      _Version        =   131072
      _ExtentX        =   2371
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1A8A
   End
   Begin EditLib.fpText fptxtState 
      Height          =   375
      Left            =   1680
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3675
      Width           =   615
      _Version        =   196608
      _ExtentX        =   1085
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   2
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
   Begin EditLib.fpDateTime fptxtPayDate 
      Height          =   345
      Left            =   8640
      TabIndex        =   2
      Tag             =   "The date you enter here will be the date that appears on the 'Payment Entry' screen. The date on that screen is not editable."
      Top             =   1200
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   609
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
      Text            =   "10/03/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
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
   Begin EditLib.fpText fptxtName 
      Height          =   375
      Left            =   1680
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7223
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpText fptxtAddress 
      Height          =   375
      Left            =   1680
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2820
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7223
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpText fptxtCity 
      Height          =   375
      Left            =   1680
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7223
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpMask fptxtZip 
      Height          =   375
      Left            =   4320
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   3675
      Width           =   1455
      _Version        =   196608
      _ExtentX        =   2566
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
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrCashPd 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   5325
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrChkChrgPd 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrTotRecd 
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6795
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrChngDue 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7230
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpText fptxtDescription 
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   7380
      Width           =   5055
      _Version        =   196608
      _ExtentX        =   8916
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
      MaxLength       =   19
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
   Begin EditLib.fpText fptxtRevTax 
      Height          =   375
      Left            =   5930
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "PRINCIPLE"
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
   Begin EditLib.fpText fptxtRevInt 
      Height          =   375
      Left            =   5925
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3075
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "INTEREST"
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
   Begin EditLib.fpText fptxtRecAdvCol 
      Height          =   375
      Left            =   5925
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3525
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "ADV/COLLECT"
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
   Begin EditLib.fpText fptxtRevLateList 
      Height          =   375
      Left            =   5925
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3975
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "LATE LISTING"
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
   Begin EditLib.fpText fptxtRevOpt1 
      Height          =   375
      Left            =   5925
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4425
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpText fptxtRevOpt3 
      Height          =   375
      Left            =   5925
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   5310
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrPrincOwed 
      Height          =   375
      Left            =   8080
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrPrincPaid 
      Height          =   375
      Left            =   9760
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrIntOwed 
      Height          =   375
      Left            =   8085
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3075
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrIntPaid 
      Height          =   375
      Left            =   9765
      TabIndex        =   9
      Top             =   3075
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrAdvColOwed 
      Height          =   375
      Left            =   8085
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3525
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrAdvColPaid 
      Height          =   375
      Left            =   9765
      TabIndex        =   10
      Top             =   3525
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrLateListOwed 
      Height          =   375
      Left            =   8085
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrLateListPaid 
      Height          =   375
      Left            =   9765
      TabIndex        =   11
      Top             =   3975
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrRevOpt1Owed 
      Height          =   375
      Left            =   8085
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4425
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrRevOpt1Paid 
      Height          =   375
      Left            =   9765
      TabIndex        =   12
      Top             =   4425
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrRevOpt2Owed 
      Height          =   375
      Left            =   8085
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4860
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrRevOpt2Paid 
      Height          =   375
      Left            =   9765
      TabIndex        =   13
      Top             =   4860
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrRevOpt3Owed 
      Height          =   375
      Left            =   8085
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5310
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrRevOpt3Paid 
      Height          =   375
      Left            =   9765
      TabIndex        =   14
      Tag             =   "1"
      Top             =   5310
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrTotOwed 
      Height          =   375
      Left            =   8085
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrTotPaid 
      Height          =   375
      Left            =   9765
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin fpBtnAtlLibCtl.fpBtn cmdInfo 
      Height          =   492
      Left            =   3186
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1224
      _Version        =   131072
      _ExtentX        =   2159
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1C9E
   End
   Begin EditLib.fpCurrency fpCurrDisc 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   6200
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrTotWDisc 
      Height          =   375
      Left            =   9765
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrPrePay 
      Height          =   375
      Left            =   9760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDrawer 
      Height          =   492
      Left            =   1716
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1344
      _Version        =   131072
      _ExtentX        =   2371
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
      ButtonDesigner  =   "frmTaxPaymentEntry.frx":1EB1
   End
   Begin VB.Label Label26 
      Caption         =   "Label26"
      Height          =   375
      Left            =   960
      TabIndex        =   78
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      TabIndex        =   77
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCurrTaxYr 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tax Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   75
      Top             =   720
      Width           =   2415
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5880
      X2              =   11400
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   240
      Top             =   1680
      Width           =   11175
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prepay Amt:"
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
      Left            =   8280
      TabIndex        =   74
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5880
      X2              =   11400
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11400
      X2              =   11400
      Y1              =   6960
      Y2              =   2280
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Paid With Discount:"
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
      TabIndex        =   73
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   5880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3120
      X2              =   5280
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   5880
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   2280
      Y2              =   7920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9720
      X2              =   9720
      Y1              =   2280
      Y2              =   6360
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discount:"
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
      Left            =   1920
      TabIndex        =   71
      Top             =   6320
      Width           =   975
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   7920
      TabIndex        =   69
      Top             =   7020
      Width           =   1455
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   9765
      X2              =   11325
      Y1              =   5805
      Y2              =   5805
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   8085
      X2              =   9645
      Y1              =   5805
      Y2              =   5805
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   8040
      X2              =   8040
      Y1              =   1680
      Y2              =   6360
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6120
      TabIndex        =   59
      Top             =   5955
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Amt Paid"
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
      Height          =   255
      Left            =   9880
      TabIndex        =   51
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Amt Owed"
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
      Height          =   255
      Left            =   8200
      TabIndex        =   50
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Revenue"
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
      Height          =   255
      Left            =   6380
      TabIndex        =   49
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   2280
      Y2              =   7920
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
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
      Left            =   480
      TabIndex        =   48
      Top             =   7335
      Width           =   2415
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
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
      Left            =   480
      TabIndex        =   47
      Top             =   6915
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check/Charge Amt Paid:"
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
      Height          =   372
      Left            =   360
      TabIndex        =   46
      Top             =   5868
      Width           =   2532
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Paid:"
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
      Left            =   720
      TabIndex        =   45
      Top             =   5415
      Width           =   2175
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
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
      Left            =   1200
      TabIndex        =   44
      Top             =   4980
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
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
      Left            =   1200
      TabIndex        =   43
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Tax  Billing Payment"
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
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
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
      Left            =   3240
      TabIndex        =   37
      Top             =   3765
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
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
      Left            =   480
      TabIndex        =   36
      Top             =   3765
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
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
      Left            =   480
      TabIndex        =   35
      Top             =   3345
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   480
      TabIndex        =   34
      Top             =   2925
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   480
      TabIndex        =   33
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   7920
      TabIndex        =   32
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source: Tax"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblOperName 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
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
      Left            =   4313
      TabIndex        =   29
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblOperNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Left            =   1320
      TabIndex        =   28
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Acct Number:"
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
      Left            =   285
      TabIndex        =   27
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2310
      Top             =   240
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment  Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   360
      Width           =   4020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   2325
      Top             =   120
      Width           =   7020
   End
End
Attribute VB_Name = "frmTaxPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim FirstBillRec As Long
  Dim BtnFnt As Double
  Public NotFirstLoad As Boolean
  Public TempAcctNum As Long
  Dim DiscXDate As Integer
  Dim ThisDiscAmt As Double
  Dim ThisDiscPct As Double
  Dim CustList() As CustPayListType
  Dim CustListCnt&
  Public EditFlag As Boolean
  Public GetNewCust As Boolean
  Dim ExitFlag As Boolean
  Dim LastPayRec&, CustPayRec&
  Dim TempBillList() As PayListType
  Dim TempBillListCnt As Integer
  Dim TempPayDate As Integer
  Dim TempAmtOwed As Double
  Dim TempTenderTY As String
  Dim TempCashAmt As Double
  Dim TempChkAmt As Double
  Dim TempChrgAmt As Double
  Dim TempAmtRecd As Double
  Dim TempChange As Double
  Dim TempDesc As String
  Dim TempPaidOwed1AmtOwed As Double
  Dim TempPaidOwed2AmtOwed As Double
  Dim TempPaidOwed3AmtOwed As Double
  Dim TempPaidOwed4AmtOwed As Double
  Dim TempPaidOwed5AmtOwed As Double
  Dim TempPaidOwed6AmtOwed As Double
  Dim TempPaidOwed7AmtOwed As Double
  Dim TempPaidOwed1AmtPaid As Double
  Dim TempPaidOwed2AmtPaid As Double
  Dim TempPaidOwed3AmtPaid As Double
  Dim TempPaidOwed4AmtPaid As Double
  Dim TempPaidOwed5AmtPaid As Double
  Dim TempPaidOwed6AmtPaid As Double
  Dim TempPaidOwed7AmtPaid As Double
  Dim TempTotOwed As Double
  Dim TempAmtPaid As Double
  Dim TempTotPaid As Double
  Dim DontExit As Boolean
  Dim CustDiscAmt As Double
  Dim DiscAmtAry() As Double
  Dim DiscRecAry() As Long
  Dim DiscAryCnt As Integer
  Dim StopWarn As Boolean
  Dim WhichRec() As Integer
  Dim DiscCnt As Integer
  Dim InClear As Boolean
  Dim TempPrincPaid As Double
  Dim TempIntPaid As Double
  Dim TempAdvColPaid As Double
  Dim TempLateListPaid As Double
  Dim TempRev1Paid As Double
  Dim TempRev2Paid As Double
  Dim TempRev3Paid As Double
  Dim TempDisc As Double
  Dim TempTotPd As Double
  Dim TempPrePay As Double
  Dim MaxDisc As Double
  Dim InOverRideDist As Boolean
  Dim InSave As Boolean
  Dim OverPay As Boolean
  Dim BillHasFocus As Boolean
  Dim RctValidate As Boolean
  Dim DiscYN As Boolean
  Dim SaveMode As Boolean
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim BegAmount As Double '10/20/06
  Dim DistrFlag As Boolean 'added 10/20/06
  Public Lookup As Boolean '2/14/06
  
  'OpenTempPayFile is the same as open TaxCPRFileName
  'OpenPayListFile is the same as open TaxLOPFileName
Public Sub cmdBills_Click()
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim ThisAmtOwed As Double
  Dim ThisCust As Long
  
  'on error goto ERRORSTUFF
  
  ThisAmtOwed = CDbl(fpCurrAmtOwed.Value)
  If fpLongAcctNum.Value = 0 Then
    frmTaxMsg.Label1.Caption = "Please supply a customer first before accessing bill information."
    frmTaxMsg.Label1.Top = 800
    frmTaxMsg.Show vbModal
    If fpLongAcctNum.Enabled = True Then
      fpLongAcctNum.SetFocus
    End If
    Exit Sub
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  
  If GetCustBalance(GCustNum, -1) = 0 Then
    frmTaxMsg.Label1.Caption = "This customer has a zero balance."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  
  frmTaxBillList.Show vbModal
  DoEvents
  If EditFlag = True Then Exit Sub 'critical

  ThisCust = 0
  If BillCnt > 0 Or Exist(TempBillRecs) Then 'BillCnt is a temporary value representing
  'the number of bills tagged for payment. TempBillRecs is a temporary file that keeps up
  'with the currently tagged bill and serves as a backup to BillCnt
    Call LoadAmtOwed
  Else
    fpCurrDisc = 0 '9/20/05
    Call ResetLeaveName
  End If
  
  GetNewCust = False
  fpcmbTenderType.SetFocus
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdBills_Click", Erl)
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

Private Sub cmdBills_GotFocus()
  BillHasFocus = True
End Sub

Private Sub cmdBills_LostFocus()
  BillHasFocus = False
End Sub

Private Sub cmdCash_Click()
  'on error goto ERRORSTUFF
  
  If fpCurrAmtOwed.DoubleValue > 0 Then
    fpcmbTenderType.Text = "CASH"
    Call ClearPaidFields
    If fpCurrDisc.Value = 0 Then
      fpCurrDisc.Value = 0
      fpCurrCashPd = OldRound(fpCurrAmtOwed.DoubleValue)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue)
    Else
      fpCurrDisc.Value = MaxDisc
      fpCurrCashPd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
    End If
    Call cmdDist_Click
    Call ReFigure
    fptxtDescription.SetFocus
  ElseIf fpCurrAmtOwed.DoubleValue = 0 Then
    Call TaxMsg(900, "Automatic distribution can only take place if there is an amount owed.")
    If fpCurrCashPd.Enabled = True Then
      fpCurrCashPd.SetFocus
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdCash_Click", Erl)
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

Private Sub cmdCharge_Click()
  'on error goto ERRORSTUFF
  
  If fpCurrAmtOwed.DoubleValue > 0 Then
    fpcmbTenderType.Text = "CHARGE"
    Call ClearPaidFields
    If fpCurrDisc.Value = 0 Then
      fpCurrDisc.Value = 0
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue)
    Else
      fpCurrDisc.Value = MaxDisc
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
    End If
    Call cmdDist_Click
    Call ReFigure
    fptxtDescription.SetFocus
  ElseIf fpCurrAmtOwed.DoubleValue = 0 Then
    Call TaxMsg(900, "Automatic distribution can only take place if there is an amount owed.")
    If fpCurrCashPd.Enabled = True Then
      fpCurrCashPd.SetFocus
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdCharge_Click", Erl)
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

Private Sub cmdCheck_Click()
  'on error goto ERRORSTUFF
  
  If fpCurrAmtOwed.DoubleValue > 0 Then
    fpcmbTenderType.Text = "CHECK"
    Call ClearPaidFields
    If fpCurrDisc.Value = 0 Then
      fpCurrDisc.Value = 0
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue)
    Else
      fpCurrDisc.Value = MaxDisc
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
    End If
    Call cmdDist_Click
    Call ReFigure
    fptxtDescription.SetFocus
  ElseIf fpCurrAmtOwed.DoubleValue = 0 Then
    Call TaxMsg(900, "Automatic distribution can only take place if there is an amount owed.")
    If fpCurrCashPd.Enabled = True Then
      fpCurrCashPd.SetFocus
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdCheck_Click", Erl)
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

Private Sub cmdDist_Click()
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim SetUpRec As TaxMasterType
  Dim SHandle As Integer
  Dim x As Integer
  Dim TotRecd As Double
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TransRecord&
  Dim WhatsLeft As Double
  Dim PaidDif As Double
  Dim ThisDif As Double
  Dim TPayRec As PayListType
  Dim PayRec As PayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim y As Integer, z As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim Operator$
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim SmallNum As Integer
  Dim HoldDate As Integer
  Dim HoldNum As Long
  Dim Nextx As Integer
  Dim Thisx As Integer
  
  'on error goto ERRORSTUFF
  
  If CDbl(fpCurrAmtOwed.Value) = 0 And Val(fpLongAcctNum.Text) = 0 Then
    Exit Sub
  End If
  
  If fpCurrCashPd.Value = 0 And fpCurrChkChrgPd.Value = 0 And fpCurrTotPaid.Value = 0 Then
    Call TaxMsg(900, "Please enter an amount paid.")
    If fpcmbTenderType = "CASH" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.BackColor = &H8080FF
        fpCurrCashPd.SetFocus
      End If
    Else
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.BackColor = &H8080FF
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    
    Exit Sub
    DoEvents
  ElseIf fpCurrCashPd.Value = 0 And fpCurrChkChrgPd.Value = 0 And fpCurrTotPaid.Value > 0 Then
    fpCurrPrincPaid.Value = 0
    fpCurrIntPaid.Value = 0
    fpCurrAdvColPaid.Value = 0
    fpCurrLateListPaid.Value = 0
    fpCurrRevOpt1Paid.Value = 0
    fpCurrRevOpt2Paid.Value = 0
    fpCurrRevOpt3Paid.Value = 0
    fpCurrTotPaid.Value = 0
  End If
    
  TotRecd = fpCurrTotRecd.Value
'  WhatsLeft = OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value))
  
  OpenTaxSetUpFile SHandle
  Get SHandle, 1, SetUpRec
  Close SHandle
  
  If BillCnt = 0 And EditFlag = True Then 'user is editing and is not accessing
  'the bill list
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    ThisPrevRec = 0
    NewRec = 0
    Operator$ = CStr(OperNum)
    Operator$ = QPTrim$(Operator$)
    OpenPayListFile PHandle, OperNum 'saved by getting data from temporary
    'bill record
    NumOfPRecs = LOF(PHandle) / Len(PayRec)
    For x = 1 To NumOfPRecs
      Get PHandle, x, PayRec
      If PayRec.CustRec = GCustNum And PayRec.PrevListRec <> -1 Then 'added <> -1 because
      '-1 means that transaction has been deleted 6/29/06
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        BillTrans(BillCnt) = PayRec.BillRec
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillDate(BillCnt) = TempRec.BillDate
      End If
    Next x
    Close PHandle
  ElseIf Exist(TempBillRecs) Then
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    BillCnt = 0
    OpenTempBillRecs TempHandle, NumOfTemps
    For x = 1 To NumOfTemps
      Get TempHandle, x, TempRec
      If TempRec.BillRec > 0 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillTrans(BillCnt) = TempRec.BillPtr
        BillDate(BillCnt) = TempRec.BillDate
        'this data should be the sane data as that where PaySeq() are loaded
      End If
    Next x
    Close TempHandle
  End If
  WhatsLeft = OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value))
  
  Call Distribute(WhatsLeft)
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdDist_Click", Erl)
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

Private Sub cmdDrawer_Click()
  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  Port$ = QPTrim$(RecpPort)
  MainLog ("Oper: " + Str(OperNum) + "Tax Pay-Open Drawer")
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  Print #PortFile, Chr$(7)
  Close PortFile
End Sub

Private Sub cmdExit_Click()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim ThisCust As Integer
  Dim Handle As Integer
  
  'on error goto ERRORSTUFF
  
  Handle = FreeFile
  If CDbl(fpLongAcctNum.Value) <> 0 Then
    If Check4Changes = True Then
      Exit Sub
    End If
  End If
  
  If DontExit = True Then 'DontExit is set in the save routine if after
  'checking for changes is true a problem is caught in the save routine
  'which would then trigger this sub because the save routine was triggered
  'from Check4Changes
    DontExit = False
    Exit Sub
  End If
  
  BillCnt = 0
  PayEntry = False 'a global that tells frmCustomerLookup that this form is
  'where to return when frmCustomerLookup is used
  KillFile TempBillRecs 'TempBillRecs is the filename for the temporary file
  'created when a bill is tagged
  ExitFlag = True
  Close
  KillFile "C:\CPWork\txpyment.dat" 'could be used to identify this form as being opened...
  'currently (4/6/05) not being used
  Call ClearTemps
  GPayNum = 0
  GCustNum = 0
  If Not Exist("C:\CPWork\editpyment.dat") Then
    frmTaxPayMenu.Show
    DoEvents
  Else
    OpenTempPayFile PayHandle, OperNum
    NumOfPRecs = LOF(PayHandle) / Len(PayRec)
    
'    If frmTaxPayEditList.fpListPay.ListCount <> NumOfPRecs Then
      frmTaxPayEditList.fpListPay.Clear
      For x = 1 To NumOfPRecs
        Get PayHandle, x, PayRec
        If PayRec.LastPayRec = 0 Then GoTo MoveOn
        frmTaxPayEditList.fpListPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtOwed))
'        Debug.Print CStr(PayRec.CustAcct)
        If PayRec.CustAcct = fpLongAcctNum.Value Then
          frmTaxPayEditList.fpListPay.ListIndex = x
        End If
        DoEvents
MoveOn:
      Next x
      Close PayHandle
'    End If
    frmTaxPayEditList.fpListPay.Action = ActionForceUpdate
      
    frmTaxPayEditList.Show
    DoEvents
  End If
  
  Close PayHandle
  
  Unload Me
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdExit_Click", Erl)
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

Private Sub cmdInfo_Click()
  If GCustNum = 0 Then
    Exit Sub
  End If
  
  Call frmTaxCustInq.LoadCust
  frmTaxCustInq.Show vbModal

End Sub

Private Sub cmdLookup_Click()
  Dim TaxRec As TaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  PayEntry = True
  Lookup = True '2/14/06
  frmTaxCustLookup.Show
  DoEvents
End Sub

Private Sub cmdSave_Click()
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim TaxPayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPayRecs As Long
  Dim TaxSetupRec As TaxMasterType
  Dim MHandle As Integer
  Dim PrinceOw#
  Dim PrincePD#
  Dim InterestOw#
  Dim InterestPd#
  Dim CollectOw#
  Dim CollectPd#
  Dim TDiscAmt#
  Dim TAmtRecv#
  Dim TAmtPaid#
  Dim ChangeAmt#
  Dim Oper$
  Dim NextListRec&
  Dim PayRecFile As Integer
  Dim NumOfRecs&
  Dim Cnt&
  Dim n As Integer
  Dim Num As Integer
  Dim ThisBal As Double
  Dim SaveFlag As Integer
  Dim Message$
  Dim ThisDisc As Double
  
  'on error goto ERRORSTUFF
  
  SaveMode = True
  SaveFlag = 2
  If GCustNum <= 0 Then
    fpLongAcctNum.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "No customer data has been supplied. No payment can be saved."
    frmTaxMsg.Label1.Top = 800
    frmTaxMsg.Show vbModal
    fpLongAcctNum.SetFocus
    DontExit = True
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If Check4ValidPaidEntries = False Then  '8/12/05'checks to make sure
  'no payments are more that the amounts owed
    Exit Sub
  End If
  
  If EditFlag = True And DistrFlag = False And fpCurrTotPaid.Value > 0 Then 'added 10/20/06
    If fpCurrTotRecd.Value <> BegAmount Then
      If TaxMsgWOpts(900, "Do you wish to distribute these entries?", "F10 Distribute", "ESC Skip") = "continue" Then
        Call ReassignDiscount 'added 7/18/07
        Call cmdDist_Click
      Else
        Call MainLog("User asked if they wanted to distribute new payment amounts and they declined.")
      End If
    End If
  End If
  
  InSave = True
  If AllTaggedPaid = False Then
    If CDbl(fpCurrDisc.Value) > 0 Then
      If CDbl(fpCurrTotOwed.Value) > OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value)) Then
        Message = "This customer cannot receive the discount entered because the bills tagged are not being paid in full. To correct this situation you can eliminate the discount or have the customer pay in full all bills tagged."
        Call TaxMsg(600, Message)
        If fpCurrDisc.Enabled = True Then
          fpCurrDisc.SetFocus
        End If
        InSave = False
        If frmTaxCustLookup.Visible = True Then
          Unload frmTaxCustLookup
        End If
        SaveMode = False
        Exit Sub
      End If
    ElseIf CDbl(fpCurrPrePay.Value) > 0 Then
      If CDbl(fpCurrTotOwed.Value) > 0 Then
        Message = "This customer cannot prepay because all bills tagged for payment are not paid in full."
        Call TaxMsg(800, Message)
        If fpCurrPrePay.Enabled = True Then
          fpCurrPrePay.SetFocus
        End If
        InSave = False
        If frmTaxCustLookup.Visible = True Then
          Unload frmTaxCustLookup
        End If
        SaveMode = False
        Exit Sub
      End If
    End If
  End If
  
  InSave = False
  
  If CDbl(fpCurrPrePay.Value) > 0 And CDbl(fpCurrDisc.Value) > 0 Then
    fpCurrPrePay.BackColor = &H8080FF
    fpCurrDisc.BackColor = &H8080FF
    Call TaxMsg(800, "Overpayment is not allowed when discounts are being used. Please eliminate the prepayment or eliminate the discount amount.")
    If fpCurrPrePay.Enabled = True Then
      fpCurrPrePay.SetFocus
    End If
    SaveMode = False
    Exit Sub
  End If
      
  If CDbl(fpCurrTotPaid.Value) = 0 Then
    fpCurrTotPaid.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "There is a zero value in the 'Total Amount Paid' field. Please make sure an amount paid has been entered and distributed among the 'Revenue Amount Paid' fields."
    frmTaxMsg.Label1.Top = 700
    frmTaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    Else
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If fpCurrTotRecd.Value = 0 Then
    fpCurrTotRecd.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "Please enter an amount paid."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If CDbl(fpCurrAmtOwed.Value) <> CDbl(fpCurrTotOwed.Value) Then
    fpCurrAmtOwed.BackColor = &H8080FF
    fpCurrTotOwed.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "The 'Tax Billing Amount Owed' and the 'Revenue Amount Owed' are not equal. Please make sure you distribute funds received by pressing F9 each time payment data has been updated."
    frmTaxMsg.Label1.Top = 600
    frmTaxMsg.Show vbModal
    Close
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      fpCurrCashPd.SetFocus
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      fpCurrChkChrgPd.SetFocus
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      fpCurrChkChrgPd.SetFocus
    End If
    DontExit = True
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
    
  If CDbl(fpCurrAmtOwed.Value) <= CDbl(fpCurrTotRecd.Value) Then
    If CDbl(fpCurrTotWDisc.Value) <> OldRound((CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value)) - CDbl(fpCurrChngDue.Value)) Then
      fpCurrTotPaid.BackColor = &H8080FF
      fpCurrChngDue.BackColor = &H8080FF
      fpCurrTotRecd.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The 'Total Amount Paid' does not equal the 'Total Amount Received' plus the 'Change Due'. Please make sure the funds distribution is accurate."
      frmTaxMsg.Label1.Top = 700
      frmTaxMsg.Show vbModal
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmTaxCustLookup.Visible = True Then
        Unload frmTaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  ElseIf CDbl(fpCurrAmtOwed.Value) > CDbl(fpCurrTotRecd.Value) And CDbl(fpCurrChngDue.Value) > 0 Then
    fpCurrAmtOwed.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrChngDue.BackColor = &H8080FF
    If CDbl(fpCurrChngDue.Value) <> 0 Then
      frmTaxMsg.Label1.Caption = "ERROR: The amount owed is more than the amount received so no change should be returned. Please re-distribute your data to try to fix this issue."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      Close
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmTaxCustLookup.Visible = True Then
        Unload frmTaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    ElseIf CDbl(fpCurrTotPaid.Value) <> CDbl(fpCurrTotRecd.Value) Then
      fpCurrTotPaid.BackColor = &H8080FF
      fpCurrTotRecd.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The total amount received does not equal the total amount paid. Please re-distribute amounts to try and fix this issue or call Southern Software at 1-800-842-8190 for assistance."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      Close
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmTaxCustLookup.Visible = True Then
        Unload frmTaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
   
  If CDbl(fpCurrTotPaid.Value) > CDbl(fpCurrTotOwed.Value) And (CDbl(fpCurrTotPaid.Value) = CDbl(fpCurrTotRecd.Value)) Then
    fpCurrTotPaid.BackColor = &H8080FF
    fpCurrTotOwed.BackColor = &H8080FF
    frmTaxMsgWOpts.Label1.Caption = "The total amount paid exceeds the total amount owed. Press F10 to continue saving allowing this customer to have a credit balance. Otherwise, press ESC to review."
    frmTaxMsgWOpts.Label1.Top = 700
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      Close
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmTaxCustLookup.Visible = True Then
        Unload frmTaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
  
  If OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrChngDue.Value)) <> CDbl(fpCurrTotRecd.Value) Then
    fpCurrTotPaid.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrChngDue.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "The total amount paid plus the total change amount do not equal the total amount received. Please re-distribute funds to try to correct this issue."
    frmTaxMsg.Label1.Top = 700
    frmTaxMsg.Show vbModal
    Close
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmTaxCustLookup.Visible = True Then 'when the customer lookup is used
    'the user can change the data on this screen by selecting a new customer...
    'this means that a check4changes must be activate from the lookup screen
    'to make sure the customer change will not leave the current customer's data
    'unsaved...if this trap catches a problem then that problem must be
    'dealt with before changing to the new customer...so if we are here by way
    'of that scenarion then the customer lookup needs unloading now
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum) 'reset the customer number to the
    'what it was before the switch was attempted
    SaveMode = False
    Exit Sub
  End If
  ThisDisc = CDbl(fpCurrDisc.Value) '9/22/05
  ThisBal = 0
  If CDbl(fpCurrAmtOwed.Value) = 0 Then
    ThisBal = GetCustBalance(GCustNum, -1)
    If ThisBal > 0 And CDbl(fpCurrPrePay.Value) > 0 Then
      fpCurrAmtOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The total amount owed field is zero. This customer has an outstanding balance of " + QPTrim$(Using$("$###,##0.00", ThisBal)) + ". Customers with outstanding balances must meet these obligations before prepaying."
      frmTaxMsg.Label1.Top = 700
      frmTaxMsg.Show vbModal
      Close
      DontExit = True
      If frmTaxCustLookup.Visible = True Then
        Unload frmTaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
  
  If CDbl(fpCurrAmtOwed.Value) < CDbl(fpCurrTotPaid.Value) Then
    ThisBal = GetCustBalance(GCustNum, -1)
    If ThisBal > CDbl(fpCurrAmtOwed.Value) Then
      fpCurrAmtOwed.BackColor = &H8080FF
      fpCurrTotPaid.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "This customer has an outstanding balance of " + QPTrim$(Using$("$###,##0.00", ThisBal)) + ". Customers with outstanding balances greater than the displayed amount owed, " + QPTrim$(Using$("$###,##0.00", CDbl(fpCurrAmtOwed.Value))) + ", cannot pay more than the displayed amount owed until all prior obligations have been fulfilled."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      Close
      DontExit = True
      If frmTaxCustLookup.Visible = True Then
        Unload frmTaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
  
  If OldRound(CDbl(fpCurrCashPd.Value) + CDbl(fpCurrChkChrgPd.Value)) <> CDbl(fpCurrTotRecd.Value) Then
    fpCurrCashPd.BackColor = &H8080FF
    fpCurrChkChrgPd.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "The total amount received does not equal the total cash paid plus the total check/charges paid. Please correct this situation."
    frmTaxMsg.Label1.Top = 700
    frmTaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If CDbl(fpCurrTotWDisc.Value) > OldRound(CDbl(fpCurrTotRecd.Value) - CDbl(fpCurrChngDue.Value) + CDbl(fpCurrDisc.Value)) Then
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrTotWDisc.BackColor = &H8080FF
    fpCurrChngDue.BackColor = &H8080FF
    fpCurrDisc.BackColor = &H8080FF
    Message = "The total amount received (minus change due) plus discount is less than the total credited. Please increase the amount received or decrease the amount credited."
    Call TaxMsg(700, Message)
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If CheckOverPay = True Then 'this routine looks to see if the customer is trying
  'to overpay one revenue before completely paying all revenues...this is not allowed
    If frmTaxCustLookup.Visible = True Then
      Unload frmTaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If VerifyPayList = False Then
    Call TaxMsg(800, "WARNING: There is a problem with this payment entry. Please delete and/or reenter this payment.")
    MainLog ("Warning issued for a payment for cust # " + CStr(GCustNum) + " telling operator # " + CStr(OperNum) + " to delete this payment.")
    Close
    Exit Sub
  End If
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxSetupRec
  Close MHandle
  
  OpenTempPayFile PayHandle, OperNum
  Num = LOF(PayHandle) / Len(TaxPayRec)
  
  If EditFlag = True Then
    GPayNum = 0
    For n = 1 To Num
    Get PayHandle, n, TaxPayRec
      If TaxPayRec.CustAcct = GCustNum Then
        GPayNum = n
        Exit For
      End If
    Next n
  End If
  
  frmTaxPrintReceipt.Show vbModal
  If frmTaxPrintReceipt.fptxtChoice.Text = "saveonly" Then
    Unload frmTaxPrintReceipt
    SaveFlag = 1
  ElseIf frmTaxPrintReceipt.fptxtChoice.Text = "abort" Then
    Unload frmTaxPrintReceipt
    Close
    SaveMode = False
    Exit Sub
  ElseIf frmTaxPrintReceipt.fptxtChoice.Text <> "both" Then
    Unload frmTaxPrintReceipt
    Close
    frmTaxMsg.Label1.Caption = "Error: There is a problem reading the 'Save' response. Please call Southern Software at 1-800-842-8190."
    SaveMode = False
    Exit Sub
  End If
 
  If EditFlag = True And GPayNum = 0 Then
    frmTaxMsg.Label1.Caption = "ERROR: The program was not able to locate the customer record being edited in the save procedure. Save attempt aborted. Please call Southern Software @ 1-800-842-8190 for assistance."
    frmTaxMsg.Label1.Top = 800
    frmTaxMsg.Show vbModal
    Close
    SaveMode = False
    Exit Sub
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxCustRec
  Close CHandle
  
  If CDbl(fpCurrDisc.Value) = 0 Then Call RemoveDiscount
  
  Call UPDateListOfPayments
  
  TaxPayRec.AmtOwed = fpCurrAmtOwed.Value
  TaxPayRec.AmtPaid = OldRound(CDbl(fpCurrCashPd.Value) + CDbl(fpCurrChkChrgPd.Value))
  TaxPayRec.AmtRecd = CDbl(fpCurrTotRecd.Value)
  TaxPayRec.CashAmt = CDbl(fpCurrCashPd.Value)
  TaxPayRec.Change = CDbl(fpCurrChngDue.Value)
  If fpcmbTenderType.Text = "CHARGE" Then
    TaxPayRec.ChkAmt = 0
    TaxPayRec.ChrgAmt = CDbl(fpCurrChkChrgPd.Value)
  ElseIf fpcmbTenderType.Text = "CHECK" Then
    TaxPayRec.ChkAmt = CDbl(fpCurrChkChrgPd.Value)
    TaxPayRec.ChrgAmt = 0
  ElseIf fpcmbTenderType.Text = "CASH AND CHECK" Then
    TaxPayRec.ChkAmt = CDbl(fpCurrChkChrgPd.Value)
    TaxPayRec.ChrgAmt = 0
  Else
    TaxPayRec.ChkAmt = 0
    TaxPayRec.ChrgAmt = 0
  End If
  TaxPayRec.CustAcct = GCustNum 'CLng(fpLongAcctNum.Value)
  TaxPayRec.CustAddr = QPTrim$(fptxtAddress.Text)
  TaxPayRec.CustName = QPTrim$(fptxtName.Text)
  TaxPayRec.CustPin = TaxCustRec.PIN
  TaxPayRec.Desc = QPTrim$(fptxtDescription.Text)
  TaxPayRec.DiscAmt = CDbl(fpCurrDisc.Value)
  If BillCnt > 0 Or CDbl(fpCurrPrePay.Value) > 0 Then
'    If EditFlag = False Then 'removed 12/8/2008
      TaxPayRec.LastPayRec = LastPayRec
      TaxPayRec.NumPayRec = BillCnt
'    End If
  End If
  TaxPayRec.OperNum = OperNum
  TaxPayRec.PaidOwed(1).AmtOwed = CDbl(fpCurrPrincOwed.Value)
  TaxPayRec.PaidOwed(1).AmtPaid = CDbl(fpCurrPrincPaid.Value)
  TaxPayRec.PaidOwed(2).AmtOwed = CDbl(fpCurrIntOwed.Value)
  TaxPayRec.PaidOwed(2).AmtPaid = CDbl(fpCurrIntPaid.Value)
  TaxPayRec.PaidOwed(3).AmtOwed = CDbl(fpCurrAdvColOwed.Value)
  TaxPayRec.PaidOwed(3).AmtPaid = CDbl(fpCurrAdvColPaid.Value)
  TaxPayRec.PaidOwed(4).AmtOwed = CDbl(fpCurrLateListOwed.Value)
  TaxPayRec.PaidOwed(4).AmtPaid = CDbl(fpCurrLateListPaid.Value)
  TaxPayRec.PaidOwed(5).AmtOwed = CDbl(fpCurrRevOpt1Owed.Value)
  TaxPayRec.PaidOwed(5).AmtPaid = CDbl(fpCurrRevOpt1Paid.Value)
  TaxPayRec.PaidOwed(6).AmtOwed = CDbl(fpCurrRevOpt2Owed.Value)
  TaxPayRec.PaidOwed(6).AmtPaid = CDbl(fpCurrRevOpt2Paid.Value)
  TaxPayRec.PaidOwed(7).AmtOwed = CDbl(fpCurrRevOpt3Owed.Value)
  TaxPayRec.PaidOwed(7).AmtPaid = CDbl(fpCurrRevOpt3Paid.Value)
  TaxPayRec.PayDate = Date2Num(fptxtPayDate.Text)
  TaxPayRec.TenderTY = QPTrim$(fpcmbTenderType.Text)
  TaxPayRec.TotOwed = fpCurrAmtOwed.Value
  TaxPayRec.TotPaid = OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrAdvColPaid))
  TaxPayRec.TotPaid = OldRound(TaxPayRec.TotPaid + CDbl(fpCurrLateListPaid.Value) + CDbl(fpCurrRevOpt1Paid.Value))
  TaxPayRec.TotPaid = OldRound(TaxPayRec.TotPaid + CDbl(fpCurrRevOpt2Paid.Value) + CDbl(fpCurrRevOpt3Paid))
  TaxPayRec.PrePayAmt = CDbl(fpCurrPrePay.Value)
  TaxPayRec.CustPin = TaxCustRec.PIN
  If Not EditFlag Then
    NumOfPayRecs = LOF(PayHandle) / Len(TaxPayRec) + 1
    CustPayRec& = NumOfPayRecs&
    If CustPayRec& = 0 Then CustPayRec& = 1
    GPayNum = CustPayRec&
    Put PayHandle, GPayNum, TaxPayRec
  Else
    Put PayHandle, GPayNum, TaxPayRec
  End If
  
  KillFile TempBillRecs 'get rid of all temporary files and records in
  'preparation for the next customer
  BillCnt = 0
  ReDim BillTrans(0 To 0) As Long
  
  Call LoadTemps 'save new temps in case a new save takes place for the
  'same customer
  If EditFlag = False Then
    If TaxPayRec.PrePayAmt > 0 Then
      MainLog ("New payment of " + QPTrim$(Using$("$###,##0.00", TaxPayRec.TotPaid)) + " for customer#/name " + QPTrim$(Using$("######", TaxPayRec.CustPin)) + "/" + QPTrim$(fptxtName.Text) + " plus overpayment of " + Using$("$##,##0.00", TaxPayRec.PrePayAmt) + " saved successfully.")
    Else
      MainLog ("New payment of " + QPTrim$(Using$("$###,##0.00", TaxPayRec.TotPaid)) + " for customer#/name " + QPTrim$(Using$("######", TaxPayRec.CustPin)) + "/" + QPTrim$(fptxtName.Text) + " saved successfully.")
    End If
  Else
    If TaxPayRec.PrePayAmt > 0 Then
      MainLog ("Edit payment of " + QPTrim$(Using$("$###,##0.00", TaxPayRec.TotPaid)) + " for customer#/name " + QPTrim$(Using$("######", TaxPayRec.CustPin)) + "/" + QPTrim$(fptxtName.Text) + " plus overpayment of " + Using$("$##,##0.00", TaxPayRec.PrePayAmt) + " saved successfully.")
    Else
      MainLog ("Edit payment of " + QPTrim$(Using$("$###,##0.00", TaxPayRec.TotPaid)) + " for customer#/name " + QPTrim$(Using$("######", TaxPayRec.CustPin)) + "/" + QPTrim$(fptxtName.Text) + " saved successfully.")
    End If
  End If
  
  If CLng(fpLongAcctNum.Value) = GCustNum Then
    EditFlag = True
  End If
  
  Close PayHandle
  
  DontExit = False
  Call Savemsg(900, "This tax payment has been saved successfully.")
  
  If SaveFlag = 2 Then
    Call PrintReceipt
    MainLog ("Receipt printed for " + QPTrim$(fptxtName.Text) + ".")
  End If
  
  Call Clearscreen
  TempAcctNum = 0
  DoEvents
  fpLongAcctNum.SetFocus
  SaveMode = False
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "cmdSave_Click", Erl)
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

Private Sub Form_Click()
  Call MakeEmWhite
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call MakeEmWhite ' done to white out any fields that were reddened when
  'errors were flagged
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      If BillHasFocus = True Then Call cmdBills_Click
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%D"
      Call cmdDist_Click
      KeyCode = 0
    Case vbKeyF10:
      If GCustNum <> fpLongAcctNum.Value Then
        KeyCode = 0 'prevents mistakenly pressing F10 (and a crash)
        Exit Sub
      End If
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%C"
      Call cmdCash_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%k"
      Call cmdCheck_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%g"
      Call cmdCharge_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%B"
      Call cmdBills_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%w"
      Call cmdDrawer_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%I"
      Call cmdInfo_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdLookup_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Call MakeEmWhite
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  NotFirstLoad = False
  StopWarn = False
  InClear = False
  InOverRideDist = False
  InSave = False
  SaveMode = False
  Me.HelpContextID = hlpEnterEdit
  Call LoadMe
  Call GetRcpInfo
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call MakeEmWhite
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      PayEntry = False
      KillFile "C:\CPWork\editpyment.dat"
      KillFile "C:\CPWork\txpyment.dat"
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxPaymentEntry.")
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

Private Sub fpcmbTenderType_Change()
  If QPTrim$(fpcmbTenderType.Text) = "" Then
    fpcmbTenderType.Text = "CASH"
    fpCurrChkChrgPd.Enabled = False
    fpCurrChkChrgPd.Value = 0
  End If
  
  If QPTrim$(fpcmbTenderType.Text) = "CASH" Then
    fpCurrChkChrgPd.Enabled = False
    fpCurrChkChrgPd.Value = 0
    fpCurrCashPd.Enabled = True
  ElseIf QPTrim$(fpcmbTenderType.Text) = "CHECK" Then
    fpCurrChkChrgPd.Enabled = True
    fpCurrCashPd.Value = 0
    fpCurrCashPd.Enabled = False
  ElseIf QPTrim$(fpcmbTenderType.Text) = "CHARGE" Then
    fpCurrChkChrgPd.Enabled = True
    fpCurrCashPd.Value = 0
    fpCurrCashPd.Enabled = False
  ElseIf QPTrim$(fpcmbTenderType.Text) = "CASH AND CHECK" Then
    fpCurrChkChrgPd.Enabled = True
    fpCurrCashPd.Enabled = True
  End If
  
End Sub

Private Sub fpcmbTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTenderType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTenderType.ListIndex = -1
  End If
  If fpcmbTenderType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      ElseIf fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
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

Private Sub fpCurrAdvColPaid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  
  If TempAdvColPaid <> CDbl(fpCurrAdvColPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrCashPd_LostFocus()
  Call ReFigure
'  If StopWarn = True Then Exit Sub
'  StopWarn = True
'  If EditFlag = False Then
'    If fpCurrAmtOwed.DoubleValue = 0 Then
'      If fpCurrTotRecd.DoubleValue > 0 Then
'        fpCurrAmtOwed.BackColor = &H8080FF
'        frmTaxMsg.Label1.Caption = "Automatic amount distribution does not take place if the Amount Due equals $0.00. Please make sure the appropriate amounts are entered manually."
'        frmTaxMsg.Label1.Top = 700
'        frmTaxMsg.Show vbModal
'        If fpcmbTenderType = "CASH" Then
'          If fpCurrCashPd.Enabled = True Then
'            fpCurrCashPd.BackColor = &H8080FF
'            fpCurrCashPd.SetFocus
'          End If
'        ElseIf fpcmbTenderType = "CASH AND CHECK" Then
'          If fpCurrCashPd.Enabled = True Then
'            fpCurrCashPd.BackColor = &H8080FF
'            fpCurrChkChrgPd.BackColor = &H8080FF
'            fpCurrCashPd.SetFocus
'          End If
'        Else
'          If fpCurrChkChrgPd.Enabled = True Then
'            fpCurrChkChrgPd.BackColor = &H8080FF
'            fpCurrChkChrgPd.SetFocus
'          End If
'        End If
'        MainLog ("The amount due is zero for this customer but a value has been entered for amount received. The user was warned to make sure the appropriate amounts were manually entered.")
'      End If
'    End If
'  End If

End Sub

Private Sub fpCurrChkChrgPd_LostFocus()
  'on error goto ERRORSTUFF
  
  Call ReFigure
  If StopWarn = True Then Exit Sub
  StopWarn = True
  If EditFlag = False Then
    If fpCurrAmtOwed.DoubleValue = 0 Then
      If fpCurrTotRecd.DoubleValue > 0 Then
        fpCurrAmtOwed.BackColor = &H8080FF
        frmTaxMsg.Label1.Caption = "Automatic amount distribution does not take place if the Amount Due equals $0.00. Please make sure the appropriate amounts are entered manually."
        frmTaxMsg.Label1.Top = 700
        frmTaxMsg.Show vbModal
        If fpcmbTenderType = "CASH" Then
          If fpCurrCashPd.Enabled = True Then
            fpCurrCashPd.BackColor = &H8080FF
            fpCurrCashPd.SetFocus
          End If
        ElseIf fpcmbTenderType = "CASH AND CHECK" Then
          If fpCurrCashPd.Enabled = True Then
            fpCurrCashPd.BackColor = &H8080FF
            fpCurrChkChrgPd.BackColor = &H8080FF
            fpCurrCashPd.SetFocus
          End If
        Else
          If fpCurrChkChrgPd.Enabled = True Then
            fpCurrChkChrgPd.BackColor = &H8080FF
            fpCurrChkChrgPd.SetFocus
          End If
        End If
        MainLog ("The amount due is zero for this customer but a value has been entered for amount received. The user was warned to make sure the appropriate amounts were manually entered.")
      End If
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "fpCurrChkChrgPd_LostFocus", Erl)
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

Private Sub fpCurrDisc_Click(Button As Integer)
  If NotFirstLoad = False Then Exit Sub
  If MaxDisc = 0 Then
    Call TaxMsg(900, "This customer is not eligible for a discount")
    fpCurrDisc.ControlType = ControlTypeReadOnly
  ElseIf MaxDisc > 0 Then
    fpCurrDisc.ControlType = ControlTypeNormal
  End If
End Sub

Private Sub fpCurrDisc_LostFocus()
  Dim ThisAmt As Double
  
  'on error goto ERRORSTUFF
  
  If CDbl(fpCurrTotOwed.Value) = 0 Then Exit Sub
  
  If CDbl(fpCurrDisc.Value) > MaxDisc Then
    Call TaxMsg(800, "The maximum discount allowed for this customer is " + QPTrim$(Using$("$##,##0.00", MaxDisc)) + ". The program will reset the discount to the maximum allowed.")
    fpCurrDisc = MaxDisc
    If CDbl(fpCurrPrincPaid.Value) > 0 Then
      fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincOwed.Value) - MaxDisc)
    End If
  ElseIf CDbl(fpCurrDisc.Value) < MaxDisc Then
    If BillCnt > 0 Then 'added 7/18/07
      Call ReassignDiscount
      Call Distribute(OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value)))
    End If
  End If
  
  fpCurrTotPaid = AddUpPaidCol
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "fpCurrDisc_LostFocus", Erl)
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

Private Sub fpCurrIntPaid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  
  If TempIntPaid <> CDbl(fpCurrIntPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
  
End Sub

Private Sub fpCurrLateListPaid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  
  If TempLateListPaid <> CDbl(fpCurrLateListPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrPrePay_LostFocus()
  'on error goto ERRORSTUFF
  
  If SaveMode = True Then Exit Sub
  If QPTrim$(fptxtName.Text) <> "" Then NotFirstLoad = True 'added 6/1/06
  If CDbl(fpCurrPrincPaid.Value) = 0 Then
    If CDbl(fpCurrPrePay.Value) > 0 And CDbl(fpCurrTotRecd.Value) = 0 Then
      fpCurrPrePay = 0
      Call TaxMsg(900, "No payment has been entered. Prepayment will be reset to zero.")
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      Else
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      Exit Sub
    End If
  End If
    
  If CDbl(fpCurrPrePay.Value) > 0 And CDbl(fpCurrAmtOwed.Value) <> 0 Then
    If CDbl(fpCurrTotWDisc.Value) < CDbl(fpCurrAmtOwed.Value) Then
      fpCurrPrePay.Value = 0
      Call TaxMsg(700, "Prepayment amounts can only be added if the total amounts owed are paid in full. Applying discounts also prevent prepayments.")
      Call ReLoadPaidTemps
    End If
  End If
  Call AddUpPaidCol
  Call LoadTempPayList
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "fpCurrPrePay_LostFocus", Erl)
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

Private Sub fpCurrPrincPaid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
    
'  If CDbl(fpCurrPrincPaid.Value) > CDbl(fpCurrPrincOwed.Value) Then
'    Call TaxMsg(800, "The amount paid for principal cannot be more than the amount owed for principal. Please correct the amount paid for principal.")
'    fpCurrPrincPaid.SetFocus
'    Exit Sub
'  End If
  
  If TempPrincPaid > CDbl(fpCurrPrincPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrRevOpt1Paid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempRev1Paid > CDbl(fpCurrRevOpt1Paid.Value) Then
    Call OverRideDist
  End If
  
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrRevOpt2Paid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempRev2Paid > CDbl(fpCurrRevOpt2Paid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrRevOpt3Paid_LostFocus()
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempRev3Paid <> CDbl(fpCurrRevOpt3Paid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrTotPaid_Change()
  Call ReFigure
End Sub

Private Sub fpCurrTotRecd_Change()
  If EditFlag = True Then '10/20/06
    DistrFlag = False
  End If
End Sub

Private Sub fpLongAcctNum_LostFocus()
  Dim ThisAcctNum As Long
  
  'on error goto ERRORSTUFF
  
  If TempAcctNum = CLng(fpLongAcctNum.Value) Then Exit Sub
  
  If frmTaxPayMenu.Visible = True Then Exit Sub
  If fpLongAcctNum.Value = 0 And BillCnt = 0 Then Exit Sub
  
  If fpLongAcctNum.Value = 0 Then
    Call Clearscreen
    Exit Sub
  End If
  
  If fpLongAcctNum.Value <> GCustNum Then
    StopWarn = False 'used to keep the warning popup regarding the need for
    'manual distribution if the customer owes nothing and prepays
    NotFirstLoad = False
    GetNewCust = True 'used to keep the _Change sub from being used needlessly
  Else
    If fpLongAcctNum.Value = GCustNum Then
      Exit Sub
    End If
  End If
  
  If Check4ValidCustNum(fpLongAcctNum.Value) = False Then
    fpLongAcctNum.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "The customer number is not valid. Please enter a valid customer number."
    frmTaxMsg.Label1.Top = 800
    frmTaxMsg.Show vbModal
    GetNewCust = False
    Call Clearscreen
    fpLongAcctNum.SetFocus
    Exit Sub
  End If

  ThisAcctNum = fpLongAcctNum.Value
  If NotFirstLoad = True Then 'screen is already loaded
    If TempAcctNum = 0 Then
      If CDbl(fpLongAcctNum.Value) = 0 Then
        Exit Sub 'no need to proceed if the acct num is 0 or unchanged
      ElseIf GCustNum = CDbl(fpLongAcctNum.Value) Then
        Exit Sub
      Else
        Call Check4Changes
        Call Clearscreen
        fpLongAcctNum = ThisAcctNum
        Call GetCust
      End If
    Else
      If TempAcctNum > 0 And TempAcctNum <> fpLongAcctNum.Value Then 'new acct num entered
        If Check4Changes = True Then Exit Sub 'stop to allow Check4Changes to handle
        'any changes
        If DontExit = True Then Exit Sub 'DontExit is set in the save routine if a trap caught
        'something
        Call Clearscreen 'All is OK so reset the data on the screen
      End If
      GetNewCust = True
      fpLongAcctNum = ThisAcctNum 'start the loading process for the new number
      Call LostFocusCheck
    End If
  Else 'brand new screen being loaded
    If TempAcctNum > 0 And TempAcctNum <> fpLongAcctNum.Value Then
      If Check4Changes = True Then Exit Sub
      Call Clearscreen
    End If
    GetNewCust = True
    fpLongAcctNum = ThisAcctNum
    Call LostFocusCheck
  End If
  
  If fpLongAcctNum.Value <> 0 Then 'we are loaded now so turn off the load flag
    NotFirstLoad = True
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "fpLongAcctNum_LostFocus", Erl)
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

Public Sub Clearscreen()
  MsgAlertTimer.Enabled = False 'added 6/29/06
  cmdInfo.ForeColor = &H80000012 'added 6/29/06
  DoEvents 'added 6/29/06
  cmdInfo.FontSize = 12 'added 6/29/06
  InClear = True
  NotFirstLoad = False
  Label25.Visible = False
  fpLongAcctNum.Value = 0
'  fptxtPayDate = Date
  fptxtPayDate = PayDate '2/14/06
  fptxtName.Text = ""
  fptxtAddress.Text = ""
  fptxtCity.Text = ""
  fptxtState.Text = ""
  fptxtZip.Text = ""
  fpCurrAmtOwed.Value = 0
  fpcmbTenderType.Text = "CASH"
  fpCurrCashPd.Value = 0
  fpCurrChkChrgPd.Value = 0
  fpCurrTotRecd.Value = 0
  fpCurrChngDue.Value = 0
  fpCurrDisc.Value = 0
  fpCurrPrincOwed.Value = 0
  fpCurrPrincPaid.Value = 0
  fpCurrIntOwed.Value = 0
  fpCurrIntPaid.Value = 0
  fpCurrAdvColOwed.Value = 0
  fpCurrAdvColPaid.Value = 0
  fpCurrLateListOwed.Value = 0
  fpCurrLateListPaid.Value = 0
  fpCurrRevOpt1Owed.Value = 0
  fpCurrRevOpt1Paid.Value = 0
  fpCurrRevOpt2Owed.Value = 0
  fpCurrRevOpt2Paid.Value = 0
  fpCurrRevOpt3Owed.Value = 0
  fpCurrRevOpt3Paid.Value = 0
  fpCurrTotOwed.Value = 0
  fpCurrTotPaid.Value = 0
  fpCurrPrePay.Value = 0
  fpCurrTotWDisc.Value = 0
  fptxtDescription.Text = ""
  NotFirstLoad = False
  ThisDiscAmt = 0
  Call LoadTemps
  Call AssignPaidTemps
  InClear = False
  BillCnt = 0
  GCustNum = 0
  GPayNum = 0
  ReDim TempBillList(1 To 1) As PayListType
  TempBillListCnt = 0
  BegAmount = 0 'added 7/18/07
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  'on error goto ERRORSTUFF
  
  Lookup = False '2/14/06
  
  BillHasFocus = False
  DiscYN = False
  ClearTemps
  OverPay = False
  MaxDisc = 0
  FileName = "C:\CPWork\txpyment.dat" 'used when using the transaction history report
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  ReDim TempBillList(1 To 1) As PayListType
  TempBillListCnt = 0
  
  fptxtPayDate.Text = PayDate
  ThisDiscAmt = 0 'reset this global for new customer
  DiscXDate = TaxMasterRec.DiscXDate
  DiscXDate = Date2Num(fptxtPayDate.Text) 'corrected 9/20/05 ...was 'Date' instead of fptxtPayDate.text
'  DiscXDate = DiscXDate + 1 'remarked 9/20/05
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  lblCurrTaxYr.Caption = "Current Tax Year: " + CStr(TaxMasterRec.TaxYear)
  'to prevent a possible entry error, if there is no description for
  'an optional revenue then the paid amount field is disabled
  fptxtRevOpt1.Text = QPTrim$(TaxMasterRec.OptRev1)
  If QPTrim$(TaxMasterRec.OptRev1) = "" Then
    fpCurrRevOpt1Paid.Enabled = False
    Opt1Desc = "NONE"
  Else
    Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  End If
  fptxtRevOpt2.Text = QPTrim$(TaxMasterRec.OptRev2)
  If QPTrim$(TaxMasterRec.OptRev2) = "" Then
    fpCurrRevOpt2Paid.Enabled = False
    Opt2Desc = "NONE"
  Else
    Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  End If
  fptxtRevOpt3.Text = QPTrim$(TaxMasterRec.OptRev3)
  If QPTrim$(TaxMasterRec.OptRev3) = "" Then
    fpCurrRevOpt3Paid.Enabled = False
    Opt3Desc = "NONE"
  Else
    Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  End If
  
  fpcmbTenderType.Clear
  fpcmbTenderType.Text = "CASH"
  fpcmbTenderType.AddItem "CASH"
  fpcmbTenderType.AddItem "CHECK"
  fpcmbTenderType.AddItem "CASH AND CHECK"
  fpcmbTenderType.AddItem "CHARGE"
  fpCurrChkChrgPd.Enabled = False 'this field is disabled when 'CASH' is
  'displayed on the tender type drop down
  EditFlag = False
  ExitFlag = False
'  OperNum = 4 'used for testing only
  lblOperNum.Caption = "Operator Number: " + CStr(OperNum)
  lblOperName.Caption = "Operator Name: " + PWUser
  ThisDiscAmt = 0
  GetNewCust = False
  NotFirstLoad = True
  If GCustNum > 0 Then 'coming from edit lookup
    If CustHasMsg(GCustNum) Then 'added 6/29/06
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdInfo.ForeColor = &H80000012
    End If
    EditFlag = True 'added 6/29/06
    Call LoadHerUpEdit
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "LoadMe", Erl)
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

Private Function Check4ValidCustNum(ThisCust As Long) As Boolean
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  
  'on error goto ERRORSTUFF
  
  Check4ValidCustNum = True
  
  If fpLongAcctNum.Value = 0 Then
    Check4ValidCustNum = False
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no tax customers saved."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If ThisCust = TaxRec.Acct Then
      If TaxRec.Deleted <> 0 Then
        Check4ValidCustNum = False
      End If
      Exit For
    End If
  Next x

  Close CHandle

  If x > NumOfCRecs Then
    Call Clearscreen
    Check4ValidCustNum = False
  End If
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "Check4ValidCustNum", Erl)
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

End Function

Public Sub GetCust()
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Integer
  Dim Number As Long
  Dim Name$
  Dim Found As Boolean
  
  On Error Resume Next
  If fpLongAcctNum.Value = 0 Then
    fpLongAcctNum.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "Please enter a customer account number."
    frmTaxMsg.Label1.Top = 800
    frmTaxMsg.Show vbModal
    If fpLongAcctNum.Enabled = True Then
      fpLongAcctNum.SetFocus
    End If
    Exit Sub
  End If
  
  Number = fpLongAcctNum.Value
  
  OpenTaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no tax customers saved."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If Number = TaxRec.Acct Then 'match the selected
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
    If QPTrim$(fptxtName.Text) <> "" And GetNewCust = False Then
      Call Clearscreen
    End If
    KillFile TempBillRecs 'get rid of all temporary files and records in
    'preparation for the this customer
    BillCnt = 0
    ReDim BillTrans(0 To 0) As Long
    Call EnterEditChk
  End If
  
  GetNewCust = False
  
End Sub

Public Sub EnterEditChk()
  Dim ONum$
  Dim ThisRec As Integer
  Dim CustNum As Long
  Dim FindStr$
  
'  'on error goto ERRORSTUFF
  
  'in conjunction with BegBalCheck this set of code determines the
  'current status of the customer the user is attempting to bring up
  'on the screen
  ONum = OperNum
  ThisRec = 0
  CustNum = GCustNum
  If CustNum > 0 Then
    If CustHasMsg(GCustNum) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdInfo.ForeColor = &H80000012
    End If
  End If
  
  Select Case BegBalCheck(CustNum, ONum$, ThisRec)
    Case 1 'normal first time transaction for this customer
      EditFlag = False
      TempAcctNum = CustNum
      Call LoadHerUpWOEdit
      FindStr = FindCustInBatchFile(CustNum)
      If FindStr <> "0" Then
        frmTaxInBatchList.ListStr = FindStr
        frmTaxInBatchList.Show vbModal
        MainLog ("User informed this customer, " + CStr(CustNum) + ", is included in the following unposted batch files: " + FindStr + ".")
      End If
      Exit Sub
    Case 2 'edit a transaction that is in progress
      EditFlag = True
      TempAcctNum = CustNum
      GPayNum = ThisRec
      Call LoadHerUpEdit
      NotFirstLoad = True
      If GetCustBalance(GCustNum, -1) < 0 Then
        Call DisablePayFields
      Else
        Call EnablePayFields
      End If
      Exit Sub
    Case 4 'a transaction for this customer is already in progress
    'so abort this attempt
      GCustNum = 0
      TempAcctNum = 0
      Call Clearscreen
      EditFlag = False
      Call LoadMe
      Exit Sub
    Case Else
      frmTaxMsg.Label1.Caption = "Error: This customer's data could not be retrieved."
      frmTaxMsg.Label1.Top = 700
      frmTaxMsg.Show vbModal
      Close
      Exit Sub
  End Select
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "EnterEditChk", Erl)
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

Private Sub LoadHerUpWOEdit()
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim ThisBalance#
  
  'on error goto ERRORSTUFF
0:
  KillFile TempBillRecs
  NotFirstLoad = False
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
1:
  Label25.Visible = False

  DiscYN = False
  ThisBalance = GetCustBalance(GCustNum, -1)
2:
'  Unload frmTaxMsg
  If ThisBalance = 0 Then
    frmTaxMsg.Label1.Caption = "This customer has a zero balance."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Call DisablePayFields
3:  ElseIf ThisBalance < 0 Then
    Call TaxMsg(900, "This customer has a balance of -" + QPTrim$(Using$("$##,##0.00", Abs(ThisBalance))) + ".")
    Label25.Visible = True
    Label25.Caption = "This customer has a balance of -" + QPTrim$(Using$("$##,##0.00", Abs(ThisBalance))) + "."
    Call DisablePayFields
4:
  Else
    Call EnablePayFields
  End If
5:  TempAcctNum = TaxRec.Acct
  fptxtName.Text = QPTrim$(TaxRec.CustName)
  If QPTrim$(TaxRec.Addr1) <> "" Then
    fptxtAddress.Text = QPTrim$(TaxRec.Addr1)
  Else
    fptxtAddress.Text = QPTrim$(TaxRec.Addr2)
  End If

  fptxtCity.Text = QPTrim$(TaxRec.City)
  fptxtState.Text = QPTrim$(TaxRec.State)
  fptxtZip = QPTrim$(TaxRec.Zip)
  ThisDiscAmt = 0
  GetNewCust = False
  Call LoadTemps
  EditFlag = False
'  NotFirstLoad = True
  cmdBills.Enabled = True
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "LoadHerUpWOEdit", Erl)
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

Private Sub LostFocusCheck()
  'on error goto ERRORSTUFF
  
  If fpLongAcctNum.Value = 0 Then
    Call Clearscreen
    Exit Sub
  End If
  
  
  Call GetCust
  
  If GCustNum = 0 Then
    Call Clearscreen
    Exit Sub
  End If
  
  If EditFlag = False And GCustNum <> fpLongAcctNum.Value Then
    Call LoadHerUpWOEdit
  End If
  
  NotFirstLoad = True
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "LostFocusCheck", Erl)
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

Private Sub fpCurrCashPd_Change()
  Dim TotAmt As Double
  Dim Cash As Double
  Dim Chrg As Double
  Dim Disc As Double
  
  Disc = CDbl(fpCurrDisc.Value)
  Cash = CDbl(fpCurrCashPd.Value)
  Chrg = CDbl(fpCurrChkChrgPd.Value)
  TotAmt = OldRound(Cash + Chrg)
  fpCurrTotRecd.Value = TotAmt
End Sub

Private Sub fpCurrChkChrgPd_Change()
  Dim TotAmt As Double
  Dim Cash As Double
  Dim Chrg As Double
  
  Cash = CDbl(fpCurrCashPd.Value)
  Chrg = CDbl(fpCurrChkChrgPd.Value)
  TotAmt = OldRound(Cash + Chrg)
  fpCurrTotRecd.Value = TotAmt

End Sub

Private Sub LoadAmtOwed()
  Dim x As Integer
  Dim TaxOwed#
  Dim IntOwed#
  Dim ColOwed#
  Dim LLOwed#
  Dim RevOpt1#
  Dim RevOpt2#
  Dim RevOpt3#
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim MasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim ThisTaxYear As Integer
  Dim Message$
  Dim ThisBal As Double
  Dim DiscCheck As Integer
  Dim Dif As Double
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, MasterRec
  Close MHandle
  ThisTaxYear = MasterRec.TaxYear
  
  ThisDiscPct = MasterRec.DisPct
  
  TaxOwed# = 0
  IntOwed# = 0
  ColOwed# = 0
  LLOwed# = 0
  RevOpt1# = 0
  RevOpt2# = 0
  RevOpt3# = 0
  OpenTaxTransFile THandle, NumOfTRecs
  
  For x = 1 To BillCnt
    Get THandle, BillTrans(x), TransRec
      TaxOwed# = OldRound(TaxOwed# + TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2 + TransRec.Revenue.Principle3 + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5 + TransRec.Revenue.Penalty)
      TaxOwed# = OldRound(TaxOwed# - (TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd + TransRec.Revenue.Principle3Pd + TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd + TransRec.Revenue.PenaltyPd))
      TaxOwed# = OldRound(TaxOwed# - TransRec.DiscAmt) '8/23/07
      IntOwed# = OldRound(IntOwed# + TransRec.Revenue.Interest)
      IntOwed# = OldRound(IntOwed# - TransRec.Revenue.InterestPd)
      ColOwed# = OldRound(ColOwed# + TransRec.Revenue.Collection)
      ColOwed# = OldRound(ColOwed# - TransRec.Revenue.CollectionPd)
      LLOwed# = OldRound(LLOwed# + TransRec.Revenue.LateList)
      LLOwed# = OldRound(LLOwed# - TransRec.Revenue.LateListPd)
      RevOpt1# = OldRound(RevOpt1# + TransRec.Revenue.RevOpt1)
      RevOpt1# = OldRound(RevOpt1# - TransRec.Revenue.RevOpt1Pd)
      RevOpt2# = OldRound(RevOpt2# + TransRec.Revenue.RevOpt2)
      RevOpt2# = OldRound(RevOpt2# - TransRec.Revenue.RevOpt2Pd)
      RevOpt3# = OldRound(RevOpt3# + TransRec.Revenue.RevOpt3)
      RevOpt3# = OldRound(RevOpt3# - TransRec.Revenue.RevOpt3Pd)
  Next x

  fpCurrPrincOwed = TaxOwed#
  fpCurrIntOwed = IntOwed#
  fpCurrAdvColOwed = ColOwed#
  fpCurrLateListOwed = LLOwed#
  fpCurrRevOpt1Owed = RevOpt1#
  fpCurrRevOpt2Owed = RevOpt2#
  fpCurrRevOpt3Owed = RevOpt3#
  
  fpCurrTotOwed = OldRound(TaxOwed# + IntOwed# + ColOwed# + LLOwed# + RevOpt1# + RevOpt2# + RevOpt3#)
  fpCurrAmtOwed = OldRound(TaxOwed# + IntOwed# + ColOwed# + LLOwed# + RevOpt1# + RevOpt2# + RevOpt3#)
  
  Close THandle
  
  MaxDisc = 0
  Call GetMaxDisc
  If MaxDisc > 0 Then
    If GetCustBalance(GCustNum, ThisTaxYear) > 0 Then
      Message = "This customer is eligible for a maximum discount of " + QPTrim$(Using$("$##,##0.00", MaxDisc)) + " but still owes money for past tax bills. If you wish to apply the discount anyway then press F10. Otherwise, press ESC to override the discount."
      If TaxMsgWOpts(600, Message, "F10 Discount OK", "ESC NO Discount") = "abort" Then
        Unload frmTaxMsgWOpts
        Call RemoveDiscount
      Else
        Unload frmTaxMsgWOpts
        fpCurrDisc = MaxDisc
        Dif = OldRound(CDbl(fpCurrTotOwed.Value) - MaxDisc)
        Call TaxMsg(900, "The total amount owed including the discount will be " + QPTrim$(Using$("$##,##0.00", Dif)) + ".")
      End If
    Else
      Message = "This customer is eligible for a maximum discount of " + QPTrim$(Using$("$##,##0.00", MaxDisc)) + ". If you wish to apply this discount then press F10. Otherwise, press ESC to override the discount."
      If TaxMsgWOpts(700, Message, "F10 Discount OK", "ESC NO Discount") = "abort" Then
        Unload frmTaxMsgWOpts
        Call RemoveDiscount
      Else
        Unload frmTaxMsgWOpts
        fpCurrDisc = MaxDisc
        Dif = OldRound(CDbl(fpCurrTotOwed.Value) - MaxDisc)
        Call TaxMsg(900, "The total amount owed including the discount will be " + QPTrim$(Using$("$##,##0.00", Dif)) + ".")
      End If
    End If
  Else
    Call RemoveDiscount
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "LoadAmtOwed", Erl)
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

Private Function Check4Discounts() As Integer
  Dim TaxSURec As TaxMasterType
  Dim MHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim Balance#
  Dim x As Integer
  Dim TaxYear As Integer
  Dim ListRec As PayListType
  Dim ListHandle As Integer
  Dim NumOfLRecs As Integer
  Dim Operator$
  Dim PayTally As Double
  Dim DiscOK As Integer
  
  'on error goto ERRORSTUFF
  
  Check4Discounts = 0
  If DiscXDate < Date2Num(fptxtPayDate.Text) Then
    fpCurrDisc.Value = 0
    Exit Function
  ElseIf DiscXDate = 0 Then
    fpCurrDisc.Value = 0
    Exit Function
  End If
  
  ReDim DiscAmtAry(1 To 1) As Double
  ReDim DiscRecAry(1 To 1) As Long
  DiscAryCnt = 0
  
  Operator$ = CStr(OperNum)
  
  ReDim WhichRec(1 To 1) As Integer
  DiscCnt = 0
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxSURec
  Close MHandle
  
  If TaxSURec.DisPct = 0 Then
    Exit Function
  End If
  TaxYear = TaxSURec.TaxYear
  
  ThisDiscAmt = 0
  PayTally = 0
  DiscOK = 0 '0 - no discounts allowed  1 - discounts allowed no warnings needed
  '2 - discounts can be allowed but...
  For x = 1 To BillCnt
    If TempBillList(x).TaxYear = TaxYear Then
      PayTally = OldRound(TempBillList(x).DiscAmt + TempBillList(x).Interest1 + TempBillList(x).Collection + TempBillList(x).LateList + TempBillList(x).Principle1 + TempBillList(x).OptRev1 + TempBillList(x).OptRev2 + TempBillList(x).OptRev3)
      If TempBillList(x).TotOwed > PayTally Then  'no discount if nothing is paid for this tax year
        Exit Function
      End If
    End If
  Next x
  
  PayTally = 0
  For x = 1 To BillCnt
    PayTally = OldRound(PayTally + TempBillList(x).TotPaid)
  Next x
  If PayTally = CDbl(fpCurrTotOwed.Value) Then
    If GetCustBalance(GCustNum, TaxYear) > PayTally Then 'if true than old balance exists
      DiscOK = 2 'discount allowed but warn that some old balance is outstanding
    Else
      DiscOK = 1 'discount allowed...no warnings necessary
    End If
  End If
  
  If InSave = True And CDbl(fpCurrTotWDisc.Value) < CDbl(fpCurrTotOwed.Value) Then
    If GetCustBalance(GCustNum, TaxYear) > 0 Then 'if true than old balance exists
      DiscOK = 2 'discount allowed but warn that some old balance is outstanding
    Else
      DiscOK = 1 'discount allowed...no warnings necessary
    End If
  End If
  
DoOver1:
  OpenTaxTransFile THandle, NumOfTRecs
  If BillCnt > 0 Then
    For x = 1 To BillCnt
      Get THandle, BillTrans(x), TaxTrans
      TempBillList(x).DiscAmt = 0
      TempBillList(x).DiscXDate = 0
'      TaxTrans.DiscXDate = DiscXDate 'for testing only
      If TaxTrans.DiscXDate > Date2Num(fptxtPayDate.Text) And TaxTrans.TaxYear = TaxYear Then
'      If TaxTrans.DiscXDate > Date2Num(fptxtPayDate.Text) And CInt(Mid(MakeRegDate(TaxTrans.TransDate), 7, 4)) = TAXYEAR Then
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        If Balance# > 0 Then 'save which transaction the discount is applied to
          DiscCnt = DiscCnt + 1
          ReDim Preserve WhichRec(1 To DiscCnt) As Integer
          WhichRec(DiscCnt) = x
          ThisDiscAmt = ThisDiscAmt + OldRound(Balance# * (TaxSURec.DisPct * 0.01))
          MaxDisc = ThisDiscAmt
          TempBillList(x).DiscAmt = ThisDiscAmt
          TempBillList(x).DiscXDate = TaxTrans.DiscXDate
          DiscAryCnt = DiscAryCnt + 1
          ReDim Preserve DiscAmtAry(1 To DiscAryCnt) As Double
          DiscAmtAry(DiscAryCnt) = ThisDiscAmt
          ReDim Preserve DiscRecAry(1 To DiscAryCnt) As Long
          DiscRecAry(DiscAryCnt) = BillTrans(x)
        End If
      End If
    Next x
    
    ThisDiscPct = TaxSURec.DisPct 'assign to global
    If ThisDiscAmt > 0 Then
      Check4Discounts = DiscOK
    End If
    
    Close THandle
    
 End If
 
 Exit Function
 
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "Check4Discounts", Erl)
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
 
End Function

Private Function Unique$(Path$)
  Dim TempName$
  Dim Seed&
  
  If Len(Path$) And Right$(Path$, 1) <> "\" Then Path$ = Path$ + "\"
  Seed& = Abs(Timer1)            'use the TIMER as a seed
  Do
    TempName$ = Path$ + Mid$(Str$(Seed&), 2)    'make a string out of it
    TempName$ = TempName$ + ".RPT"
    Seed& = Seed& + 1           'increment for next time
  Loop Until Not Exist(TempName$)              'loop and try another name
  Unique$ = TempName$           'this is the function output
  
End Function

Private Sub LoadHerUpEdit()
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  
  'on error goto ERRORSTUFF
  
  DistrFlag = False 'added 10/20/06
  KillFile TempBillRecs
  Label25.Visible = False
  NotFirstLoad = False
  BillCnt = 0
  OpenTempPayFile PayHandle, OperNum
  Get PayHandle, GPayNum, PayRec
  Close PayHandle
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  
  TempAcctNum = TaxRec.Acct
  fpLongAcctNum = TaxRec.Acct
  fptxtName.Text = QPTrim$(TaxRec.CustName)
  If QPTrim$(TaxRec.Addr1) <> "" Then
    fptxtAddress.Text = QPTrim$(TaxRec.Addr1)
  Else
    fptxtAddress.Text = QPTrim$(TaxRec.Addr2)
  End If
  fptxtCity.Text = QPTrim$(TaxRec.City)
  fptxtState.Text = QPTrim$(TaxRec.State)
  fptxtZip = QPTrim$(TaxRec.Zip)
  
  fpCurrDisc = PayRec.DiscAmt
  fpCurrPrincOwed = PayRec.PaidOwed(1).AmtOwed
  fpCurrPrincPaid = PayRec.PaidOwed(1).AmtPaid
  fpCurrIntOwed = PayRec.PaidOwed(2).AmtOwed
  fpCurrIntPaid = PayRec.PaidOwed(2).AmtPaid
  fpCurrAdvColOwed = PayRec.PaidOwed(3).AmtOwed
  fpCurrAdvColPaid = PayRec.PaidOwed(3).AmtPaid
  fpCurrLateListOwed = PayRec.PaidOwed(4).AmtOwed
  fpCurrLateListPaid = PayRec.PaidOwed(4).AmtPaid
  fpCurrRevOpt1Owed = PayRec.PaidOwed(5).AmtOwed
  fpCurrRevOpt1Paid = PayRec.PaidOwed(5).AmtPaid
  fpCurrRevOpt2Owed = PayRec.PaidOwed(6).AmtOwed
  fpCurrRevOpt2Paid = PayRec.PaidOwed(6).AmtPaid
  fpCurrRevOpt3Owed = PayRec.PaidOwed(7).AmtOwed
  fpCurrRevOpt3Paid = PayRec.PaidOwed(7).AmtPaid
  fpCurrTotOwed = PayRec.TotOwed
  fpCurrAmtOwed = PayRec.AmtOwed
  fpCurrCashPd = PayRec.CashAmt
  If PayRec.ChkAmt > 0 Then
    fpCurrChkChrgPd = PayRec.ChkAmt
  ElseIf PayRec.ChrgAmt > 0 Then
    fpCurrChkChrgPd = PayRec.ChrgAmt
  Else
    fpCurrChkChrgPd = 0
  End If
  fpcmbTenderType.Text = QPTrim$(PayRec.TenderTY)
  fptxtPayDate.Text = MakeRegDate(PayRec.PayDate)
  fpCurrTotRecd = PayRec.AmtRecd
  fpCurrChngDue = PayRec.Change
  fptxtDescription.Text = QPTrim$(PayRec.Desc)
  fpCurrTotPaid = OldRound(PayRec.TotPaid + PayRec.PrePayAmt)
  fpCurrPrePay = PayRec.PrePayAmt
  fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + PayRec.DiscAmt)
  GetNewCust = False
  Call GetMaxDisc
  Call LoadTempPayList
  Call LoadTemps
  Call AssignPaidTemps
  If OverPay = False Then
    Call EnablePayFields
  End If
  EditFlag = True
'  cmdBills.Enabled = False
  BegAmount = fpCurrTotRecd.Value 'added 7/18/07
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "LoadHerUpEdit", Erl)
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


Private Sub UPDateListOfPayments()
  'Keeps up with which tagged bills go with which customer
  'If the bill list is not accessed then this sub is not used
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim Operator$
  Dim TPayRec As PayListType
  Dim PayRec As PayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer, y As Integer, z As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim TotPaid#
  Dim PrevCnt As Integer
  Dim FoundCnt As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim ThisPrePay As Double
  Dim Nextx As Integer
  
  'on error goto ERRORSTUFF
  
  ThisPrePay = CDbl(fpCurrPrePay.Value)
  ThisPrevRec = 0
  NewRec = 0
  
  Operator$ = CStr(OperNum)
  Operator$ = QPTrim$(Operator$)
  OpenPayListFile PHandle, OperNum
  NumOfPRecs = LOF(PHandle) / Len(PayRec)
  
  LastPayRec = 0
  
  If BillCnt = 0 And GetCustBalance(GCustNum, -1) <= 0 Then 'customer
  'owes nothing and wants to prepay
    OpenTaxSetUpFile MHandle
    Get MHandle, 1, TaxMasterRec
    Close MHandle
    
    TotPaid# = 0
    PayRec.BillRec = -GCustNum
    PayRec.CustRec = GCustNum
    PayRec.PrevListRec = 0
    'the following should always be zero
    PayRec.Principle1 = CDbl(fpCurrPrincPaid.Value)
    TotPaid# = CDbl(fpCurrPrincPaid.Value)
    PayRec.Interest1 = CDbl(fpCurrIntPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrIntPaid.Value))
    PayRec.Collection = CDbl(fpCurrAdvColPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrAdvColPaid.Value))
    PayRec.LateList = CDbl(fpCurrLateListPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrLateListPaid.Value))
    PayRec.OptRev1 = CDbl(fpCurrRevOpt1Paid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrRevOpt1Paid.Value))
    PayRec.OptRev2 = CDbl(fpCurrRevOpt2Paid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrRevOpt2Paid.Value))
    PayRec.OptRev3 = CDbl(fpCurrRevOpt3Paid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrRevOpt3Paid.Value))
    PayRec.DiscAmt = CDbl(fpCurrDisc.Value) 'should be zero always
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrDisc.Value))
    PayRec.TaxYear = TaxMasterRec.TaxYear
    PayRec.Description = QPTrim$(fptxtDescription.Text)
    'ThisPrePay will have a value
    PayRec.PrePayAmt = ThisPrePay
    ThisPrePay = 0
    NumOfPRecs = NumOfPRecs + 1
    LastPayRec = NumOfPRecs
    PayRec.PrevListRec = 0
    Put PHandle, NumOfPRecs, PayRec
    Close PHandle
    Exit Sub
  End If
    
  If NumOfPRecs = 0 Then 'first record saved in the bills list
    For y = 1 To BillCnt
      If TempBillList(y).BillRec > 0 Then
      'this is saving the totals of all bills tagged
        TotPaid# = 0
        PayRec.BillRec = TempBillList(y).BillRec  'TempRec.BillPtr
        PayRec.CustRec = GCustNum
        PayRec.PrevListRec = y - 1
        PayRec.Principle1 = TempBillList(y).Principle1  'CDbl(fpCurrPrincPaid.Value)
        TotPaid# = PayRec.Principle1
        PayRec.Interest1 = TempBillList(y).Interest1  ' CDbl(fpCurrIntPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.Interest1)
        PayRec.Collection = TempBillList(y).Collection  'CDbl(fpCurrAdvColPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.Collection)
        PayRec.LateList = TempBillList(y).LateList  'CDbl(fpCurrLateListPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.LateList)
        PayRec.OptRev1 = TempBillList(y).OptRev1  'CDbl(fpCurrRevOpt1Paid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.OptRev1)
        PayRec.OptRev2 = TempBillList(y).OptRev2  ' CDbl(fpCurrRevOpt2Paid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.OptRev2)
        PayRec.OptRev3 = TempBillList(y).OptRev3  'CDbl(fpCurrRevOpt3Paid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.OptRev3)
        PayRec.TotPaid = TotPaid#  'CDbl(fpCurrTotPaid.Value)
        PayRec.DiscAmt = TempBillList(y).DiscAmt  'CDbl(fpCurrDisc.Value)
        PayRec.TaxYear = TempBillList(y).TaxYear
        PayRec.Description = QPTrim$(fptxtDescription.Text)
        PayRec.PrePayAmt = ThisPrePay 'prepay applies once
        ThisPrePay = 0
        NewRec = NewRec + 1
        LastPayRec = NewRec
        Put PHandle, NewRec, PayRec
      End If
    Next y
    Close PHandle
    Exit Sub
  End If
  
  If BillCnt > 0 And EditFlag = False Then 'bill list has been processed
  'for a new customer
    For y = 1 To BillCnt 'NumOfTemps 'NumOfTemps is how many bill records there are
    'fpr thios customer after having accessed the bill tag screen
      If TempBillList(y).BillRec > 0 Then
        PayRec.BillRec = TempBillList(y).BillRec
        PayRec.CustRec = GCustNum
        PayRec.Principle1 = TempBillList(y).Principle1
        TotPaid# = PayRec.Principle1
        PayRec.Interest1 = TempBillList(y).Interest1
        TotPaid# = OldRound(TotPaid# + PayRec.Interest1)
        PayRec.Collection = TempBillList(y).Collection
        TotPaid# = OldRound(TotPaid# + PayRec.Collection)
        PayRec.LateList = TempBillList(y).LateList
        TotPaid# = OldRound(TotPaid# + PayRec.LateList)
        PayRec.OptRev1 = TempBillList(y).OptRev1
        TotPaid# = OldRound(TotPaid# + PayRec.OptRev1)
        PayRec.OptRev2 = TempBillList(y).OptRev2
        TotPaid# = OldRound(TotPaid# + PayRec.OptRev2)
        PayRec.OptRev3 = TempBillList(y).OptRev3
        TotPaid# = OldRound(TotPaid# + PayRec.OptRev3)
        PayRec.TotPaid = TotPaid#
        PayRec.DiscAmt = TempBillList(y).DiscAmt
        PayRec.TaxYear = TempBillList(y).TaxYear
        PayRec.Description = QPTrim$(fptxtDescription.Text)
        PayRec.PrePayAmt = ThisPrePay
        ThisPrePay = 0
        NumOfPRecs = NumOfPRecs + 1
        LastPayRec = NumOfPRecs
        If y = 1 Then
          PayRec.PrevListRec = 0
        Else
          PayRec.PrevListRec = NumOfPRecs - 1
        End If
        Put PHandle, NumOfPRecs, PayRec
      End If
    Next y
    Close PHandle
    Exit Sub
  End If
   
  FoundCnt = 0
  For x = 1 To NumOfPRecs 'number of bills tagged
    Get PHandle, x, PayRec
    If PayRec.CustRec = GCustNum Then
      For y = 1 To BillCnt
        If PayRec.BillRec = TempBillList(y).BillRec And TempBillList(y).BillRec > 0 Then
          FoundCnt = FoundCnt + 1
          PayRec.Principle1 = TempBillList(y).Principle1
          TotPaid# = PayRec.Principle1
          PayRec.Interest1 = TempBillList(y).Interest1
          TotPaid# = OldRound(TotPaid# + PayRec.Interest1)
          PayRec.Collection = TempBillList(y).Collection
          TotPaid# = OldRound(TotPaid# + PayRec.Collection)
          PayRec.LateList = TempBillList(y).LateList
          TotPaid# = OldRound(TotPaid# + PayRec.LateList)
          PayRec.OptRev1 = TempBillList(y).OptRev1
          TotPaid# = OldRound(TotPaid# + PayRec.OptRev1)
          PayRec.OptRev2 = TempBillList(y).OptRev2
          TotPaid# = OldRound(TotPaid# + PayRec.OptRev2)
          PayRec.OptRev3 = TempBillList(y).OptRev3
          TotPaid# = OldRound(TotPaid# + PayRec.OptRev3)
          PayRec.TotPaid = TotPaid#
          PayRec.DiscAmt = TempBillList(y).DiscAmt
          PayRec.BillRec = TempBillList(y).BillRec
          PayRec.TaxYear = TempBillList(y).TaxYear
          PayRec.Description = QPTrim$(fptxtDescription.Text)
          PayRec.PrePayAmt = ThisPrePay
          ThisPrePay = 0
          LastPayRec = x
          Put PHandle, x, PayRec
        ElseIf PayRec.BillRec = TempBillList(y).BillRec And TempBillList(y).BillRec < 0 Then
          PayRec.PrePayAmt = ThisPrePay 'added this elseif on 8/23/07
          ThisPrePay = 0
          LastPayRec = x
          Put PHandle, x, PayRec
        End If
      Next y
    End If
  Next x
  
  Nextx = 1
  If FoundCnt < BillCnt Then
    For z = NumOfPRecs To 1 Step -1
      Get PHandle, z, PayRec
      If PayRec.CustRec = GCustNum Then
        ThisPrevRec = z
      End If
    Next z
    
    For y = 1 To BillCnt 'take one bill at a time
    'if the bill rec #s match then they were already saved above
    'if they don't match then we have a new billrec and it will
    'be saved below
      For x = 1 To NumOfPRecs
        Get PHandle, x, PayRec
        If PayRec.BillRec = TempBillList(y).BillRec Then
          GoTo NextOne
        End If
      Next x
      
      PrevCnt = 0
      PayRec.CustRec = GCustNum
      PayRec.Principle1 = TempBillList(y).Principle1
      TotPaid# = PayRec.Principle1
      PayRec.Interest1 = TempBillList(y).Interest1
      TotPaid# = OldRound(TotPaid# + PayRec.Interest1)
      PayRec.Collection = TempBillList(y).Collection
      TotPaid# = OldRound(TotPaid# + PayRec.Collection)
      PayRec.LateList = TempBillList(y).LateList
      TotPaid# = OldRound(TotPaid# + PayRec.LateList)
      PayRec.OptRev1 = TempBillList(y).OptRev1
      TotPaid# = OldRound(TotPaid# + PayRec.OptRev1)
      PayRec.OptRev2 = TempBillList(y).OptRev2
      TotPaid# = OldRound(TotPaid# + PayRec.OptRev2)
      PayRec.OptRev3 = TempBillList(y).OptRev3
      TotPaid# = OldRound(TotPaid# + PayRec.OptRev3)
      PayRec.TotPaid = TotPaid#
      PayRec.DiscAmt = TempBillList(y).DiscAmt
      PayRec.BillRec = TempBillList(y).BillRec
      PayRec.TaxYear = TempBillList(y).TaxYear
      PayRec.PrePayAmt = ThisPrePay
      ThisPrePay = 0
      NumOfPRecs = NumOfPRecs + 1
      LastPayRec = NumOfPRecs
      PayRec.PrevListRec = ThisPrevRec
      Put PHandle, NumOfPRecs, PayRec
      ThisPrevRec = NumOfPRecs
NextOne:
    Next y
  End If
  Close PHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "UPDateListOfPayments", Erl)
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
  
Private Function AddUpPaidCol() As Double
  Dim ThisAdd As Double
  Dim MatchAdd As Double
  Dim Message$
  
  'on error goto ERRORSTUFF
  
  MatchAdd = CDbl(fpCurrTotPaid.Value)
  If NotFirstLoad = False Then Exit Function
  AddUpPaidCol = 0
  ThisAdd = OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrAdvColPaid.Value))
  ThisAdd = OldRound(ThisAdd + CDbl(fpCurrLateListPaid.Value) + CDbl(fpCurrRevOpt1Paid.Value))
  ThisAdd = OldRound(ThisAdd + CDbl(fpCurrRevOpt2Paid.Value) + CDbl(fpCurrRevOpt3Paid.Value) + CDbl(fpCurrPrePay.Value))
  If ThisAdd = MatchAdd Then
    AddUpPaidCol = ThisAdd
    fpCurrTotPaid = ThisAdd
    fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
    If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
    fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))
    Exit Function
  End If
  
  fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
  If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
  If ThisAdd > OldRound(CDbl(fpCurrTotRecd.Value) - fpCurrChngDue.Value) Then
'    fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
    fpCurrTotPaid = ThisAdd
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrTotPaid.BackColor = &H8080FF
    frmTaxMsg.Label1.Caption = "The amount distributed exceeds the amount received. Please re-distribute the amounts paid or add funds to the 'Cash Amount Paid' or the 'Check/Charge Amount Paid' fields."
    frmTaxMsg.Label1.Top = 700
    frmTaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    Else
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
  End If
  AddUpPaidCol = ThisAdd
  fpCurrTotPaid = ThisAdd
'  fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
'  If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
  fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "AddUpPaidCol", Erl)
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

End Function

Private Sub MakeEmWhite()
  fpLongAcctNum.BackColor = &H80000005
  fpCurrAmtOwed.BackColor = &H80000005
  fpCurrCashPd.BackColor = &H80000005
  fpCurrChkChrgPd.BackColor = &H80000005
  fpCurrTotRecd.BackColor = &H80000005
  fpCurrChngDue.BackColor = &H80000005
  fpCurrTotOwed.BackColor = &H80000005
  fpCurrTotPaid.BackColor = &H80000005
  fpCurrPrincPaid.BackColor = &H80000005
  fpCurrPrincOwed.BackColor = &H80000005
  fpCurrIntPaid.BackColor = &H80000005
  fpCurrIntOwed.BackColor = &H80000005
  fpCurrAdvColPaid.BackColor = &H80000005
  fpCurrAdvColOwed.BackColor = &H80000005
  fpCurrLateListPaid.BackColor = &H80000005
  fpCurrLateListOwed.BackColor = &H80000005
  fpCurrRevOpt1Paid.BackColor = &H80000005
  fpCurrRevOpt1Owed.BackColor = &H80000005
  fpCurrRevOpt2Paid.BackColor = &H80000005
  fpCurrRevOpt2Owed.BackColor = &H80000005
  fpCurrRevOpt3Paid.BackColor = &H80000005
  fpCurrRevOpt3Owed.BackColor = &H80000005
  fpCurrDisc.BackColor = &H80000005
  fpCurrTotWDisc.BackColor = &H80000005
  fpCurrPrePay.BackColor = &H80000005
End Sub

Private Sub ReFigure()
  fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - CDbl(fpCurrTotPaid.Value))
  If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
End Sub

Private Sub LoadTemps()
  'Temp variables are used to reset changes back to what the values were
  'before changes were made...used extensively in Check4Changes
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim x As Integer
  Dim NumOfPayRecs As Integer
  
  TempAcctNum = fpLongAcctNum.Value
  OpenTempPayFile PayHandle, OperNum
  NumOfPayRecs = LOF(PayHandle) / Len(PayRec)
  For x = 1 To NumOfPayRecs
    Get PayHandle, x, PayRec
    If PayRec.CustAcct = CDbl(fpLongAcctNum.Text) Then
      TempPayDate = PayRec.PayDate '2/14/06
      TempAmtOwed = PayRec.AmtOwed
      TempTenderTY = QPTrim$(PayRec.TenderTY)
      TempCashAmt = PayRec.CashAmt
      TempChkAmt = PayRec.ChkAmt
      TempChrgAmt = PayRec.ChrgAmt
      TempAmtRecd = PayRec.AmtRecd
      TempChange = PayRec.Change
      TempDesc = QPTrim$(PayRec.Desc)
      TempPaidOwed1AmtOwed = PayRec.PaidOwed(1).AmtOwed
      TempPaidOwed2AmtOwed = PayRec.PaidOwed(2).AmtOwed
      TempPaidOwed3AmtOwed = PayRec.PaidOwed(3).AmtOwed
      TempPaidOwed4AmtOwed = PayRec.PaidOwed(4).AmtOwed
      TempPaidOwed5AmtOwed = PayRec.PaidOwed(5).AmtOwed
      TempPaidOwed6AmtOwed = PayRec.PaidOwed(6).AmtOwed
      TempPaidOwed7AmtOwed = PayRec.PaidOwed(7).AmtOwed
      TempPaidOwed1AmtPaid = PayRec.PaidOwed(1).AmtPaid
      TempPaidOwed2AmtPaid = PayRec.PaidOwed(2).AmtPaid
      TempPaidOwed3AmtPaid = PayRec.PaidOwed(3).AmtPaid
      TempPaidOwed4AmtPaid = PayRec.PaidOwed(4).AmtPaid
      TempPaidOwed5AmtPaid = PayRec.PaidOwed(5).AmtPaid
      TempPaidOwed6AmtPaid = PayRec.PaidOwed(6).AmtPaid
      TempPaidOwed7AmtPaid = PayRec.PaidOwed(7).AmtPaid
      TempTotOwed = PayRec.TotOwed
      TempAmtPaid = PayRec.AmtPaid
      TempTotPaid = PayRec.TotPaid
'      TempAcctNum = CLng(fpLongAcctNum.Value)
      Exit For
    End If
  Next x
  If x > NumOfPayRecs Then
    TempPayDate = Date2Num(PayDate) '2/14/06
    TempAmtOwed = 0
    TempTenderTY = "CASH"
    TempCashAmt = 0
    TempChkAmt = 0
    TempChrgAmt = 0
    TempAmtRecd = 0
    TempChange = 0
    TempDesc = ""
    TempPaidOwed1AmtOwed = 0
    TempPaidOwed2AmtOwed = 0
    TempPaidOwed3AmtOwed = 0
    TempPaidOwed4AmtOwed = 0
    TempPaidOwed5AmtOwed = 0
    TempPaidOwed6AmtOwed = 0
    TempPaidOwed7AmtOwed = 0
    TempPaidOwed1AmtPaid = 0
    TempPaidOwed2AmtPaid = 0
    TempPaidOwed3AmtPaid = 0
    TempPaidOwed4AmtPaid = 0
    TempPaidOwed5AmtPaid = 0
    TempPaidOwed6AmtPaid = 0
    TempPaidOwed7AmtPaid = 0
    TempTotOwed = 0
    TempAmtPaid = 0
    TempTotPaid = 0
'    TempAcctNum = 0
  End If
  
  Close PayHandle
End Sub

Public Function Check4Changes() As Boolean
  Dim PayRec As TaxPaymentRecType
  Dim PayHandle As Integer
  Dim x As Integer
  Dim NumOfPayRecs As Integer
  Dim Operator$
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisDesc As String
  Dim ThatDesc As String
  Dim ThisText As String
  Dim ThisDbl As Double
  Dim ThatDbl As Double
  Dim ThisInt As Integer
  Dim ThatInt As Integer
  
  'on error goto ERRORSTUFF
  Check4Changes = False
  If fpLongAcctNum.Value = 0 And BillCnt = 0 Then Exit Function
  
  Set ThisControl = fptxtPayDate
  ThisText = fptxtPayDate.Text
  ThisDesc = MakeRegDate(TempPayDate)
  If ThisText <> ThisDesc Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
'  Set ThisControl = fpCurrAmtOwed
'  thisDbl = CDbl(fpCurrAmtOwed.Value)
'  thatDbl = TempAmtOwed
'  If thisDbl <> thatDbl Then
'    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
'    frmTaxMsgW3Opts.Label1.Top = 800
'    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
'    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
'    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
'    frmTaxMsgW3Opts.Show vbModal
'    choice = frmTaxMsgW3Opts.fptxtChoice.Text
'    Unload frmTaxMsgW3Opts
'    If choice = "continue" Then
'      DontExit = False
''      Close PayHandle
'      Call cmdSave_Click
'      Exit Function
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fpcmbTenderType
  ThisDesc = QPTrim$(fpcmbTenderType.Text)
  ThatDesc = TempTenderTY
  If ThisDesc <> ThatDesc Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrCashPd
  ThisDbl = CDbl(fpCurrCashPd.Value)
  ThatDbl = TempCashAmt
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrChkChrgPd
  ThisDbl = CDbl(fpCurrChkChrgPd.Value)
  If TempTenderTY = "CHECK" Or TempTenderTY = "CASH AND CHECK" Then
    ThatDbl = TempChkAmt
  ElseIf TempTenderTY = "CHARGE" Then
    ThatDbl = TempChrgAmt
  Else
    ThatDbl = 0
  End If
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
'  Set ThisControl = fpCurrTotRecd
'  thisDbl = CDbl(fpCurrTotRecd.Value)
'  thatDbl = TempAmtRecd
'  If thisDbl <> thatDbl Then
'    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
'    frmTaxMsgW3Opts.Label1.Top = 800
'    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
'    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
'    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
'    frmTaxMsgW3Opts.Show vbModal
'    choice = frmTaxMsgW3Opts.fptxtChoice.Text
'    Unload frmTaxMsgW3Opts
'    If choice = "continue" Then
'      DontExit = False
''      Close PayHandle
'      Call cmdSave_Click
'      Exit Function
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fpCurrChngDue
  ThisDbl = CDbl(fpCurrChngDue.Value)
  ThatDbl = TempChange
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
'      Close PayHandle
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDescription
  ThisDesc = QPTrim$(fptxtDescription.Text)
  If ThisDesc = "" Then ThisDesc = "BLANK"
  ThatDesc = TempDesc
  If ThatDesc = "" Then ThatDesc = "BLANK"
  If ThisDesc <> ThatDesc Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.Show vbModal
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
'      Close PayHandle
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrPrePay
  ThisDbl = CDbl(fpCurrPrePay.Value)
  ThatDbl = TempPrePay
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
'      Close PayHandle
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrIntPaid
  ThisDbl = CDbl(fpCurrIntPaid.Value)
  ThatDbl = TempPaidOwed2AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrAdvColPaid
  ThisDbl = CDbl(fpCurrAdvColPaid.Value)
  ThatDbl = TempPaidOwed3AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrLateListPaid
  ThisDbl = CDbl(fpCurrLateListPaid.Value)
  ThatDbl = TempPaidOwed4AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrPrincPaid
  ThisDbl = CDbl(fpCurrPrincPaid.Value)
  ThatDbl = TempPaidOwed1AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrRevOpt1Paid
  ThisDbl = CDbl(fpCurrRevOpt1Paid.Value)
  ThatDbl = TempPaidOwed5AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrRevOpt2Paid
  ThisDbl = CDbl(fpCurrRevOpt2Paid.Value)
  ThatDbl = TempPaidOwed6AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrRevOpt3Paid
  ThisDbl = CDbl(fpCurrRevOpt3Paid.Value)
  ThatDbl = TempPaidOwed7AmtPaid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrTotOwed
  ThisDbl = CDbl(fpCurrTotOwed.Value)
  ThatDbl = TempTotOwed
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrTotRecd
  ThisDbl = CDbl(fpCurrTotRecd.Value)
  ThatDbl = TempAmtRecd 'Paid
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrTotPaid
  ThisDbl = CDbl(fpCurrTotPaid.Value)
  ThatDbl = TempTotPaid + TempPrePay
  If ThisDbl <> ThatDbl Then
    frmTaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmTaxMsgW3Opts.Label1.Top = 800
    frmTaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmTaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmTaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmTaxMsgW3Opts.Show vbModal
    choice = frmTaxMsgW3Opts.fptxtChoice.Text
    Unload frmTaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Exit Function
  
HandleChoice:
  Close PayHandle
  Select Case choice
    Case "abort" 'don't save
      If Exist("C:\CPWork\editpyment.dat") Or Lookup = True Then '2/14/06 added Or Lookup
        Exit Function 'trying to access another customer
      ElseIf TempAcctNum = CLng(fpLongAcctNum.Value) Then
        frmTaxPayMenu.Show
        DoEvents
        KillFile "C:\CPWork\txpyment.dat"
        Unload Me
      End If
      Exit Function
    Case "option" 'review
      fpLongAcctNum = TempAcctNum
      If ThisControl.Enabled = True Then
        ThisControl.SetFocus
      Else
        ThisControl.BackColor = &H8080FF
      End If
      Close PayHandle
      Check4Changes = True
      Exit Function
    Case Else
  End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "Check4Changes", Erl)
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

Private Sub fptxtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyTab Then
    fpLongAcctNum.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrRevOpt3Paid.Enabled = True Then
      fpCurrRevOpt3Paid.SetFocus
    ElseIf fpCurrRevOpt2Paid.Enabled = True Then
      fpCurrRevOpt2Paid.SetFocus
    ElseIf fpCurrRevOpt1Paid.Enabled = True Then
      fpCurrRevOpt1Paid.SetFocus
    Else
      fpCurrLateListPaid.SetFocus
    End If
  End If
End Sub

Private Sub fptxtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpLongAcctNum.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtDescription.SetFocus
  End If
End Sub

Private Function CheckOverPay() As Boolean

  'on error goto ERRORSTUFF
  
  'looks for overpayment of revenues if others are not fully paid...not allowed
  CheckOverPay = False
  If CDbl(fpCurrIntPaid.Value) = CDbl(fpCurrIntOwed.Value) Then
    GoTo IntOK
  ElseIf CDbl(fpCurrIntOwed.Value) > CDbl(fpCurrIntPaid.Value) Then
    If OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal (plus Discount) while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Adv/Collect while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrLateListOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Late Listing while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt1Owed.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt1.Text) + " while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt2Owed.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt2.Text) + " while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt3Owed.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt3.Text) + " while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
  
IntOK:
      
  If OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) = CDbl(fpCurrPrincOwed.Value) Then
    GoTo PrincOK
  ElseIf CDbl(fpCurrPrincOwed.Value) > CDbl(fpCurrPrincPaid.Value) Then
    If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Principal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Principal (plus Discount). Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Adv/Collect while underpaying Principal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrLateListOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Late Listing while underpaying Principal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt1Paid.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt1.Text) + " while underpaying Principal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt2Paid.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt2.Text) + " while underpaying Principal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt3Paid.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt3.Text) + " while underpaying Principal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
PrincOK:

  If CDbl(fpCurrAdvColPaid.Value) = CDbl(fpCurrAdvColOwed.Value) Then
    GoTo AdvColOK
  ElseIf CDbl(fpCurrAdvColOwed.Value) > CDbl(fpCurrAdvColPaid.Value) Then
    If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrAdvColOwed.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal (plus Discount) while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrLateListOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Late Listing while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt1Paid.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt1.Text) + " while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt2Paid.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt2.Text) + " while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt3Paid.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt3.Text) + " while underpaying Adv/Collect. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
AdvColOK:

  If CDbl(fpCurrLateListPaid.Value) = CDbl(fpCurrLateListOwed.Value) Then
    GoTo LateListOK
  ElseIf CDbl(fpCurrLateListOwed.Value) > CDbl(fpCurrLateListPaid.Value) Then
    If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal (and Discount) while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Adv/Collect while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt1Paid.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt1.Text) + " while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt2Paid.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt2.Text) + " while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt3Paid.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt3.Text) + " while underpaying Late Listing. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
LateListOK:

  If CDbl(fpCurrRevOpt1Paid.Value) = CDbl(fpCurrRevOpt1Owed.Value) Then
    GoTo Rev1OK
  ElseIf CDbl(fpCurrRevOpt1Owed.Value) > CDbl(fpCurrRevOpt1Paid.Value) Then
    If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal (plus Discount) while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Adv/Collect while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Late Listing while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt2Paid.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt2.Text) + " while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt3Paid.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt3.Text) + " while underpaying " + QPTrim$(fptxtRevOpt1.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
Rev1OK:

  If CDbl(fpCurrRevOpt2Paid.Value) = CDbl(fpCurrRevOpt2Owed.Value) Then
    GoTo Rev2OK
  ElseIf CDbl(fpCurrRevOpt2Owed.Value) > CDbl(fpCurrRevOpt2Paid.Value) Then
    If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal (plus Discount) while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Adv/Collect while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Late Listing while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt1Paid.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt1.Text) + " while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt3Paid.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt3.Text) + " while underpaying " + QPTrim$(fptxtRevOpt2.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
Rev2OK:

  If CDbl(fpCurrRevOpt3Paid.Value) = CDbl(fpCurrRevOpt3Owed.Value) Then
    GoTo Rev3OK
  ElseIf CDbl(fpCurrRevOpt3Owed.Value) > CDbl(fpCurrRevOpt3Paid.Value) Then
    If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrPrincOwed.BackColor = &H8080FF
      fpCurrPrincPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Principal (plus Discount) while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrAdvColPaid.BackColor = &H8080FF
      fpCurrAdvColOwed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Adv/Collect while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
      fpCurrLateListOwed.BackColor = &H8080FF
      fpCurrLateListPaid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay Late Listing while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt1Paid.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
      fpCurrRevOpt1Owed.BackColor = &H8080FF
      fpCurrRevOpt1Paid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt1.Text) + " while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrRevOpt2Paid.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
      fpCurrRevOpt2Owed.BackColor = &H8080FF
      fpCurrRevOpt2Paid.BackColor = &H8080FF
      fpCurrRevOpt3Paid.BackColor = &H8080FF
      fpCurrRevOpt3Owed.BackColor = &H8080FF
      frmTaxMsg.Label1.Caption = "The customer is attempting to overpay " + QPTrim$(fptxtRevOpt2.Text) + " while underpaying " + QPTrim$(fptxtRevOpt3.Text) + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmTaxMsg.Label1.Top = 600
      frmTaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
Rev3OK:
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "CheckOverPay", Erl)
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

End Function

Private Sub PrintReceipt()
  Dim PayRec As TaxPaymentRecType
  Dim PHandle As Integer
  Dim RHandle As Integer
  Dim Oper$
  Dim MasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim TownName$
  Dim PostDate$
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  Dim PayRecpName$
  Dim RHandle2 As Integer, PayRecpName2 As String, RptHandle2 As Integer
  
  'on error goto ERRORSTUFF
  
  PayRecpName2$ = StartPath$ + "TXVLD" + Oper$ + ".Rpt"
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, MasterRec
  Close MHandle
  
  Oper$ = CStr(OperNum)
  OpenTempPayFile PHandle, OperNum
  Get PHandle, GPayNum, PayRec
  Close PHandle
 
  TownName$ = QPTrim$(MasterRec.Name)
  PostDate$ = MakeRegDate(PayRec.PayDate)
  PayRecpName$ = "C:\CPWork\TAXRCP" + Oper$ + ".RPT"
  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  Print #RHandle, Chr$(7)
  Print #RHandle, TownName$
  Print #RHandle, "TAX PAYMENT"
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, QPTrim$(PayRec.CustName)
  Print #RHandle, QPTrim$(PayRec.CustAddr)
  Print #RHandle, QPTrim$(PayRec.Desc)
  Print #RHandle, "Acct. No. "; PayRec.CustAcct
  Print #RHandle,
  Print #RHandle, "Total Owed: "; QPTrim$(Using("$##,##0.00", PayRec.AmtOwed))
  Print #RHandle, "Total Paid: "; QPTrim$(Using("$##,##0.00", PayRec.AmtPaid))
  Print #RHandle, "  Discount: "; QPTrim$(Using("$##,##0.00", PayRec.DiscAmt))
  Print #RHandle, "Change Due: "; QPTrim$(Using("$##,##0.00", PayRec.Change))
  Print #RHandle,
  Print #RHandle,
  Print #RHandle, "Operator: "; CStr(OperNum)
  Print #RHandle, '"Receipt#: "; Using("$##,##0.00", FileSize("TAXCPR" + Oper$ + ".DAT") \ TaxPayRecLen)
  Print #RHandle,
  Print #RHandle, "       T H A N K   Y O U !"
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,

  Close RHandle
  
  If CntrlDef = 1 Then
    Call cmdDrawer_Click
  End If
10:
  DefPrinter = RecpPort
20:
  
  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
      If frmTaxPrint.cmdCancel = False Then
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
  If QPTrim(PayRec.TenderTY) = "CHECK" Or QPTrim(PayRec.TenderTY) = "CASH AND CHECK" Then
   If RctValidate Then
     RHandle2 = FreeFile
     Open PayRecpName2$ For Output As RHandle2
     Print #LPTHandle, Chr$(27); Chr$(&H63); Chr$(&H30); Chr$(&H4)
     Print #LPTHandle, Chr$(13); Chr$(10)
     Print #RHandle2, TownName$
     Print #RHandle2, "FOR DEPOSIT ONLY"
     Print #RHandle2, "Acct. No. "; PayRec.CustAcct
     Print #RHandle2, "Date: "; PostDate$
     Print #RHandle2, "Time: "; Time
     Print #RHandle2,
     Print #LPTHandle, Chr$(12)
     Close RHandle2
     LPTHandle = FreeFile
     Open DefPrinter For Output As LPTHandle
     RptHandle2 = FreeFile
     Open PayRecpName2$ For Input As RptHandle2
     Do
       If frmTaxPrint.cmdCancel = False Then
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
    MainLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(PayRec.CustAcct)
  End If
 End If

70:
 MainLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(PayRec.CustAcct)
 KillFile PayRecpName$
80:
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "PrintReceipt", Erl)
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

Private Sub fptxtPayDate_LostFocus()
  Dim WhichRec() As Integer
  
  'on error goto ERRORSTUFF
  
  'warns user if the date changes causing an existing discount to change
  If NotFirstLoad = False Then Exit Sub
  If CDbl(fpCurrDisc.Value) > 0 And Check4Discounts = 0 Then
    frmTaxMsgWOpts.Label1.Caption = "Changing the date from " + MakeRegDate(TempPayDate) + " to " + QPTrim$(fptxtPayDate.Text) + " will disqualify this customer from an existing discount. If you wish to continue with the new date which will automatically recalculate the amounts owed then press F10. Otherwise, press ESC to leave the date untouched."
    frmTaxMsgWOpts.Label1.Top = 600
    frmTaxMsgWOpts.cmdExit.Text = "ESC Restore Date"
    frmTaxMsgWOpts.cmdCont.Text = "F10 New Date OK"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtPayDate = MakeRegDate(TempPayDate)
    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      Call RemoveDiscount
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "fptxtPayDate_LostFocus", Erl)
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

Private Sub ApplyDiscount()
  Dim ColTot As Double
  Dim x As Integer
  
  'on error goto ERRORSTUFF
  
  For x = 1 To BillCnt
    TempBillList(x).Principle1 = OldRound(TempBillList(x).Principle1 - TempBillList(x).DiscAmt)
  Next x
  
  ColTot = 0
  fpCurrDisc = MaxDisc
  fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincOwed.Value) - CDbl(fpCurrDisc.Value))
  GoSub AddCol
  
  If CDbl(fpCurrPrincPaid.Value) < ThisDiscAmt Then
    fpCurrDisc = CDbl(fpCurrPrincPaid.Value)
    fpCurrPrincPaid = 0
    Call AddUpPaidCol
  ElseIf OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPrincOwed.Value) Then
    fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincOwed.Value) - CDbl(fpCurrDisc.Value))
  Else
    If ColTot > CDbl(fpCurrTotRecd.Value) Then
      If OldRound(ColTot - CDbl(fpCurrTotRecd.Value)) = CDbl(fpCurrDisc.Value) Then
        fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincPaid.Value) - CDbl(fpCurrDisc.Value))
      End If
    End If
  End If
  
  Call ReFigure
  
  MainLog ("frmTaxPaymentEntry: Customer, " + fptxtName.Text + ", is eligible for a discount of " + QPTrim$(Using$("$#,##0.00", ThisDiscAmt)) + " and the user allowed the discount to apply.")
  
  Exit Sub
  
AddCol:
'  ColTot = OldRound(CDbl(fpCurrDisc.Value) + CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrIntPaid.Value))
  ColTot = OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrIntPaid.Value))
  ColTot = OldRound(ColTot# + CDbl(fpCurrAdvColPaid.Value) + CDbl(fpCurrLateListPaid.Value))
  ColTot = OldRound(ColTot# + CDbl(fpCurrRevOpt1Paid.Value) + CDbl(fpCurrRevOpt2Paid.Value) + CDbl(fpCurrRevOpt3Paid.Value))
  fpCurrTotPaid = ColTot
  fpCurrTotWDisc = OldRound(ColTot + CDbl(fpCurrDisc.Value))
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "ApplyDiscount", Erl)
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

Private Sub Distribute(WhatsLeft As Double)
  Dim SetUpRec As TaxMasterType
  Dim SHandle As Integer
  Dim x As Integer
  Dim TotRecd As Double
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TransRecord&
  Dim PaidDif As Double
  Dim ThisDif As Double
  Dim TPayRec As PayListType
  Dim PayRec As PayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim y As Integer, z As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim Nextx As Integer
  Dim SmallNum As Integer
  Dim HoldNum As Long
  Dim HoldDate As Integer
  Dim Thisx As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim ThisTaxYear As Integer
  Dim Message$
  Dim ThisBal As Double
  Dim DiscCheck As Integer
  Dim ThisPct As Double
  Dim TotPaid#
  
  'on error goto ERRORSTUFF
  
  If CDbl(fpCurrAmtOwed.Value) = 0 Then '8/12/05
    If TaxMsgWOpts(800, "Since this customer does not owe any money automatic distribution will place the amount paid in the 'Prepay Amt' field. Press F10 to OK this distribution.", "F10 OK", "ESC Abort") = "abort" Then
      Unload frmTaxMsgWOpts
      Exit Sub
    Else
      fpCurrPrePay = WhatsLeft
      fpCurrChngDue = 0
      NotFirstLoad = True 'added 6/1/06
      Call AddUpPaidCol
      Call LoadTempPayList
      Exit Sub
    End If
  End If
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  
  ThisTaxYear = TaxMasterRec.TaxYear
  fpCurrIntPaid.Value = 0
  fpCurrAdvColPaid.Value = 0
  fpCurrLateListPaid.Value = 0
  fpCurrPrincPaid.Value = 0
  fpCurrRevOpt1Paid.Value = 0
  fpCurrRevOpt2Paid.Value = 0
  fpCurrRevOpt3Paid.Value = 0
'  PrePayAmt = CDbl(fpCurrPrePay.Value)
  Nextx = 1
  SmallNum = 30000
  ReDim TransSeq(1 To BillCnt) As Long
  Do While Nextx <= BillCnt
    For x = Nextx To BillCnt
      If BillDate(x) <= SmallNum Then
        SmallNum = BillDate(x)
        Thisx = x
      End If
    Next x
    HoldNum = BillTrans(Nextx)
    HoldDate = BillDate(Nextx)
    BillTrans(Nextx) = BillTrans(Thisx)
    BillDate(Nextx) = BillDate(Thisx)
    BillTrans(Thisx) = HoldNum
    BillDate(Thisx) = HoldDate
    Nextx = Nextx + 1
    SmallNum = 30000
  Loop
  
  ReDim Preserve TempBillList(1 To BillCnt) As PayListType
  TempBillListCnt = 0
  For x = 1 To BillCnt
    TempBillList(x).Interest1 = 0
    TempBillList(x).Collection = 0
    TempBillList(x).LateList = 0
    TempBillList(x).Principle1 = 0
    TempBillList(x).OptRev1 = 0
    TempBillList(x).OptRev2 = 0
    TempBillList(x).OptRev3 = 0
    TempBillList(x).BillRec = 0
    TempBillList(x).CustRec = 0
    TempBillList(x).TaxYear = 0
    TempBillList(x).TotPaid = 0
  Next x
   
  ReDim PaySeq(1 To BillCnt, 1 To 7) As Double 'Payments are applied by priority. The first
  '4 are hard coded. The final 3 are determined by the order the user enters
  'them on the System Setup screen (last tab)
  If EditFlag = False Or (EditFlag = True And BillCnt > 0) Then 'If EditFlag is
  'false then this is a new customer and BillCnt will be > 0 since this function
  'is not accessible unless there is an amount in the amount owed field
    OpenTaxTransFile THandle, NumOfTRecs
    For x = 1 To BillCnt
      Get THandle, BillTrans(x), TaxTrans
        TempBillList(x).BillRec = BillTrans(x)
        TempBillList(x).CustRec = GCustNum
        TempBillList(x).TaxYear = TaxTrans.TaxYear
        PaySeq(x, 1) = OldRound(PaySeq(x, 1) + TaxTrans.Revenue.Interest)
        PaySeq(x, 1) = OldRound(PaySeq(x, 1) - TaxTrans.Revenue.InterestPd)
        TempBillList(x).TotOwed = PaySeq(x, 1)
        PaySeq(x, 2) = OldRound(PaySeq(x, 2) + TaxTrans.Revenue.Collection)
        PaySeq(x, 2) = OldRound(PaySeq(x, 2) - TaxTrans.Revenue.CollectionPd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 2))
        PaySeq(x, 3) = OldRound(PaySeq(x, 3) + TaxTrans.Revenue.LateList)
        PaySeq(x, 3) = OldRound(PaySeq(x, 3) - TaxTrans.Revenue.LateListPd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 3))
        PaySeq(x, 4) = OldRound(PaySeq(x, 4) + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        PaySeq(x, 4) = OldRound(PaySeq(x, 4) + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        PaySeq(x, 4) = OldRound(PaySeq(x, 4) - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
        PaySeq(x, 4) = OldRound(PaySeq(x, 4) - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 4))
        PaySeq(x, 5) = OldRound(PaySeq(x, 5) + TaxTrans.Revenue.RevOpt1)
        PaySeq(x, 5) = OldRound(PaySeq(x, 5) - TaxTrans.Revenue.RevOpt1Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 5))
        PaySeq(x, 6) = OldRound(PaySeq(x, 6) + TaxTrans.Revenue.RevOpt2)
        PaySeq(x, 6) = OldRound(PaySeq(x, 6) - TaxTrans.Revenue.RevOpt2Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 6))
        PaySeq(x, 7) = OldRound(PaySeq(x, 7) + TaxTrans.Revenue.RevOpt3)
        PaySeq(x, 7) = OldRound(PaySeq(x, 7) - TaxTrans.Revenue.RevOpt3Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 7))
   Next x
   
'   fpCurrPrePay = 0
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 1) Then
       fpCurrIntPaid.Value = CDbl(fpCurrIntPaid.Value) + PaySeq(x, 1)
       TempBillList(x).Interest1 = PaySeq(x, 1)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest1)
     Else
       fpCurrIntPaid.Value = CDbl(fpCurrIntPaid.Value) + WhatsLeft
       TempBillList(x).Interest1 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest1)
     End If
 
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 1))
     If WhatsLeft <= 0 Then GoTo PlayedOut
   
     If WhatsLeft >= PaySeq(x, 2) Then
       fpCurrAdvColPaid.Value = CDbl(fpCurrAdvColPaid.Value) + PaySeq(x, 2)
       TempBillList(x).Collection = PaySeq(x, 2)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Collection)
     Else
       fpCurrAdvColPaid.Value = CDbl(fpCurrAdvColPaid.Value) + WhatsLeft
       TempBillList(x).Collection = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Collection)
     End If

     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 2))
     If WhatsLeft <= 0 Then GoTo PlayedOut
     
     If WhatsLeft >= PaySeq(x, 3) Then
       fpCurrLateListPaid.Value = CDbl(fpCurrLateListPaid.Value) + PaySeq(x, 3)
       TempBillList(x).LateList = PaySeq(x, 3)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).LateList)
     Else
       fpCurrLateListPaid.Value = CDbl(fpCurrLateListPaid.Value) + WhatsLeft
       TempBillList(x).LateList = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).LateList)
     End If
 
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 3))
     If WhatsLeft <= 0 Then GoTo PlayedOut
     
     If WhatsLeft >= PaySeq(x, 4) Then
       fpCurrPrincPaid.Value = CDbl(fpCurrPrincPaid.Value) + PaySeq(x, 4)
       TempBillList(x).Principle1 = PaySeq(x, 4)
       If TempBillList(x).DiscAmt > 0 Then
         If CDbl(fpCurrDisc.Value) = TempBillList(x).DiscAmt Then
           fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincPaid.Value) - TempBillList(x).DiscAmt)
           TempBillList(x).Principle1 = OldRound(TempBillList(x).Principle1 - TempBillList(x).DiscAmt)
         Else
           fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincPaid.Value - TempBillList(x).DiscAmt)) 'CDbl(fpCurrDisc.Value))changed on 1/19/07
           TempBillList(x).Principle1 = OldRound(TempBillList(x).Principle1 - TempBillList(x).DiscAmt) 'CDbl(fpCurrDisc.Value))changed on 1/19/07
         End If
       End If
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Principle1)
     Else
       fpCurrPrincPaid.Value = CDbl(fpCurrPrincPaid.Value) + WhatsLeft
       TempBillList(x).Principle1 = WhatsLeft
       If TempBillList(x).DiscAmt > 0 Then
         If CDbl(fpCurrDisc.Value) = TempBillList(x).DiscAmt Then
           fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincPaid.Value) - TempBillList(x).DiscAmt)
           TempBillList(x).Principle1 = OldRound(TempBillList(x).Principle1 - TempBillList(x).DiscAmt)
         Else
           fpCurrPrincPaid = OldRound(CDbl(fpCurrPrincPaid.Value - TempBillList(x).DiscAmt)) 'CDbl(fpCurrDisc.Value))changed on 1/19/07
           TempBillList(x).Principle1 = OldRound(TempBillList(x).Principle1 - TempBillList(x).DiscAmt) 'CDbl(fpCurrDisc.Value))changed on 1/19/07
         End If
       End If
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Principle1)
     End If
     
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 4))
     If WhatsLeft <= 0 Then GoTo PlayedOut

     If WhatsLeft >= PaySeq(x, 5) Then '1/16/07
       fpCurrRevOpt1Paid.Value = CDbl(fpCurrRevOpt1Paid.Value) + PaySeq(x, 5) '1/16/07
       TempBillList(x).OptRev1 = PaySeq(x, 5) '1/16/07
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev1)
     Else
       fpCurrRevOpt1Paid.Value = CDbl(fpCurrRevOpt1Paid.Value) + WhatsLeft '1/16/07
       TempBillList(x).OptRev1 = WhatsLeft '1/16/07
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev1)
     End If
 
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 5))
     If WhatsLeft <= 0 Then GoTo PlayedOut

     If WhatsLeft >= PaySeq(x, 6) Then
       fpCurrRevOpt2Paid.Value = CDbl(fpCurrRevOpt2Paid.Value) + PaySeq(x, 6)
       TempBillList(x).OptRev2 = PaySeq(x, 6)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev2)
     Else
       fpCurrRevOpt2Paid.Value = CDbl(fpCurrRevOpt2Paid.Value) + WhatsLeft
       TempBillList(x).OptRev2 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev2)
     End If
 
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 6))
     If WhatsLeft <= 0 Then GoTo PlayedOut

     If WhatsLeft >= PaySeq(x, 7) Then
       fpCurrRevOpt3Paid.Value = CDbl(fpCurrRevOpt3Paid.Value) + PaySeq(x, 7)
       TempBillList(x).OptRev3 = PaySeq(x, 7)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev3)
     Else
       fpCurrRevOpt3Paid.Value = CDbl(fpCurrRevOpt3Paid.Value) + WhatsLeft
       TempBillList(x).OptRev3 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev3)
     End If
 
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 7))
     If WhatsLeft <= 0 Then GoTo PlayedOut 'added 6/2/2011
   Next x
   
   Call AssignPaidTemps
   
 End If
 
PlayedOut:
  TotPaid# = OldRound(CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrAdvColPaid.Value) + CDbl(fpCurrLateListPaid.Value))
  TotPaid# = OldRound(TotPaid# + CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrRevOpt1Paid.Value) + CDbl(fpCurrRevOpt2Paid.Value) + CDbl(fpCurrRevOpt3Paid.Value))
  TotPaid# = OldRound(TotPaid# + CDbl(fpCurrPrePay.Value))
  fpCurrTotPaid = TotPaid#
  fpCurrChngDue.Value = OldRound(CDbl(fpCurrTotRecd.Value) - CDbl(fpCurrTotPaid.Value))
  If CDbl(fpCurrChngDue.Value) < 0 Then
    fpCurrChngDue = 0
  End If
  Close THandle 'added THandle on 7/18/07
 
  fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))

  GetNewCust = False
  
'  If MaxDisc > 0 And CDbl(fpCurrDisc.Value) = 0 Then 'remarked 9/20/05
'    If TaxMsgWOpts(800, "This customer is eligible for a discount. Press F10 if you wish to apply the discount. Otherwise, press ESC to override the discount.", "F10 Apply Discount", "ESC NO Discount") = "abort" Then
'      Unload frmTaxMsgWOpts
'      If CDbl(fpCurrDisc.Value) > 0 Then
'        Call RemoveDiscount
'      End If
'      Exit Sub
'    Else
'      Unload frmTaxMsgWOpts
'      ApplyDiscount
'      Call ReassignDiscount
'    End If
'  End If
  If EditFlag = True Then DistrFlag = True 'added 10/20/06
    
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "Distribute", Erl)
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

Private Sub RemoveDiscount()
  Dim x As Integer
  Dim TempPayRec As PayListType
  Dim THandle As Integer
  Dim NumOfTRecs As Integer
  
  If BillCnt = 0 Then
    OpenPayListFile THandle, OperNum
    NumOfTRecs = LOF(THandle) / Len(TempPayRec)
    For x = 1 To NumOfTRecs
      Get THandle, x, TempPayRec
      If TempPayRec.CustRec = GCustNum Then
        TempPayRec.DiscAmt = 0
        Put THandle, x, TempPayRec
      End If
    Next x
    Close THandle
  Else
    For x = 1 To BillCnt
      TempBillList(x).DiscAmt = 0
      TempBillList(x).Principle1 = OldRound(TempBillList(x).Principle1 + TempBillList(x).DiscAmt)
    Next x
  End If
  
  ThisDiscAmt = 0
  fpCurrDisc = 0
  
  Call AddUpPaidCol
  
  Call AssignPaidTemps
End Sub

Private Sub OverRideDist()
  'this sub handles user entered overrides after auto distribution takes place
  Dim Message$
  Dim Top As Integer
  Dim OptRev$
  
  'on error goto ERRORSTUFF
  
  'we are looking for amounts owed versus amounts paid to see if there are
  'any shortfalls...if found then shortfalls negate discounts
  If CDbl(fpCurrDisc.Value) <= 0 Then Exit Sub
  
  InOverRideDist = True
  
  If CDbl(fpCurrIntOwed.Value) > CDbl(fpCurrIntPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the interest portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic interest payment of " + QPTrim$(Using("$##,##0.00", TempIntPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new interest amount of " + fpCurrIntPaid.Text + ".")
    End If
  End If
  
  If CDbl(fpCurrAdvColOwed.Value) > CDbl(fpCurrAdvColPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the adv/collect portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic adv/collect payment of " + QPTrim$(Using("$##,##0.00", TempAdvColPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new adv/collect amount of " + fpCurrAdvColPaid.Text + ".")
    End If
  End If
  
  If CDbl(fpCurrLateListOwed.Value) > CDbl(fpCurrLateListPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the late list portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic late list payment of " + QPTrim$(Using("$##,##0.00", TempLateListPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new late list amount of " + fpCurrLateListPaid.Text + ".")
    End If
  End If
  
  If CDbl(fpCurrPrincOwed.Value) > OldRound(CDbl(fpCurrPrincPaid.Value) + CDbl(fpCurrDisc.Value)) Then
    Message = "This payment configuration eliminates the discount because now the principle portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic principal payment of " + QPTrim$(Using("$##,##0.00", TempPrincPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new principal amount of " + fpCurrPrincPaid.Text + ".")
    End If
  End If
  
  If fpCurrRevOpt1Paid.Enabled = False Then GoTo Next2
  OptRev = QPTrim$(fptxtRevOpt1.Text)
  If CDbl(fpCurrRevOpt1Owed.Value) > CDbl(fpCurrRevOpt1Paid.Value) Then
    Message = "This payment configuration eliminates the discount because now the " + OptRev + " portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic " + OptRev + " payment of " + QPTrim$(Using("$##,##0.00", TempRev1Paid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new " + OptRev + " amount of " + fpCurrRevOpt1Paid.Text + ".")
    End If
  End If
  
Next2:
  If fpCurrRevOpt2Paid.Enabled = False Then GoTo Next3
  OptRev = QPTrim$(fptxtRevOpt2.Text)
  If CDbl(fpCurrRevOpt2Owed.Value) > CDbl(fpCurrRevOpt2Paid.Value) Then
    Message = "This payment configuration eliminates the discount because now the " + OptRev + " portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic " + OptRev + " payment of " + QPTrim$(Using("$##,##0.00", TempRev2Paid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new " + OptRev + " amount of " + fpCurrRevOpt2Paid.Text + ".")
    End If
  End If
  
Next3:
  If fpCurrRevOpt2Paid.Enabled = False Then GoTo Next4
  OptRev = QPTrim$(fptxtRevOpt3.Text)
  If CDbl(fpCurrRevOpt3Owed.Value) > CDbl(fpCurrRevOpt3Paid.Value) Then
    Message = "This payment configuration eliminates the discount because now the " + OptRev + " portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If TaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmTaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      fpCurrDisc.Value = 0
      MainLog ("WARNING: User warned that overriding the automatic " + OptRev + " payment of " + QPTrim$(Using("$##,##0.00", TempRev3Paid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new " + OptRev + " amount of " + fpCurrRevOpt3Paid.Text + ".")
    End If
  End If
  
Next4:
  Call AddUpPaidCol
  InOverRideDist = False
  Call AssignPaidTemps
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "OverRideDist", Erl)
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

Private Sub AssignPaidTemps()
  TempPrincPaid = CDbl(fpCurrPrincPaid.Value)
  TempIntPaid = CDbl(fpCurrIntPaid.Value)
  TempAdvColPaid = CDbl(fpCurrAdvColPaid.Value)
  TempLateListPaid = CDbl(fpCurrLateListPaid.Value)
  TempRev1Paid = CDbl(fpCurrRevOpt1Paid.Value)
  TempRev2Paid = CDbl(fpCurrRevOpt2Paid.Value)
  TempRev3Paid = CDbl(fpCurrRevOpt3Paid.Value)
  TempDisc = CDbl(fpCurrDisc.Value)
  TempTotPd = CDbl(fpCurrTotPaid.Value)
  TempPrePay = CDbl(fpCurrPrePay.Value)
End Sub

Private Sub ReLoadPaidTemps()
  fpCurrPrincPaid = TempPrincPaid
  fpCurrIntPaid = TempIntPaid
  fpCurrAdvColPaid = TempAdvColPaid
  fpCurrLateListPaid = TempLateListPaid
  fpCurrRevOpt1Paid = TempRev1Paid
  fpCurrRevOpt2Paid = TempRev2Paid
  fpCurrRevOpt3Paid = TempRev3Paid
  fpCurrDisc = TempDisc
  fpCurrTotPaid = TempTotPd
  fpCurrPrePay = TempPrePay
End Sub

Private Sub GetMaxDisc()
  Dim TPayRec As PayListType
  Dim PayRec As PayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim Operator$
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim x As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim MHandle As Integer
  Dim ThisDiscPct As Double
  Dim TaxTRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim ThisTaxYear As Integer
  Dim Balance#
  Dim Nextx As Integer
  Dim SmallNum As Integer
  Dim HoldNum As Long
  Dim HoldDate As Integer
  Dim Thisx As Integer
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  ThisDiscPct = TaxMasterRec.DisPct
  MaxDisc = 0
'  If TaxMasterRec.DiscXDate > 0 Then
'    DiscYN = True
'  End If
  
  ThisTaxYear = TaxMasterRec.TaxYear
  
  If BillCnt = 0 And EditFlag = True Then 'Or Exist("C:\CPWork\editpyment.dat") Then 'user is editing and is not accessing
  'the bill list
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    ThisPrevRec = 0
    NewRec = 0
    Operator$ = CStr(OperNum)
    Operator$ = QPTrim$(Operator$)
    OpenPayListFile PHandle, OperNum 'saved by getting data from temporary
    'bill record
    NumOfPRecs = LOF(PHandle) / Len(PayRec)
    For x = 1 To NumOfPRecs
      Get PHandle, x, PayRec
      If PayRec.CustRec = GCustNum And PayRec.PrevListRec <> -1 Then 'added <> -1 because
      '-1 means that transaction has been deleted 6/29/06
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        BillTrans(BillCnt) = PayRec.BillRec
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillDate(BillCnt) = TempRec.BillDate
      End If
    Next x
    Close PHandle
  ElseIf Exist(TempBillRecs) Then
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    BillCnt = 0
    OpenTempBillRecs TempHandle, NumOfTemps
    For x = 1 To NumOfTemps
      Get TempHandle, x, TempRec
      If TempRec.BillRec > 0 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillTrans(BillCnt) = TempRec.BillPtr
        BillDate(BillCnt) = TempRec.BillDate
        'this data should be the same data as that where PaySeq() are loaded
      End If
    Next x
    Close TempHandle
  End If
  
  If BillCnt = 0 Then Exit Sub
  Nextx = 1
  SmallNum = 30000
  ReDim TransSeq(1 To BillCnt) As Long
  Do While Nextx <= BillCnt
    For x = Nextx To BillCnt
      If BillDate(x) <= SmallNum Then
        SmallNum = BillDate(x)
        Thisx = x
      End If
    Next x
    HoldNum = BillTrans(Nextx)
    HoldDate = BillDate(Nextx)
    BillTrans(Nextx) = BillTrans(Thisx)
    BillDate(Nextx) = BillDate(Thisx)
    BillTrans(Thisx) = HoldNum
    BillDate(Thisx) = HoldDate
    Nextx = Nextx + 1
    SmallNum = 30000
  Loop
  
  ReDim Preserve TempBillList(1 To BillCnt) As PayListType
  
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To BillCnt
    If BillTrans(x) <= 0 Then
      Call DisablePayFields
      OverPay = True
      GoTo OverPay
    End If
    Get THandle, BillTrans(x), TaxTRec
      Balance = 0
    'next line...added TaxTRec.DiscXDate > 0 And DiscXDate <= TaxTRec.DiscXDate in place of Date2Num(fptxtPayDate.Text)
    If TaxTRec.TaxYear = ThisTaxYear And TaxTRec.DiscXDate > 0 And DiscXDate <= TaxTRec.DiscXDate Then 'Date2Num(fptxtPayDate.Text) Then
'remarked out the following 7 lines on 2/8/07
      Balance# = OldRound(Balance# + TaxTRec.Revenue.Principle1 + TaxTRec.Revenue.Principle2 + TaxTRec.Revenue.Principle3)
      Balance# = OldRound(Balance# + TaxTRec.Revenue.Principle4 + TaxTRec.Revenue.Principle5 + TaxTRec.Revenue.LateList + TaxTRec.Revenue.Interest)
      Balance# = OldRound(Balance# + TaxTRec.Revenue.Collection + TaxTRec.Revenue.Penalty + TaxTRec.Revenue.RevOpt1 + TaxTRec.Revenue.RevOpt2 + TaxTRec.Revenue.RevOpt3)
      Balance# = OldRound(Balance# - (TaxTRec.Revenue.Principle1Pd + TaxTRec.Revenue.Principle2Pd + TaxTRec.Revenue.Principle3Pd))
      Balance# = OldRound(Balance# - (TaxTRec.Revenue.Principle4Pd + TaxTRec.Revenue.Principle5Pd + TaxTRec.Revenue.LateListPd + TaxTRec.Revenue.InterestPd))
      Balance# = OldRound(Balance# - TaxTRec.Revenue.CollectionPd + TaxTRec.Revenue.PenaltyPd + TaxTRec.Revenue.RevOpt1Pd + TaxTRec.Revenue.RevOpt2Pd + TaxTRec.Revenue.RevOpt3Pd)
      Balance# = OldRound(Balance# - TaxTRec.DiscAmt)
'      Balance = TaxTRec.Amount 'added 2/9/07 'remarked out 8/24/07
      If Balance# > 0 Then 'save which transaction the discount is applied to
        MaxDisc = MaxDisc + OldRound(Balance# * ThisDiscPct * 0.01)
        TempBillList(x).DiscAmt = OldRound(Balance# * ThisDiscPct * 0.01)
      End If
    End If
  Next x
OverPay:
  Close THandle
  
  If MaxDisc = 0 Then
    fpCurrDisc.ControlType = ControlTypeReadOnly
    If CDbl(fpCurrDisc.Value) > 0 Then
      fpCurrDisc = 0
    End If
  Else
    fpCurrDisc.ControlType = ControlTypeNormal
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "GetMaxDisc", Erl)
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

Private Sub LoadTempPayList()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim WhatsLeft As Double
  Dim x As Integer
  
  'on error goto ERRORSTUFF
  
  'this sub loads the tempbilllist with amounts that are
  'not generated by the automatic distribution such as when
  'an existing edit is brought up or when the user overrides
  'automatically distributed amounts
  If BillCnt = 0 Then Exit Sub
  ReDim PaySeq(1 To BillCnt, 1 To 7) As Double
  ReDim Preserve TempBillList(1 To BillCnt) As PayListType
  TempBillListCnt = BillCnt
  'BillTrans are in oldest first order
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To BillCnt
    If BillTrans(x) <= 0 Then
      Call DisablePayFields
      OverPay = True
      GoTo OverPay
    End If
    Get THandle, BillTrans(x), TaxTrans
OverPay:
      TempBillList(x).BillRec = BillTrans(x)
      TempBillList(x).CustRec = GCustNum
      TempBillList(x).TaxYear = TaxTrans.TaxYear
      PaySeq(x, 1) = OldRound(PaySeq(x, 1) + TaxTrans.Revenue.Interest)
      PaySeq(x, 1) = OldRound(PaySeq(x, 1) - TaxTrans.Revenue.InterestPd)
      PaySeq(x, 2) = OldRound(PaySeq(x, 2) + TaxTrans.Revenue.Collection)
      PaySeq(x, 2) = OldRound(PaySeq(x, 2) - TaxTrans.Revenue.CollectionPd)
      PaySeq(x, 3) = OldRound(PaySeq(x, 3) + TaxTrans.Revenue.LateList)
      PaySeq(x, 3) = OldRound(PaySeq(x, 3) - TaxTrans.Revenue.LateListPd)
      PaySeq(x, 4) = OldRound(PaySeq(x, 4) + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
      PaySeq(x, 4) = OldRound(PaySeq(x, 4) + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      PaySeq(x, 4) = OldRound(PaySeq(x, 4) - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
      PaySeq(x, 4) = OldRound(PaySeq(x, 4) - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      PaySeq(x, 5) = OldRound(PaySeq(x, 5) + TaxTrans.Revenue.RevOpt1)
      PaySeq(x, 5) = OldRound(PaySeq(x, 5) - TaxTrans.Revenue.RevOpt1Pd)
      PaySeq(x, 6) = OldRound(PaySeq(x, 6) + TaxTrans.Revenue.RevOpt2)
      PaySeq(x, 6) = OldRound(PaySeq(x, 6) - TaxTrans.Revenue.RevOpt2Pd)
      PaySeq(x, 7) = OldRound(PaySeq(x, 7) + TaxTrans.Revenue.RevOpt3)
      PaySeq(x, 7) = OldRound(PaySeq(x, 7) - TaxTrans.Revenue.RevOpt3Pd)
   Next x
   
   WhatsLeft = CDbl(fpCurrIntPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 1) Then
       TempBillList(x).Interest1 = PaySeq(x, 1)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest1)
     Else
       TempBillList(x).Interest1 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest1)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 1))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   WhatsLeft = CDbl(fpCurrAdvColPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 2) Then
       TempBillList(x).Collection = PaySeq(x, 2)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Collection)
     Else
'       fpCurrAdvColPaid.Value = CDbl(fpCurrAdvColPaid.Value) + WhatsLeft
       TempBillList(x).Collection = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Collection)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 2))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrLateListPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 3) Then
       TempBillList(x).LateList = PaySeq(x, 3)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).LateList)
     Else
       TempBillList(x).LateList = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).LateList)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 3))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrPrincPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 4) Then
       TempBillList(x).Principle1 = PaySeq(x, 4)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Principle1)
     Else
       TempBillList(x).Principle1 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Principle1)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 4))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrRevOpt1Paid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 5) Then
       TempBillList(x).OptRev1 = PaySeq(x, 5)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev1)
     Else
       TempBillList(x).OptRev1 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev1)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 5))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrRevOpt2Paid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 6) Then
       TempBillList(x).OptRev2 = PaySeq(x, 6)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev2)
     Else
       TempBillList(x).OptRev2 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev2)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 6))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrRevOpt3Paid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 7) Then
       TempBillList(x).OptRev3 = PaySeq(x, 7)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev3)
     Else
       TempBillList(x).OptRev3 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).OptRev3)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 7))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "LoadTempPayList", Erl)
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

Private Sub DisablePayFields()
  fpCurrPrincPaid.Enabled = False
  fpCurrIntPaid.Enabled = False
  fpCurrAdvColPaid.Enabled = False
  fpCurrLateListPaid.Enabled = False
  fpCurrRevOpt1Paid.Enabled = False
  fpCurrRevOpt2Paid.Enabled = False
  fpCurrRevOpt3Paid.Enabled = False
End Sub

Private Sub EnablePayFields()
  fpCurrPrincPaid.Enabled = True
  fpCurrIntPaid.Enabled = True
  fpCurrAdvColPaid.Enabled = True
  fpCurrLateListPaid.Enabled = True
  fpCurrRevOpt1Paid.Enabled = True
  fpCurrRevOpt2Paid.Enabled = True
  fpCurrRevOpt3Paid.Enabled = True
End Sub

Private Sub ClearTemps()
   TempPayDate = Date2Num(PayDate) '2/14/06
   TempAmtOwed = 0
   TempTenderTY = 0
   TempCashAmt = 0
   TempChkAmt = 0
   TempChrgAmt = 0
   TempAmtRecd = 0
   TempChange = 0
   TempDesc = 0
   TempPaidOwed1AmtOwed = 0
   TempPaidOwed2AmtOwed = 0
   TempPaidOwed3AmtOwed = 0
   TempPaidOwed4AmtOwed = 0
   TempPaidOwed5AmtOwed = 0
   TempPaidOwed6AmtOwed = 0
   TempPaidOwed7AmtOwed = 0
   TempPaidOwed1AmtPaid = 0
   TempPaidOwed2AmtPaid = 0
   TempPaidOwed3AmtPaid = 0
   TempPaidOwed4AmtPaid = 0
   TempPaidOwed5AmtPaid = 0
   TempPaidOwed6AmtPaid = 0
   TempPaidOwed7AmtPaid = 0
   TempTotOwed = 0
   TempAmtPaid = 0
   TempTotPaid = 0
   TempPrincPaid = 0
   TempIntPaid = 0
   TempAdvColPaid = 0
   TempLateListPaid = 0
   TempRev1Paid = 0
   TempRev2Paid = 0
   TempRev3Paid = 0
   TempDisc = 0
   TempTotPd = 0
   TempPrePay = 0
   TempAcctNum = 0
End Sub

Private Sub ReassignDiscount()
  Dim x As Integer
  Dim DiscEntered As Double
  
  DiscEntered = CDbl(fpCurrDisc.Value)
  ReDim ThisPct(1 To BillCnt) As Double
  Call GetMaxDisc
'  fpCurrPrincPaid = fpCurrPrincPaid
  For x = 1 To BillCnt
    If TempBillList(x).DiscAmt > 0 Then
      ThisPct(x) = OldRound(TempBillList(x).DiscAmt / MaxDisc)
    Else
      ThisPct(x) = 0
    End If
  Next x
  
  For x = 1 To BillCnt
    TempBillList(x).DiscAmt = OldRound(ThisPct(x) * DiscEntered)
  Next x
  
End Sub

Private Function AllTaggedPaid() As Boolean
  Dim x As Integer
  Dim ThisTot#
  
  AllTaggedPaid = True
  
  For x = 1 To BillCnt
    If TempBillList(x).TotOwed > OldRound(TempBillList(x).TotPaid + TempBillList(x).DiscAmt) Then
      AllTaggedPaid = False
      Exit For
    End If
  Next x
  
End Function

Private Sub ClearPaidFields()
  fpCurrPrePay = 0
  fpCurrPrincPaid = 0
  fpCurrAdvColPaid = 0
  fpCurrLateListPaid = 0
  fpCurrRevOpt1Paid = 0
  fpCurrRevOpt2Paid = 0
  fpCurrRevOpt3Paid = 0
  fpCurrTotPaid = 0
  fpCurrTotWDisc = 0
'  fpCurrDisc = 0
End Sub

Private Sub ResetLeaveName()
  Call ClearPaidFields
  fpCurrAmtOwed = 0
  fpCurrTotRecd = 0
  fpCurrCashPd = 0
  fpCurrChkChrgPd = 0
  fpcmbTenderType.Text = "CASH"
  fpCurrChngDue = 0
  fpCurrPrincOwed = 0
  fpCurrPrincPaid = 0
  fpCurrIntOwed = 0
  fpCurrIntPaid = 0
  fpCurrAdvColOwed = 0
  fpCurrAdvColPaid = 0
  fpCurrLateListOwed = 0
  fpCurrLateListPaid = 0
  fpCurrRevOpt1Owed = 0
  fpCurrRevOpt1Paid = 0
  fpCurrRevOpt2Owed = 0
  fpCurrRevOpt2Paid = 0
  fpCurrRevOpt3Owed = 0
  fpCurrRevOpt3Paid = 0
  fpCurrTotOwed = 0
  TempPayDate = Date2Num(PayDate) '2/14/06
  TempAmtOwed = 0
  TempTenderTY = "CASH"
  TempCashAmt = 0
  TempChkAmt = 0
  TempChrgAmt = 0
  TempAmtRecd = 0
  TempChange = 0
  TempDesc = ""
  TempPaidOwed1AmtOwed = 0
  TempPaidOwed2AmtOwed = 0
  TempPaidOwed3AmtOwed = 0
  TempPaidOwed4AmtOwed = 0
  TempPaidOwed5AmtOwed = 0
  TempPaidOwed6AmtOwed = 0
  TempPaidOwed7AmtOwed = 0
  TempPaidOwed1AmtPaid = 0
  TempPaidOwed2AmtPaid = 0
  TempPaidOwed3AmtPaid = 0
  TempPaidOwed4AmtPaid = 0
  TempPaidOwed5AmtPaid = 0
  TempPaidOwed6AmtPaid = 0
  TempPaidOwed7AmtPaid = 0
  TempTotOwed = 0
  TempAmtPaid = 0
  TempTotPaid = 0
End Sub


Public Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdInfo.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdInfo.ForeColor = &H80000012
      cmdInfo.FontSize = BtnFnt
    Case 2
      cmdInfo.ForeColor = &H80000011
      cmdInfo.FontSize = BtnFnt - 0.7
    Case 3
      cmdInfo.ForeColor = &H80000011
      cmdInfo.FontSize = BtnFnt - 1.4
    Case 4
      cmdInfo.ForeColor = &H80000010
      cmdInfo.FontSize = BtnFnt - 2.1
    Case 5
      cmdInfo.ForeColor = &H80000010
      cmdInfo.FontSize = BtnFnt - 2.8
    Case 6
      cmdInfo.ForeColor = &H8000000F
      cmdInfo.FontSize = BtnFnt - 3.5
    Case 7
      cmdInfo.ForeColor = &H8000000F
      cmdInfo.FontSize = BtnFnt - 4.2
    Case 8
      cmdInfo.ForeColor = &H8000000E
      cmdInfo.FontSize = BtnFnt - 4.9
    Case 9
      cmdInfo.ForeColor = &H8000000E
      cmdInfo.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
''  DoEvents
End Sub

Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  
  'on error goto ERRORSTUFF
  
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
'  If Exist("C:\RcptPrn.dat") Then
'    Open "c:\RcptPrn.dat" For Random Shared As RP1 Len = lenRP
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP '2/14/08
    Get RP1, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
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
      RctValidate = True
    Else
      RctValidate = False
    End If
  Close RP1
  Else
    RecpDef = 99
  End If
Exit Sub
nofound:
  RecpDef = 99
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "GetRcpInfo", Erl)
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

Private Function Check4ValidPaidEntries() As Boolean '8/12/05

  'on error goto ERRORSTUFF
  
  Check4ValidPaidEntries = True
  If CDbl(fpCurrPrincPaid.Value) > CDbl(fpCurrPrincOwed.Value) Then
    Call TaxMsg(800, "The amount entered for principal paid cannot be more than the amount owed for principal. Press F9 to correct this problem.")
    fpCurrPrincPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
    Call TaxMsg(800, "The amount entered for interest paid cannot be more than the amount owed for interest. Press F9 to correct this problem.")
    fpCurrIntPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrAdvColPaid.Value) > CDbl(fpCurrAdvColOwed.Value) Then
    Call TaxMsg(800, "The amount entered for adv/collect paid cannot be more than the amount owed for adv/collect. Press F9 to correct this problem.")
    fpCurrAdvColPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrLateListPaid.Value) > CDbl(fpCurrLateListOwed.Value) Then
    Call TaxMsg(800, "The amount entered for late listing paid cannot be more than the amount owed for late listing. Press F9 to correct this problem.")
    fpCurrLateListPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrRevOpt1Paid.Value) > CDbl(fpCurrRevOpt1Owed.Value) Then
    Call TaxMsg(800, "The amount entered for " + Opt1Desc + " paid cannot be more than the amount owed for " + Opt1Desc + ". Press F9 to correct this problem.")
    fpCurrRevOpt1Paid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrRevOpt2Paid.Value) > CDbl(fpCurrRevOpt2Owed.Value) Then
    Call TaxMsg(800, "The amount entered for " + Opt2Desc + " paid cannot be more than the amount owed for " + Opt2Desc + ". Press F9 to correct this problem.")
    fpCurrRevOpt2Paid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrRevOpt3Paid.Value) > CDbl(fpCurrRevOpt3Owed.Value) Then
    Call TaxMsg(800, "The amount entered for " + Opt3Desc + " paid cannot be more than the amount owed for " + Opt3Desc + ". Press F9 to correct this problem.")
    fpCurrRevOpt3Paid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPaymentEntry", "Check4ValidPaidEntries", Erl)
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


End Function
Private Function VerifyPayList() As Boolean
  Dim PayRec As PayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer, y As Integer
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim ThisRec As Long, NextRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  VerifyPayList = False
  If fpCurrAmtOwed.Value = 0 Then 'no bill has been selected and is probably a pre pay
    VerifyPayList = True
    Exit Function
  End If
  If GCustNum = 0 Then Exit Function
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  NextRec = TaxCust.LastTrans
  Close TCHandle
  
  OpenTaxTransFile THandle, NumOfTRecs
  If EditFlag = False Then GoTo NewPayment
  OpenPayListFile PHandle, OperNum
  NumOfPRecs = LOF(PHandle) / Len(PayRec)
  Do While NextRec > 0
    Get THandle, NextRec, TaxTrans
    For x = 1 To NumOfPRecs 'number of bills tagged
      Get PHandle, x, PayRec
      If PayRec.CustRec = GCustNum Then
        For y = 1 To BillCnt
          If PayRec.BillRec = NextRec Then
            VerifyPayList = True
          End If
        Next y
      End If
    Next x
    NextRec = TaxTrans.LastTrans
  Loop
  GoTo Done
  
NewPayment:
  Do While NextRec > 0
    Get THandle, NextRec, TaxTrans
      For y = 1 To BillCnt
        If TempBillList(y).BillRec = NextRec Then
          VerifyPayList = True
        End If
      Next y
    NextRec = TaxTrans.LastTrans
  Loop

Done:
  Close PHandle
  Close THandle
  
End Function

