VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxReprintPosted 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Reprints of Posted Bills"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxReprintPosted.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbFile 
      Height          =   405
      Left            =   4395
      TabIndex        =   1
      Top             =   1680
      Width           =   4410
      _Version        =   196608
      _ExtentX        =   7779
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
      BackColor       =   16777215
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
      ColDesigner     =   "frmTaxReprintPosted.frx":08CA
   End
   Begin LpLib.fpList fpList1 
      Height          =   3120
      Left            =   1560
      TabIndex        =   5
      Top             =   4200
      Width           =   8535
      _Version        =   196608
      _ExtentX        =   15055
      _ExtentY        =   5503
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Columns         =   4
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmTaxReprintPosted.frx":0C6D
   End
   Begin VB.OptionButton OptMulti 
      BackColor       =   &H008F8265&
      Caption         =   "Multi-Select From Bill List"
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
      Height          =   372
      Left            =   3480
      TabIndex        =   15
      ToolTipText     =   $"frmTaxReprintPosted.frx":1081
      Top             =   3360
      Width           =   3372
   End
   Begin VB.OptionButton OptRange 
      BackColor       =   &H008F8265&
      Caption         =   "Range Of Bills"
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
      Height          =   372
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "The program will print all bills from the 'First Bill' selection to the 'Last Bill' selection."
      Top             =   2880
      Width           =   1932
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmTaxReprintPosted.frx":111F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmTaxReprintPosted.frx":12FE
   End
   Begin EditLib.fpDoubleSingle fpDblSnglLastBill 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   2850
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
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
      DecimalPlaces   =   -1
      DecimalPoint    =   ""
      FixedPoint      =   0   'False
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "0"
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
   Begin EditLib.fpDoubleSingle fpDblSnglFirstBill 
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   2850
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
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
      DecimalPlaces   =   -1
      DecimalPoint    =   ""
      FixedPoint      =   0   'False
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "0"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmTaxReprintPosted.frx":14DA
   End
   Begin EditLib.fpText fptxtCurrForm 
      Height          =   396
      Left            =   5538
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Late notices are selected on the System Setup screen."
      Top             =   1200
      Width           =   2856
      _Version        =   196608
      _ExtentX        =   5038
      _ExtentY        =   698
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
      AlignTextH      =   1
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpDateTime fptxtPostDate 
      Height          =   375
      Left            =   6045
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1005
      _Version        =   196608
      _ExtentX        =   1773
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
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0D0D0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year:"
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
      Height          =   345
      Left            =   4605
      TabIndex        =   19
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Format In Use:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3258
      TabIndex        =   17
      Top             =   1260
      Width           =   2028
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1095
      Left            =   480
      Top             =   2760
      Width           =   10695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0D0D0&
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
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
      Height          =   345
      Left            =   3000
      TabIndex        =   14
      Top             =   3390
      Width           =   300
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0D0D0&
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
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
      Height          =   345
      Left            =   3000
      TabIndex        =   13
      Top             =   2910
      Width           =   300
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select One Of The Following Options:"
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
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Bill:"
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
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Bill:"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   2910
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3375
      Left            =   1440
      Top             =   4080
      Width           =   8775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0D0D0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select File:"
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
      Height          =   345
      Left            =   2835
      TabIndex        =   4
      Top             =   1770
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Reprinting of Posted Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3120
      TabIndex        =   0
      Top             =   516
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   348
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   240
      Width           =   8652
   End
End
Attribute VB_Name = "frmTaxReprintPosted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim FirstLoad As Boolean
  Dim FirstNum As Long
  Dim LastNum As Long
  Dim DirContents() As String
  Dim DirCnt As Integer
  Dim MyPath$
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim TownName As String
  Dim BillIdx() As Long
  Dim BillCnt As Long
  
Private Sub cmdExit_Click()
  frmTaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
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
  FirstLoad = True
  Me.HelpContextID = hlpReprintPostedTax
  Call LoadMe
  FirstLoad = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxReprintPosted.")
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

Private Sub LoadMe()
  Dim x As Integer
'  Dim ThisFile As Integer
'  Dim ThisFile$
  Dim GotIt As Boolean
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MyName$
  
  'on error goto ERRORSTUFF
  DirCnt = 0
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    MyName = Dir
    If Len(MyName) > 4 Then
      DirCnt = DirCnt + 1
      ReDim Preserve DirContents(DirCnt) As String
      DirContents(DirCnt) = MyPath + MyName
      If DirCnt = 1 Then
        fpcmbFile.Text = MyName
      End If
       fpcmbFile.AddItem MyName
    End If
  Loop
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.City)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20007 Then
    cmdAlign.Enabled = False
  End If
  
  Select Case TaxMasterRec.TaxForm
    Case 21837
      fptxtCurrForm.Text = "MULTI-PART"
    Case 20304
      fptxtCurrForm.Text = "POSTCARD"
    Case 16716
      fptxtCurrForm.Text = "LASER"
    Case 29999
      fptxtCurrForm.Text = "EXPORT COMBINED"
    Case 20000
      fptxtCurrForm.Text = "EXPORT REAL"
    Case 20001
      fptxtCurrForm.Text = "EXPORT PERSONAL"
    Case 20002
      fptxtCurrForm.Text = "HMLT24TF"
    Case 20003
      fptxtCurrForm.Text = "PH24TF"
    Case 20004
      fptxtCurrForm.Text = "SYL23TF"
    Case 20005
      fptxtCurrForm.Text = "BSC32TF"
    Case 20006
      fptxtCurrForm.Text = "LLN21TF"
    Case 20007
      fptxtCurrForm.Text = "LASER LEGAL"
    Case 20008
      fptxtCurrForm.Text = "LASER LEGAL HP"
    Case Else
      fptxtCurrForm.Text = "UNKNOWN"
  End Select
  
  Call LoadList
  OptRange.Value = True
'  fpList1.Enabled = False
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPOsted", "LoadMe", Erl)
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

Private Sub LoadList()
  Dim ThisFile$
  Dim THandle As Integer
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim FirstBill$
  Dim LastBill$
  Dim ThisLen As Integer
  Dim ThisName As String * 35
  Dim BillArr() As Long
  Dim BillArrCnt As Long
  Dim BillNum() As Long
  Dim BigNum As Long
  Dim BigNumStatic As Long
  Dim Nextx As Long
  Dim HoldThis As Long
  Dim Thisx As Long
  
  'on error goto ERRORSTUFF
  
  FirstBill$ = ""
  LastBill$ = ""
  fpList1.Clear
  ThisFile = MyPath + QPTrim$(fpcmbFile.Text)
  If ThisFile = "" Then
    Call TaxMsg(900, "Please make a selection from the list of files.")
    fpcmbFile.SetFocus
    Exit Sub
  End If
  
  OpenPostedReprintFile THandle, NumOfTBRecs, ThisFile
  
  
  If NumOfTBRecs = 0 Then
    Call TaxMsg(900, "No records available for " + ThisFile)
    Close
    Exit Sub
  End If
  
  GoSub FillIdx

  For x = 1 To BillArrCnt '8/6/08
    Get THandle, BillIdx(x), TaxBill '8/6/08
'  For x = 1 To NumOfTBRecs
'    Get THandle, x, TaxBill
    If TaxBill.BillNumber >= 0 Then
      If FirstBill = "" Then FirstBill = CStr(TaxBill.BillNumber)
      FirstNum = FirstBill
      ThisName = QPTrim$(TaxBill.CustName)
'      fpList1.InsertRow = CStr(TaxBill.BillNumber) + Chr(9) + ThisName + Chr(9) + Using$("$###,###,##0.00", TaxBill.TotalBillDue) + Chr(9) + CStr(x)
      fpList1.InsertRow = CStr(TaxBill.BillNumber) + Chr(9) + ThisName + Chr(9) + Using$("$###,###,##0.00", TaxBill.TotalBillDue) + Chr(9) + CStr(BillIdx(x)) '8/6/08
    End If
  Next x
  
  For x = BillArrCnt To 1 Step -1 '8/6/08
    Get THandle, BillIdx(x), TaxBill '8/6/08
'  For x = NumOfTBRecs To 1 Step -1
'    Get THandle, x, TaxBill
    If TaxBill.BillNumber >= 0 Then
      LastBill = CStr(TaxBill.BillNumber)
      LastNum = LastBill
      Exit For
    End If
  Next x
  
  Close
  fptxtPostDate.Text = CStr(TaxBill.TaxYear)
  fpDblSnglFirstBill = CLng(FirstBill)
  fpDblSnglLastBill = CLng(LastBill)
  
  Exit Sub
  
FillIdx: '8/6/08
  BigNum = 0
  For x = 1 To NumOfTBRecs
    Get THandle, x, TaxBill
    If TaxBill.BillNumber > 0 Then
      BillArrCnt = BillArrCnt + 1
      ReDim Preserve BillArr(1 To BillArrCnt) As Long
      ReDim Preserve BillNum(1 To BillArrCnt) As Long
      BillArr(BillArrCnt) = x
      BillNum(BillArrCnt) = TaxBill.BillNumber
      If TaxBill.BillNumber > BigNum Then
        BigNum = TaxBill.BillNumber
      End If
    End If
  Next x
  
  BigNumStatic = BigNum + 1
  BigNum = BigNumStatic
  BillCnt = BillArrCnt 'global
  ReDim BillIdx(1 To BillCnt) As Long 'global
  Nextx = 1
  Do
    For x = Nextx To BillArrCnt
      If BillNum(x) < BigNum Then
        BigNum = BillNum(x)
        Thisx = x
      End If
    Next x
    HoldThis = BillNum(Nextx)
    BillNum(Nextx) = BillNum(Thisx)
    BillNum(Thisx) = HoldThis
    HoldThis = BillArr(Nextx)
    BillArr(Nextx) = BillArr(Thisx)
    BillArr(Thisx) = HoldThis
    
    If Nextx = BillArrCnt Then Exit Do
    Nextx = Nextx + 1
    BigNum = BigNumStatic
  Loop
  
  For x = 1 To BillArrCnt
    BillIdx(x) = BillArr(x)
  Next x
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPosted", "LoadList", Erl)
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

Private Sub PrintBills(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim DueDate$
  Dim TAXRATE#
  Dim NetTaxVal#
  Dim LC As Integer
  Dim ThisFile$
  Dim PrintCnt As Long
  
  'on error goto ERRORSTUFF
  
  ThisFile$ = QPTrim$(fpcmbFile.Text)
  DueDate$ = "12-31-" + QPTrim$(Str$(TaxBill.TaxYear))
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)

  Print #RptHandle,
  Print #RptHandle, Tab(29); "TOWN OF MAGGIE VALLEY"
  Print #RptHandle, Tab(29); "    3987 SOCO RD."
  Print #RptHandle, Tab(29); "MAGGIE VALLEY NC 28751"
  Print #RptHandle, Tab(29); "  PROPERTY TAX BILL"

  For LC = 1 To 3
    Print #RptHandle, " "
  Next

  Print #RptHandle, Tab(12); "ACCT # "; TaxBill.CustRec;
  Print #RptHandle, Tab(65); "BILL #"; Using("#####0", TaxBill.BillNumber)
  Print #RptHandle, Tab(12); Left$(CustName$, 25);
  Print #RptHandle, Tab(63); "TAX YEAR "; TaxBill.TaxYear
  Print #RptHandle, Tab(12); Left$(TaxBill.CustAdd1, 25);
  Print #RptHandle, Tab(63); "TAX RATE "; Using("##0.00", TAXRATE#)
  Print #RptHandle, Tab(12); Left$(TaxBill.CustAdd2, 25)
  Print #RptHandle, Tab(12); QPTrim$(TaxBill.CustAdd3); " "; Left$(TaxBill.CustZip, 5) + "-" + Mid$(TaxBill.CustZip, 6, 4)
  For LC = 1 To 4
    Print #RptHandle, " "
  Next
  Print #RptHandle, Tab(39); "[--------- VALUATIONS --------]"
  Print #RptHandle, Tab(2); "PROPERTY DESCRIPTION"; Tab(30); "RATE"; Tab(40); "REAL"; Tab(48); "PERSONAL"; Tab(61); "EXEMPT"; Tab(72); "TOTAL"
  Print #RptHandle, " "
  'Line 23 Starts Here
  Print #RptHandle, Tab(30); Using(".##", TAXRATE#);
  Print #RptHandle, Tab(35); Using("###,###,##0", TaxBill.RealValue);
  Print #RptHandle, Tab(47); Using("###,###,##0", TaxBill.PersValue);
  Print #RptHandle, Tab(59); Using("###,##0", TaxBill.ExptValue);
  Print #RptHandle, Tab(68); Using("###,###,##0", (TaxBill.PersValue + TaxBill.RealValue))
  Print #RptHandle, Tab(2); QPTrim$(TaxBill.RDesc1)
  Print #RptHandle, Tab(2); QPTrim$(PINTemp)


   Print #RptHandle, ""
   Print #RptHandle, ""
   Print #RptHandle, ""
   Print #RptHandle, Tab(2); "NOTE: "
   Print #RptHandle, Tab(2); "      A 2% PENALTY WILL BE ADDED AFTER DUE DATE."
   Print #RptHandle, Tab(2); "      .75% ADDED ON FIRST OF EACH MONTH THEREAFTER."
   Print #RptHandle, ""
   Print #RptHandle, Tab(2); "      PLEASE SUBMIT YOUR ACCOUNT# OR, PARCEL ID# ON CHECK"
   Print #RptHandle, Tab(2); "      TO PROCESS YOUR PAYMENT. CONTACT THE TAX COLLECTOR"
   Print #RptHandle, Tab(2); "      IF YOU HAVE IF ANY QUESTIONS."
   Print #RptHandle, Tab(2); "      (PHONE: 828-926-0866  EXT. 101)"
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle, Tab(39); " TAX DUE DATE: "; DueDate$
   Print #RptHandle,
   Print #RptHandle, Tab(39); "TOTAL TAX DUE: "; Using("###,###,##0", TaxBill.TotalBillDue)
   Print #RptHandle,
   Print #RptHandle,
   Print #RptHandle, "BN"; Using("#####0", PrnCnt)
   Print #RptHandle, Chr$(12);
   
   Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPosted", "PrintBills", Erl)
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
 
  
End Sub

Private Sub PrintLaser1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$
  Dim TBDRec As TxBill1DefaultsType
  Dim TBDHandle As Integer
  Dim FBill&, PrnCnt&
  Dim LBill&
  Dim PCnt As Integer
  Dim NCnt As Integer
  Dim ThisRate As Double
  Dim ThisFile$
  Dim PrintCnt As Long
  
  'on error goto ERRORSTUFF
  
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  dlm$ = "~"
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  ReportFile$ = StartPath$ + "/TaxBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  OpenTxBill1File TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  ARptTempTaxBill.Head1 = QPTrim(TBDRec.TxtHead1)
  ARptTempTaxBill.Head2 = QPTrim(TBDRec.TxtHead2)
  ARptTempTaxBill.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  ARptTempTaxBill.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  ARptTempTaxBill.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  ARptTempTaxBill.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  ARptTempTaxBill.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  ARptTempTaxBill.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  ARptTempTaxBill.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  ARptTempTaxBill.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  ARptTempTaxBill.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  ARptTempTaxBill.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  ARptTempTaxBill.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  ARptTempTaxBill.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  ARptTempTaxBill.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  ARptTempTaxBill.LblHead4 = QPTrim(TBDRec.txtHead4)
  ARptTempTaxBill.LblHead5 = QPTrim(TBDRec.txtHead5)
  ARptTempTaxBill.LblHead6 = QPTrim(TBDRec.txtHead6)
  ARptTempTaxBill.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  ARptTempTaxBill.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      ARptTempTaxBill.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      ARptTempTaxBill.Image1.Visible = True
    End If
  End If
  
  frmTaxShowPctComp.Label1 = "Printing Tax Bills"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  OpenPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
  NumOfTBRecs = BillCnt '8/6/08
  If OptMulti.Value = True Then
    ReDim PrintThis(1 To 1) As Long
    For x = 0 To fpList1.ListCount - 1
      fpList1.Row = x
      If fpList1.Selected = True Then
        fpList1.ListIndex = x
        PrintCnt = PrintCnt + 1
        ReDim Preserve PrintThis(1 To PrintCnt) As Long
        fpList1.Col = 3
        PrintThis(PrintCnt) = CLng(fpList1.ColText)
      End If
    Next x
    NumOfTBRecs = PrintCnt
  End If
      
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TBRec
      Get TBHandle, BillIdx(x), TBRec '8/6/08
    Else
      Get TBHandle, PrintThis(x), TBRec
      GoTo PrintIt '8/21/07
    End If
      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
PrintIt: '8/21/07
        If TBRec.TotalBillDue > 0 Then
          '                       0                           1
          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
          '                           2                             3
          Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
          '                          4                         5
          Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
          '                         6                              7
          Print #RptHandle, QPTrim$(TBRec.RealPin); dlm; QPTrim$(TBRec.RDesc1); dlm;
          '                        8                     9                     10
          Print #RptHandle, TBRec.RealValue; dlm; TBRec.PersValue; dlm; TBRec.ExptValue; dlm;
          If TBRec.RealTaxDue > 0 And TBRec.PersTaxDue > 0 Then
            ThisRate = TBRec.RealTaxRate
          ElseIf TBRec.RealTaxDue <= 0 And TBRec.PersTaxDue > 0 Then
            ThisRate = TBRec.PersTaxRate
          ElseIf TBRec.RealTaxDue > 0 And TBRec.PersTaxDue <= 0 Then
            ThisRate = TBRec.RealTaxRate
          Else
            ThisRate = 0
          End If
          '                                   11                                12
          Print #RptHandle, OldRound(TBRec.RealValue + TBRec.PersValue); dlm; ThisRate; dlm;
          If TBRec.OverPayAmt > 0 Then
            '                                        13                             14            15
            Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm; ""; dlm; TBDRec.dologo; dlm;
            '                         16                   17                     18
            Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
            '                    19             20             21                 22                          23                    24
            Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; QPTrim$(TBRec.CustZip); dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; "False" '  TBRec.OverPayAmt
          Else
            '                        13                14             15
            Print #RptHandle, TBRec.TotalBillDue; dlm; ""; dlm; TBDRec.dologo; dlm;
            '                         16                   17                     18
            Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
          '                    19             20             21                 22                          23                     24
          Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; QPTrim$(TBRec.CustZip); dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; "False" ' TBRec.OverPayAmt
          End If
          PCnt = PCnt + 1
        End If
      End If
    frmTaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Close TBHandle
  
  Unload frmTaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  ARptTempTaxBill.GetName ReportFile$
  ARptTempTaxBill.Show

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPosted", "PrintLaser1", Erl)
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
 
End Sub

Private Sub PrintStandard(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim TAXRATE#
  Dim NetTaxVal#
  
  'on error goto ERRORSTUFF
  
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  Print #RptHandle, "~"; Tab(50); Using("###0", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, TaxBill.TaxYear;
  Print #RptHandle, Tab(7); Using("######", TaxBill.BillNumber);
  Print #RptHandle, Tab(16); Using("#####", TaxBill.CustRec);
  Print #RptHandle, Tab(23); QPTrim$(PINTemp); Tab(37); Using("####", TaxBill.TaxYear);

  Print #RptHandle, Tab(42); Using("########", TaxBill.CustRec);
  'PRINT #RptHandle, TAB(49); USING "######,#.##"; TaxBill.TotalBillDue + TaxBill.PriorYrBalance
  Print #RptHandle, Tab(51); Using("######", TaxBill.CustRec)
  Print #RptHandle,
  Print #RptHandle, Tab(11); Left$(QPTrim$(CustName$), 21)
  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc1), 21)
  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc2), 21)
  'v line 12
  Print #RptHandle,
  
  Print #RptHandle, Using("###,###,##0", TaxBill.RealValue); Tab(15); TaxBill.PersValue;
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle, Tab(25); Using("#,###,##0", TaxBill.ExptValue);
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(36); Left$(QPTrim$(CustName$), 25)
  Print #RptHandle, Using("###,###,##0", NetTaxVal#);
  Print #RptHandle, Tab(16); Using("#.##", TAXRATE#);
  Print #RptHandle, Tab(21); Using("##,###,##0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(36); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
  Print #RptHandle, Tab(21); Using("##,###,##0.00", TaxBill.LateTaxDue)
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using("###,##0.0", TaxBill.PriorYrBalance);
  If TaxBill.OverPayAmt > 0 Then
    Print #RptHandle, Tab(21); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance); Tab(47); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance - TaxBill.OverPayAmt)
  Else
    Print #RptHandle, Tab(21); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance); Tab(47); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance)
  End If
  If TaxBill.OverPayAmt > 0 Then
    Print #RptHandle, "Credit of " + QPTrim$(Using$("$###,##0.00", TaxBill.OverPayAmt)) + " has been applied."
  Else
    Print #RptHandle,
  End If
  Print #RptHandle, "~"

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPosted", "PrintStandard", Erl)
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
 
End Sub
Private Sub cmdProcess_Click()
  Dim RptHandle As Integer
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim WhatRec&, PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim RptFile$, FBill&, LBill&
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisFile$, FF$
  Dim PrintCnt As Long
  
  'on error goto ERRORSTUFF
  
  FF$ = Chr(12)
  Select Case QPTrim$(fptxtCurrForm.Text)
    Case "EXPORT COMBINED"
      Call TaxMsg(900, "There are no reprints possible when the bill format is " + "EXPORT COMBINED.")
      Exit Sub
    Case "EXPORT REAL"
      Call TaxMsg(900, "There are no reprints possible when the bill format is " + "EXPORT REAL.")
      Exit Sub
    Case "EXPORT PERSONAL"
      Call TaxMsg(900, "There are no reprints possible when the bill format is " + "EXPORT PERSONAL.")
      Exit Sub
    Case "UNKNOWN"
      Call TaxMsg(900, "There are no reprints possible when the bill format is " + "UNKNOWN.")
      Exit Sub
    Case Else
  End Select
  
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  LBill = fpDblSnglLastBill.Value
  FBill = fpDblSnglFirstBill.Value
  
  If OptRange.Value = True Then
    If LBill < FBill Then
      Call TaxMsg(900, "The last bill number comes before the first bill number. Please correct this error.")
      fpDblSnglLastBill.SetFocus
      Close
      Exit Sub
    End If
  End If
    
  If FBill < FirstNum Then
    Call TaxMsg(900, "The first bill cannot be less than " + CStr(FirstNum) + ". Please re-enter and try again.")
    fpDblSnglFirstBill.SetFocus
    Exit Sub
  End If
  
  If LBill > LastNum Then
    Call TaxMsg(900, "The last bill cannot be greater than " + CStr(LastNum) + ". Please re-enter and try again.")
    fpDblSnglLastBill.SetFocus
    Exit Sub
  End If
  
  If TaxMasterRec.TaxForm = 16716 Then
    Call PrintLaser1
    Exit Sub
  End If
  
  If TaxMasterRec.TaxForm = 20007 Then
    Call PrintLaserLegal
    Exit Sub
  End If
  
  If TaxMasterRec.TaxForm = 20008 Then
    Call PrintLaserLegalHP
    Exit Sub
  End If
  
  RptHandle = FreeFile
  RptFile$ = "TAXBLRE.PRN"
  
  Open RptFile For Output As RptHandle
  
  OpenPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
  NumOfTBRecs = BillCnt '8/6/08

  PrnCnt = 0 ' 8/21/07fpDblSnglFirstBill
  
  If fptxtCurrForm.Text = "LLN21TF" Then 'new for 10/24/06
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  ElseIf fptxtCurrForm.Text = "HMLT24TF" Then
    Call TaxMsg(800, "12 Pitch is recommended for this form. Please be sure to always start on bill #1 and not on bill #2.")
  ElseIf fptxtCurrForm.Text = "PH24TF" Then
    Call TaxMsg(900, "10 Pitch is recommended for this form.")
  ElseIf fptxtCurrForm.Text = "SYL23TF" Then
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  ElseIf fptxtCurrForm.Text = "BSC32TF" Then
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  ElseIf fptxtCurrForm.Text = "POSTCARD" Then
    Call TaxMsg(900, "12 Pitch is recommended for this form.")
  ElseIf fptxtCurrForm.Text = "MULTI-PART" Then
'    Call PrintStandard(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
  Else
    Call TaxMsg(900, "The bill format is not recognized. Printing aborted.")
    Close
    Exit Sub
  End If
  
  If OptMulti.Value = True Then
    ReDim PrintThis(1 To 1) As Long
    For x = 0 To fpList1.ListCount - 1
      fpList1.Row = x
      If fpList1.Selected = True Then
        fpList1.ListIndex = x
        PrintCnt = PrintCnt + 1
        ReDim Preserve PrintThis(1 To PrintCnt) As Long
        fpList1.Col = 3
        PrintThis(PrintCnt) = CLng(fpList1.ColText)
      End If
    Next x
    NumOfTBRecs = PrintCnt
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill '8/6/08
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoTo PrintIt '8/21/07
    End If
    If TaxBill.BillPrinted Then
      PrnCnt = PrnCnt + 1
      If PrnCnt >= FBill And PrnCnt <= LBill Then
PrintIt:
        If QPTrim$(TaxBill.RealPin) <> "" Then
          RSet PINTemp = QPTrim$(TaxBill.RealPin)
        ElseIf QPTrim$(TaxBill.PersPin) <> "" Then
          RSet PINTemp = QPTrim$(TaxBill.PersPin)
        Else
          RSet PINTemp = "0"
        End If
        CustName$ = QPTrim$(TaxBill.CustName)
        If fptxtCurrForm.Text = "LLN21TF" Then
          Call PrintLLN21TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf fptxtCurrForm.Text = "HMLT24TF" Then
          Call PrintHMLT24TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf fptxtCurrForm.Text = "PH24TF" Then
          Call PrintPH24TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf fptxtCurrForm.Text = "SYL23TF" Then
          Call PrintSYL23TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf fptxtCurrForm.Text = "BSC32TF" Then
          Call PrintBSC32TF(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf fptxtCurrForm.Text = "POSTCARD" Then
          Call PrintPostCard(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        ElseIf fptxtCurrForm.Text = "MULTI-PART" Then
          Call PrintStandard(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
        End If
      End If
    End If
  Next x
  
SkipIt:
  Close TBHandle
  Print #RptHandle, FF$
  Close RptHandle
  
  ViewPrint RptFile$, "Posted Tax Bill Reprinting", True
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPosted", "cmdProcess_Click", Erl)
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
 
End Sub

Private Sub fpcmbFile_Click()
  Call LoadList
End Sub

Private Sub OptMulti_Click()
  fpList1.MultiSelect = MultiSelectSimple
'  fpList1.Enabled = True
End Sub

Private Sub OptRange_Click()
  fpList1.MultiSelect = MultiSelectNone
'  fpList1.Enabled = False
End Sub
Private Sub cmdAlign_Click()
  Dim Handle As Integer
  Dim TempHandle As Integer
  Dim Cnt As Integer
  Dim TextLine$
  Dim BillFormat$
  
  'on error goto ERRORSTUFF
  
  BillFormat$ = fptxtCurrForm.Text
  
  If BillFormat = "POSTCARD" Then
    If Exist("TAXMSKPC1.DAT") Then
      alnRpt = "TAXMSKPC1.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TAXMSKPC1.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = "HMLT24TF" Then
    If Exist("TXMSKHMLT24TF.DAT") Then
      Call TaxMsg(900, "12 Pitch is recommended for this form. Each mask prints 2 bills to match the way the bill forms have been printed.")
      alnRpt = "TXMSKHMLT24TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKHMLT24TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = "PH24TF" Then
    If Exist("TXMSKPH24TF.DAT") Then
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      alnRpt = "TXMSKPH24TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKPH24TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = "SYL23TF" Then
    If Exist("TXMSKSYL23TF.DAT") Then
      Call TaxMsg(900, "12 Pitch is recommended for this form.")
      alnRpt = "TXMSKSYL23TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKSYL23TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = "BSC32TF" Then
    If Exist("TXMSKBSC32TF.DAT") Then
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      alnRpt = "TXMSKBSC32TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKBSC32TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = "LLN21TF" Then
    If Exist("TXMSKLLN21TF.DAT") Then
      Call TaxMsg(900, "10 Pitch is recommended for this form.")
      alnRpt = "TXMSKLLN21TF.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TXMSKLLN21TF.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  ElseIf BillFormat = "MULTI-PART" Then
    If Exist("TAXBLMSK.DAT") Then
      alnRpt = "TAXBLMSK.DAT"
    Else
      Call TaxMsg(900, "The mask for this bill format, " + "TAXBLMSK.DAT," + " could not be found.")
      Close
      Exit Sub
    End If
  Else
    Call TaxMsg(900, "The mask for this bill format could not be found.")
    Close
    Exit Sub
  End If
  
  Handle = FreeFile
  Open alnRpt For Input As #Handle
  TempHandle = FreeFile
  Open "TAXALIGN.MSK" For Output As #TempHandle
  Do While Not eof(Handle)
    Line Input #Handle, TextLine   ' Read line into variable.
    Print #TempHandle, TextLine
  Loop
  Close
  alnRpt = "TAXALIGN.MSK"
  doAlign = True
  frmTaxPrint.Show vbModal
  alnRpt = ""
  doAlign = False

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxReprintPosted", "cmdAlign_Click", Erl)
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
 

End Sub

Private Sub PrintHMLT24TF(ByVal RptHandle As Integer, TBHandle As Integer, TBRec As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Static Cnt As Integer
  
  Print #RptHandle, "~"
  Print #RptHandle, Using("#####", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(16); Using("######", TBRec.BillNumber);
  If TBRec.RealValue < -1000 Then
    TBRec.RealValue = 0
  End If
  If TBRec.PersValue < -1000 Then
    TBRec.PersValue = 0
  End If
  Print #RptHandle, Tab(48); Using("###,###,###", TBRec.RealValue)
  Print #RptHandle, Tab(16); CStr(TBRec.TaxYear);
  Print #RptHandle, Tab(48); Using("###,###,###", TBRec.PersValue)
  If TBRec.RealTaxRate# = 0 Then
    TBRec.RealTaxRate# = TBRec.PersTaxRate#
  End If
  Print #RptHandle, Tab(16); Using("#.0##", TBRec.RealTaxRate#);
  NetTaxVal# = OldRound#(TBRec.RealValue + TBRec.PersValue)
  Print #RptHandle, Tab(48); Using("###,###,###", NetTaxVal#);
  Print #RptHandle, Tab(84); Using("$#,###,##0.00", TBRec.TotalBillDue - TBRec.LateTaxDue)
  Print #RptHandle, Tab(20); Using("#####0", TBRec.CustRec);
  Print #RptHandle, Tab(55); Using("#,###,###", TBRec.ExptValue);
  Print #RptHandle, Tab(81); Using("$#,###,##0.00", TBRec.LateTaxDue)
  Print #RptHandle, Tab(14); QPTrim$(TBRec.TownShip);
  NetTaxVal# = OldRound#(NetTaxVal# - TBRec.ExptValue)
  Print #RptHandle, Tab(48); Using("###,###,###", NetTaxVal#)
  Print #RptHandle, Tab(18); QPTrim$(TBRec.LotOrAcre);
  Print #RptHandle, " "; QPTrim$(TBRec.LASize)
  If QPTrim$(TBRec.RealPin) <> "" Then
    Print #RptHandle, Tab(18); QPTrim$(TBRec.RealPin);
  Else
    Print #RptHandle, Tab(18); "";
  End If
  Print #RptHandle, Tab(86); Using("$##,###,##0.00", TBRec.TotalBillDue - TBRec.OverPayAmt) 'added OverPayAmt
  If QPTrim$(TBRec.RDesc1) <> "" Then
    Print #RptHandle, Tab(18); QPTrim$(TBRec.RDesc1)
  Else
    Print #RptHandle, ""
  End If
  If QPTrim$(TBRec.RDesc2) <> "" Then
    Print #RptHandle, Tab(18); QPTrim$(TBRec.RDesc2)
  Else
    Print #RptHandle, ""
  End If
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(8); Left$(QPTrim$(CustName$), 35)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TBRec.CustAdd1), 35)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TBRec.CustAdd2), 35)
  Print #RptHandle, Tab(8); QPTrim$(TBRec.CustAdd3); " "; QPTrim$(TBRec.CustZip)
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, ""
  Cnt = Cnt + 1
  If Cnt <> 1 Then
    Cnt = 0
    Print #RptHandle, "" '8/16/06
  End If
  
End Sub

Private Sub PrintLLN21TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#
  
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If
  
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  
  Print #RptHandle, "~"; Tab(78); "~"
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(50); TaxBill.TaxYear; Tab(59); Using("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(69); QPTrim$(TaxBill.RealPin)
  Print #RptHandle,
  Print #RptHandle, Tab(50); Using("#####0", TaxBill.CustRec)
  Print #RptHandle,
  Print #RptHandle, Tab(51); QPTrim$(Left$(TaxBill.RDesc1, 21))
  Print #RptHandle,
  Print #RptHandle, Tab(51); Using("##,###,##0", TaxBill.RealValue); Tab(61); Using("##,###,##0", TaxBill.PersValue); Tab(71); Using("###,###,##0", NetTaxVal#)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); Left$(CustName$, 25);
  Print #RptHandle, Tab(66); Using("#.00", TAXRATE#);
  Print #RptHandle, Tab(71); Using("####0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue))
  Print #RptHandle, Tab(5); Left$(TaxBill.CustAdd1, 25)
  Print #RptHandle, Tab(5); Left$(TaxBill.CustAdd2, 25);
  Print #RptHandle, Tab(71); Using("####.00", TaxBill.LateTaxDue)
  Print #RptHandle, Tab(5); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle, Tab(71); Using("####0.00", TaxBill.TotalBillDue - TaxBill.OverPayAmt) 'added OverPayAmt 8/15/06
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"; Tab(78); "~"

End Sub

Private Sub PrintPH24TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#
  
  If InStr(TownName, "WHITAKERS") Then
   If QPTrim$(TaxBill.RealPin) = "" Then GoTo EmptyPin
   RSet PINTemp = Mid(TaxBill.RealPin, Len(QPTrim$(TaxBill.RealPin)) - 3, 4)
  End If
  
EmptyPin:

  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If
  
  Print #RptHandle, "~"; Tab(40); Using("###0", PrnCnt)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Using("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(5); Using("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(15); Using("####0", TaxBill.CustRec);
  Print #RptHandle, Tab(23); QPTrim$(PINTemp); Tab(34); Using("###0", TaxBill.TaxYear);

  Print #RptHandle, Tab(38); Using("#######0", TaxBill.CustRec);
  Print #RptHandle, Tab(46); Using("#,###,##0.00", OldRound#(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle, Tab(11); Left$(CustName$, 21)
  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc1), 21) 'added QPTrim$ 10/24/06
  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc2), 21) 'added QPTrim$ 10/24/06
  'v line 12
  Print #RptHandle,
  
  Print #RptHandle, Using("###,###,##0", TaxBill.RealValue); Tab(12); Using("###,###,##0", TaxBill.PersValue);
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)

  Print #RptHandle, Tab(23); Using("#,###,##0", TaxBill.ExptValue);
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(36); Left$(CustName$, 25)
  Print #RptHandle, Using("###,###,##0", NetTaxVal#);
  Print #RptHandle, Tab(14); Using("#.000", TAXRATE#);
  Print #RptHandle, Tab(19); Using("##,###,##0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(36); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle, Tab(21); Using("#######0.00", TaxBill.LateTaxDue)
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using("###,##0.00", TaxBill.PriorYrBalance);
  Print #RptHandle, Tab(19); Using("##,###,##0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance)) '; Tab(47); Using("##,###,##0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"

End Sub

Private Sub PrintSYL23TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#
  
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If
  
  Print #RptHandle, Chr$(27); Chr$(58); "~"
  Print #RptHandle, 'added 6.23.06
  Print #RptHandle, Tab(32); Using$("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(44); Using("#.00", TAXRATE#);
  Print #RptHandle, Tab(78); Using("#####0", TaxBill.CustRec);
  Print #RptHandle, Tab(90); Using("#####0", TaxBill.BillNumber)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  Print #RptHandle, Using("#,###,##0", TaxBill.PersValue);

  Print #RptHandle, Tab(13); Using("##,###,##0", TaxBill.RealValue);

  Print #RptHandle, Tab(26); Using("##,###,##0", NetTaxVal#);
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle, Tab(37); Using("#####0.00", OldRound(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
  Print #RptHandle, Tab(49); Using("##,###,##0", TaxBill.ExptValue);
  Print #RptHandle, Tab(82); Using("###0.00", TaxBill.LateTaxDue);
  Print #RptHandle, Tab(89); Using("####0.00", TaxBill.TotalBillDue - TaxBill.OverPayAmt) 'added OverPayAmt 8/15/06
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(8); '"IF TAXES ARE ESCROWED SEND BILL TO"
  Print #RptHandle, Tab(8); '"MORTGAGE COMPANY."
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, 'added 6.23.06
  Print #RptHandle, Tab(8); Left$(QPTrim$(CustName$), 25)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(8); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(8); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"

End Sub

Private Sub PrintBSC32TF(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim NetTaxVal#
  Dim TAXRATE#
  
  TAXRATE# = TaxBill.RealTaxRate
  If TAXRATE# = 0 Then
    TAXRATE# = TaxBill.PersTaxRate
  End If

  Print #RptHandle, Chr$(27); Chr$(48); "~"

  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Using$("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(10); Using("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(19); Using("####0", TaxBill.CustRec);
  Print #RptHandle, Tab(34); Using("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(43); Using("#0.00", TAXRATE#);
  Print #RptHandle, Tab(48); Using("#####0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(2); Left$(TaxBill.RDesc1, 21)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, QPTrim$(PINTemp); Tab(20); Using("#######0", TaxBill.ExptValue)
  Print #RptHandle,
  Print #RptHandle, Tab(48); Using("###,##0.00", OldRound(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
  Print #RptHandle, Using("#######0", TaxBill.RealValue); Tab(11); Using("#######0", TaxBill.PersValue); Tab(22); Using("########0", NetTaxVal#)
  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  
  Print #RptHandle, Tab(24); Using$("###0", TaxBill.TaxYear);
  Print #RptHandle, Tab(30); Using$("#####0", TaxBill.BillNumber);
  Print #RptHandle, Tab(39); Using$("####0", TaxBill.CustRec);
  Print #RptHandle, Tab(48); Using$("###,##0.00", OldRound#(TaxBill.TotalBillDue + TaxBill.PriorYrBalance))
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(30); Left$(CustName$, 25)
  Print #RptHandle, Tab(30); Left$(QPTrim$(TaxBill.CustAdd1), 25)
  Print #RptHandle, Tab(30); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptHandle, Tab(30); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, "~"

End Sub
Private Sub PrintLaserLegal()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim dlm$
  Dim LA$, PrintCnt As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrnCnt As Integer
  Dim ThisFile$, FBill As Long, LBill As Long
  Dim CustCSZ$
  Dim PinNum$
  Dim Desc As String * 29
  Dim Name As String * 28
  
  dlm$ = "~"
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  RptFile$ = "TAXRPTS\TXLSRLEGAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
  NumOfTBRecs = BillCnt '8/6/08
  If OptMulti.Value = True Then
    ReDim PrintThis(1 To 1) As Long
    For x = 0 To fpList1.ListCount - 1
      fpList1.Row = x
      If fpList1.Selected = True Then
        fpList1.ListIndex = x
        PrintCnt = PrintCnt + 1
        ReDim Preserve PrintThis(1 To PrintCnt) As Long
        fpList1.Col = 3
        PrintThis(PrintCnt) = CLng(fpList1.ColText)
      End If
    Next x
    NumOfTBRecs = PrintCnt
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill '8/6/08
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoTo PrintIt '8/21/07
    End If
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
PrintIt: '8/21/07
        If TaxBill.LotOrAcre = "A" Then
          LA = "Acre"
        ElseIf TaxBill.LotOrAcre = "L" Then
          LA = "Lot"
        Else
          LA = "NA"
        End If
        If QPTrim$(TaxBill.LASize) <> "" Then
          LA = LA + "  " + "Parcel Size: " + CStr(TaxBill.LASize)
        End If
        TaxBill.CustZip = InsertZipDash(QPTrim$(TaxBill.CustZip))
        CustCSZ = QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
        If QPTrim$(TaxBill.RealPin) <> "" Then
          PinNum = QPTrim$(TaxBill.RealPin)
        ElseIf QPTrim$(TaxBill.PersPin) <> "" Then
          PinNum = QPTrim$(TaxBill.PersPin)
        Else
          PinNum = ""
        End If
        Desc = QPTrim$(QPTrim$(TaxBill.RDesc1))
        Name = QPTrim$(TaxBill.CustName)
        '                              0                      1                        2               3
        Print #RptHandle, CStr(TaxBill.TaxYear); dlm; TaxBill.BillNumber; dlm; TaxBill.CustRec; dlm; PinNum; dlm;
        '                   4         5         6               7                         8
        Print #RptHandle, Name; dlm; Desc; dlm; LA; dlm; TaxBill.RealValue; dlm; TaxBill.PersValue; dlm;
        '                        9                                       10                                     11
        Print #RptHandle, TaxBill.ExptValue; dlm; OldRound(TaxBill.PersValue + TaxBill.RealValue); dlm; TaxBill.RealTaxRate; dlm;
        '                                     12                                         13                                   14
        Print #RptHandle, OldRound(TaxBill.RealTaxDue + TaxBill.PersTaxDue); dlm; TaxBill.LateTaxDue; dlm; TaxBill.TotalBillDue - TaxBill.OverPayAmt; dlm; 'added OverPayAmt 8/15/06
        '                         15                           16                      17
'        Print #RptHandle, QPTrim$(TaxMasterRec.Name); dlm; QPTrim$(TaxMasterRec.Add1); dlm; QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip); dlm;
        Print #RptHandle, "                       "; dlm; "                  "; dlm; "                     "; dlm;
        '                   18                    19                           20                     21
        Print #RptHandle, Name; dlm; QPTrim$(TaxBill.CustAdd1); dlm; QPTrim$(TaxBill.CustAdd2); dlm; CustCSZ
        PrnCnt = PrnCnt + 1
      End If
    Next x
  
  Close
  
  arTaxLsrLegal.Show
  
End Sub

Private Sub PrintPostCard(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As TaxBillType, CustName$, PINTemp$, PrnCnt As Long)
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim ThisDesc As String * 20
  Dim LotsAcres As String * 20
  Dim ThisCName As String * 22
  Dim ThisName As String * 28
  Dim ThisAdd1 As String * 28
  Dim ThisAdd2 As String * 28
  Dim ThisAdd3 As String * 28
  Dim FF$
  Dim PastDue#
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  ThisName = QPTrim$(TaxBill.CustName)
  ThisCName = QPTrim$(TaxBill.CustName)
  ThisAdd1 = QPTrim$(TaxBill.CustAdd1)
  ThisAdd2 = QPTrim$(TaxBill.CustAdd2)
  ThisAdd3 = QPTrim$(TaxBill.CustAdd3) & " " & QPTrim$(TaxBill.CustZip)
  OpenPersPropFile PHandle, NumOfPRecs
  OpenRealPropFile RHandle, NumOfRRecs
  FF$ = Chr(12)
  
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  '-----------------------------------------
  Print #RptHandle, Tab(1); Using$("###0", TaxBill.TaxYear); Tab(7); Using$("####0", TaxBill.BillNumber); Tab(14); Using$("######0", TaxBill.CustPin);
'  Print #RptHandle, Tab(22); ""; Tab(33); Using$("###0", TaxBill.TaxYear); Tab(40); Using$("####0", TaxBill.BillNumber); Tab(50); Using$("######0", TaxBill.CustPin)
  If InStr(TaxMasterRec.Name, "SEVEN DEVILS") Then
    Print #RptHandle, Tab(22); ""; Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); ""; Tab(60); Using$("######0", TaxBill.CustPin)
  ElseIf InStr(TaxMasterRec.Name, "ANDREWS") Then
    Print #RptHandle, Tab(22); ""; Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); Using$("######0", TaxBill.CustPin); Tab(60); Using$("$##,##0.00", TaxBill.TotalBillDue)
  Else
    Print #RptHandle, Tab(26); Right(QPTrim$(TaxBill.RealPin), 11); Tab(40); Using$("###0", TaxBill.TaxYear); Tab(50); Using$("####0", TaxBill.BillNumber); Tab(60); Using$("######0", TaxBill.CustPin)
  End If
  '---end of line 1-------------------------
  Print #RptHandle, Tab(10); ThisCName 'end of line 2
  If TaxBill.RealPropRecord > 0 And TaxBill.PersPropRecord > 0 Then
    Get RHandle, TaxBill.RealPropRecord, RealRec
    Get PHandle, TaxBill.PersPropRecord, PersRec
    ThisDesc = "Real And Personal"
    LotsAcres = QPTrim$(RealRec.LOTACRE) + "/" + QPTrim$(RealRec.LOTNUMB)
  ElseIf TaxBill.RealPropRecord > 0 Then
    Get RHandle, TaxBill.RealPropRecord, RealRec
    ThisDesc = QPTrim$(RealRec.PROPNOT1)
    LotsAcres = QPTrim$(RealRec.LOTACRE) + "/" + QPTrim$(RealRec.LOTNUMB)
  ElseIf TaxBill.PersPropRecord > 0 Then
    Get PHandle, TaxBill.PersPropRecord, PersRec
    ThisDesc = QPTrim$(PersRec.DESC1)
    LotsAcres = ""
  Else
    ThisDesc = ""
    LotsAcres = ""
  End If
  If TaxBill.OverPayAmt > 0 Then 'late tax should not come into play because of the overpay amt
    TaxBill.TotalBillDue = OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt)
  End If
  Print #RptHandle, Tab(10); ThisDesc
  Print #RptHandle, Tab(10); LotsAcres
  Print #RptHandle,
  Print #RptHandle, Tab(2); Using$("$###,###.00", TaxBill.RealValue); Tab(15); Using$("$###,##0.00", TaxBill.PersValue); Tab(26); Using$("$###,##0.00", TaxBill.ExptValue);
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  If TaxBill.RealTaxRate > 0 Then
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", TaxBill.RealTaxRate);
  ElseIf TaxBill.PersTaxRate > 0 Then
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", TaxBill.PersTaxRate);
  Else
    Print #RptHandle, Tab(2); Using$("$###,##0.00", OldRound(TaxBill.RealValue + TaxBill.PersValue)); Tab(16); Using$("##0.00", 0);
  End If
  Print #RptHandle, Tab(26); Using$("$###,##0.00", OldRound(TaxBill.OverPayAmt + TaxBill.TotalBillDue - TaxBill.LateTaxDue)); Tab(42); ThisName '8/16/06 added overpayment
  Print #RptHandle, Tab(42); ThisAdd1
  Print #RptHandle, Tab(42); ThisAdd2
  Print #RptHandle, Tab(28); Using$("$#,##0.00", TaxBill.LateTaxDue); Tab(42); ThisAdd3
  Print #RptHandle,
  Print #RptHandle, Using$("$###,##0.00", PastDue); Tab(26); Using$("$###,##0.00", TaxBill.TotalBillDue)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  
  Close RHandle
  Close PHandle
End Sub

Private Sub PrintLaserLegalHP()
  Dim RptHandle As Integer
  Dim RptFile$
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim dlm$
  Dim LA$, PrintCnt As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PrnCnt As Integer
  Dim ThisFile$, FBill As Long, LBill As Long
  Dim Desc As String * 29
  Dim Name As String * 28
  Dim CustCSZ$
  Dim PinNum$
  
  dlm$ = "~"
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  RptFile$ = "TAXRPTS\TXLSRLEGALHP.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
  NumOfTBRecs = BillCnt '8/6/08

  If OptMulti.Value = True Then
    ReDim PrintThis(1 To 1) As Long
    For x = 0 To fpList1.ListCount - 1
      fpList1.Row = x
      If fpList1.Selected = True Then
        fpList1.ListIndex = x
        PrintCnt = PrintCnt + 1
        ReDim Preserve PrintThis(1 To PrintCnt) As Long
        fpList1.Col = 3
        PrintThis(PrintCnt) = CLng(fpList1.ColText)
      End If
    Next x
    NumOfTBRecs = PrintCnt
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill '8/6/08
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoTo PrintIt '8/21/07
    End If
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
PrintIt: '8/21/07
        If TaxBill.LotOrAcre = "A" Then
          LA = "Acre"
        ElseIf TaxBill.LotOrAcre = "L" Then
          LA = "Lot"
        Else
          LA = "NA"
        End If
        If QPTrim$(TaxBill.LASize) <> "" Then
          LA = LA + "  " + "Parcel Size: " + CStr(TaxBill.LASize)
        End If
        Desc = QPTrim$(TaxBill.RDesc1)
        Name = QPTrim$(TaxBill.CustName)
        TaxBill.CustZip = InsertZipDash(QPTrim$(TaxBill.CustZip))
        CustCSZ = QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
        If QPTrim$(TaxBill.RealPin) <> "" Then
          PinNum = QPTrim$(TaxBill.RealPin)
        ElseIf QPTrim$(TaxBill.PersPin) <> "" Then
          PinNum = QPTrim$(TaxBill.PersPin)
        Else
          PinNum = ""
        End If
        '                              0                      1                        2                3
        Print #RptHandle, CStr(TaxBill.TaxYear); dlm; TaxBill.BillNumber; dlm; TaxBill.CustRec; dlm; PinNum; dlm;
        '                   4         5          6               7                         8
        Print #RptHandle, Name; dlm; Desc; dlm; LA; dlm; TaxBill.RealValue; dlm; TaxBill.PersValue; dlm;
        '                        9                                       10                                     11
        Print #RptHandle, TaxBill.ExptValue; dlm; OldRound(TaxBill.PersValue + TaxBill.RealValue); dlm; TaxBill.RealTaxRate; dlm;
        '                                     12                                         13                                    14
        Print #RptHandle, OldRound(TaxBill.RealTaxDue + TaxBill.PersTaxDue); dlm; TaxBill.LateTaxDue; dlm; TaxBill.TotalBillDue - TaxBill.OverPayAmt; dlm; 'added OverPayAmt 8/15/06
        '                         15                           16                      17
'        Print #RptHandle, QPTrim$(TaxMasterRec.Name); dlm; QPTrim$(TaxMasterRec.Add1); dlm; QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip); dlm;
        Print #RptHandle, "                       "; dlm; "                  "; dlm; "                     "; dlm;
        '                  18                     19                             20                    21
        Print #RptHandle, Name; dlm; QPTrim$(TaxBill.CustAdd1); dlm; QPTrim$(TaxBill.CustAdd2); dlm; CustCSZ
        PrnCnt = PrnCnt + 1
      End If
    Next x
  
  Close
  
  arTaxLsrLegalHP.Show
  
End Sub

