VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxReprintPosted 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Reprints of Posted Bills"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxReprintPosted.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   3120
      Left            =   1560
      TabIndex        =   17
      Top             =   4080
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
      ColDesigner     =   "frmVATaxReprintPosted.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbFile 
      Height          =   405
      Left            =   4755
      TabIndex        =   1
      Top             =   1680
      Width           =   3945
      _Version        =   196608
      _ExtentX        =   6959
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
      ColDesigner     =   "frmVATaxReprintPosted.frx":0CA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   495
      Left            =   8400
      TabIndex        =   21
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
      ButtonDesigner  =   "frmVATaxReprintPosted.frx":1009
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
      Left            =   3240
      TabIndex        =   11
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
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "The program will print all bills from the 'First Bill' selection to the 'Last Bill' selection."
      Top             =   2880
      Width           =   1932
   End
   Begin EditLib.fpDoubleSingle fpDblSnglLastBill 
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   2850
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
      Left            =   6360
      TabIndex        =   3
      Top             =   2850
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
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
   Begin EditLib.fpText fptxtCurrForm 
      Height          =   396
      Left            =   5520
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Late notices are selected on the System Setup screen."
      Top             =   1200
      Width           =   2856
      _Version        =   196608
      _ExtentX        =   5027
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   6000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmVATaxReprintPosted.frx":11EA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   1200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmVATaxReprintPosted.frx":13C9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   492
      Left            =   3600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7800
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
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
      ButtonDesigner  =   "frmVATaxReprintPosted.frx":15A5
   End
   Begin EditLib.fpDateTime fptxtPostDate 
      Height          =   375
      Left            =   5790
      TabIndex        =   19
      Top             =   2160
      Width           =   1755
      _Version        =   196608
      _ExtentX        =   3096
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
      Text            =   "02/24/2005"
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
      Appearance      =   0
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
      Caption         =   "Post Date:"
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
      Left            =   4110
      TabIndex        =   20
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "label"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4800
      TabIndex        =   18
      Top             =   650
      Width           =   2028
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
      Left            =   3240
      TabIndex        =   13
      Top             =   1260
      Width           =   2028
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1095
      Left            =   240
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
      Left            =   2760
      TabIndex        =   10
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
      Left            =   2760
      TabIndex        =   9
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
      Left            =   240
      TabIndex        =   7
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
      Left            =   8160
      TabIndex        =   6
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
      Left            =   5280
      TabIndex        =   4
      Top             =   2910
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3615
      Left            =   1440
      Top             =   3960
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
      Left            =   2955
      TabIndex        =   2
      Top             =   1770
      Width           =   1620
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
      Top             =   276
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   768
      Index           =   1
      Left            =   1500
      Top             =   240
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   1500
      Top             =   120
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxReprintPosted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim FirstLoad As Boolean
  Dim FirstNum As Long
  Dim LastNum As Long
  Dim Real As Integer
  Dim TownName$
  Dim Add1$
  Dim Add2$
  Dim Add3$
  Dim DirRContents() As String
  Dim DirRCnt As Integer
  Dim DirPContents() As String
  Dim DirPCnt As Integer
  Dim MyPath$
  Dim PPTRARmvl As Boolean
  Dim BillIdx() As Long
  Dim BillCnt As Long
  
Private Sub cmdClear_Click()
  fpList1.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmVATaxBillingMenu.Show
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FirstLoad = True
  Me.HelpContextID = hlpReprintPostedTax
  Real = -1
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxReprintPosted.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub LoadMe()
  Dim x As Integer
  Dim ThisYear As Integer
  Dim ThisFile$
  Dim GotIt As Boolean
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MyName$
  
'  On Error GoTo ERRORSTUFF
  If frmVATaxBillingMenu.Real = True Then
    Real = 1
    lblType.Caption = "Real Property"
  Else
    Real = 2
    lblType.Caption = "Personal Property"
  End If
  DirRCnt = 0
  DirPCnt = 0
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    MyName = Dir
    If Len(MyName) > 4 Then
      If Real = 1 Then
        If Mid(MyName, 5, 1) = "R" Then
          DirRCnt = DirRCnt + 1
          ReDim Preserve DirRContents(DirRCnt) As String
          DirRContents(DirRCnt) = MyName
          If DirRCnt = 1 Then
            fpcmbFile.Text = MyName
          End If
          fpcmbFile.AddItem MyName
        End If
      ElseIf Real = 2 Then
        If Mid(MyName, 5, 1) = "P" Then
          DirPCnt = DirPCnt + 1
          ReDim Preserve DirPContents(DirPCnt) As String
          DirPContents(DirPCnt) = MyName
          If DirPCnt = 1 Then
            fpcmbFile.Text = MyName
          End If
          fpcmbFile.AddItem MyName
        End If
      End If
    End If
  Loop
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName$ = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20002 Then
    cmdAlign.Enabled = False
  End If
  Select Case TaxMasterRec.TaxForm
    Case 30000
      fptxtCurrForm.Text = "STANDARD"
'    Case 21837
'      fptxtCurrForm.Text = "MULTI-PART"
'    Case 20304
'      fptxtCurrForm.Text = "POSTCARD"
    Case 16716
      fptxtCurrForm.Text = "LASER"
    Case 29999
      fptxtCurrForm.Text = "EXPORT COMBINED"
    Case 20000
      fptxtCurrForm.Text = "EXPORT REAL"
    Case 20001
      fptxtCurrForm.Text = "EXPORT PERSONAL"
    Case 20002
      fptxtCurrForm.Text = "LASER ITEMIZED"
    Case 20003
      fptxtCurrForm.Text = "MDLTWN"
    Case 20004
      fptxtCurrForm.Text = "CDRBLUFF"
    Case Else
      fptxtCurrForm.Text = "UNKNOWN"
  End Select
  
'  If frmVATaxBillingMenu.Real = True Then
'    Real = 1
'  Else
'    Real = 2
'  End If
  
  Call LoadList
  OptRange.Value = True
  Close
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPOsted", "LoadMe", Erl)
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
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim ThisFile$
  Dim THandle As Integer
  Dim RTaxBill As VARETaxBillType
  Dim PTaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, y As Integer
  Dim FirstBill$
  Dim LastBill$
  Dim ThisLen As Integer
  Dim BillArr() As Long
  Dim BillArrCnt As Long
  Dim BillNum() As Long
  Dim BigNum As Long
  Dim BigNumStatic As Long
  Dim Nextx As Long
  Dim HoldThis As Long
  Dim Thisx As Long
  
'  On Error GoTo ERRORSTUFF
  PPTRARmvl = False
 
  If Real < 0 Then Exit Sub
  
  FirstBill$ = ""
  LastBill$ = ""
  fpList1.Clear
  ThisFile = QPTrim$(fpcmbFile.Text)
  If ThisFile = "" Then
    Call TaxMsg(900, "Please make a selection from the list of files.")
    fpcmbFile.SetFocus
    Exit Sub
  End If
  
  If Real = 1 Then
    If DirRCnt = 0 Then
      Call TaxMsg(900, "No real tax billing records available.")
      Close
      Exit Sub
    End If
    
    For y = 1 To DirRCnt
      ThisLen = Len(DirRContents(y))
      If DirRContents(y) = fpcmbFile.Text Then
        OpenRealPostedReprintFile THandle, NumOfTBRecs, MyPath + DirRContents(y)
        Exit For
      End If
    Next y

    If y > DirRCnt Then
      Call TaxMsg(900, "ERROR: There is a problem loading the selected file. Please try again.")
      Close
      Exit Sub
    End If
  
    GoSub FillIdxR '8/6/08
    
    For x = 1 To BillArrCnt '8/6/08
      Get THandle, BillIdx(x), RTaxBill '8/6/08
'    For x = 1 To NumOfTBRecs
'      Get THandle, x, RTaxBill
      If RTaxBill.BillNumber >= 0 Then
        fptxtPostDate.Text = MakeRegDate(RTaxBill.PostDate)
        If FirstBill = "" Then FirstBill = CStr(RTaxBill.BillNumber)
        FirstNum = FirstBill
'        fpList1.InsertRow = CStr(RTaxBill.BillNumber) + Chr(9) + QPTrim$(RTaxBill.CustName) + Chr(9) + Using$("$###,###,##0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt)) + Chr(9) + CStr(x)
        fpList1.InsertRow = CStr(RTaxBill.BillNumber) + Chr(9) + QPTrim$(RTaxBill.CustName) + Chr(9) + Using$("$###,###,##0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt)) + Chr(9) + CStr(BillIdx(x)) '8/6/08
      End If
    Next x
    
    For x = BillArrCnt To 1 Step -1 '8/6/08
      Get THandle, BillIdx(x), RTaxBill '8/6/08
'    For x = NumOfTBRecs To 1 Step -1
'      Get THandle, x, RTaxBill
      If RTaxBill.BillNumber >= 0 Then
        LastBill = CStr(RTaxBill.BillNumber)
        LastNum = LastBill
        Exit For
      End If
    Next x
  
    Close
    fpDblSnglFirstBill = CLng(FirstBill)
    fpDblSnglLastBill = CLng(LastBill)
  ElseIf Real = 2 Then
    If DirPCnt = 0 Then
      Call TaxMsg(900, "No personal tax billing records available.")
      Close
      Exit Sub
    End If
    
    OpenBillPostDateFile PostHandle, NumOfPostRecs
    For x = 1 To NumOfPostRecs
      Get PostHandle, x, PostRec
        If QPTrim$(PostRec.BackUpName) = "TAXBILLBU\" + fpcmbFile.Text Then
          If PostRec.PPTRAPosted = "Y" Then
            PPTRARmvl = True
          End If
          Exit For
        End If
    Next x
    Close PostHandle
    
    For y = 1 To DirPCnt
      ThisLen = Len(DirPContents(y))
      If DirPContents(y) = fpcmbFile.Text Then
        OpenPersPostedReprintFile THandle, NumOfTBRecs, MyPath + DirPContents(y)
        Exit For
      End If
    Next y
    
    If y > DirPCnt Then
      Call TaxMsg(900, "ERROR: There is a problem loading the selected file. Please try again.")
      Close
      Exit Sub
    End If
    
    GoSub FillIdxP '8/6/08
    
    For x = 1 To BillArrCnt '8/6/08
      Get THandle, BillIdx(x), PTaxBill '8/6/08
'    For x = 1 To NumOfTBRecs
'      Get THandle, x, PTaxBill
      If PTaxBill.BillNumber >= 0 Then
        fptxtPostDate.Text = MakeRegDate(PTaxBill.PostDate)
        If FirstBill = "" Then FirstBill = CStr(PTaxBill.BillNumber)
        FirstNum = FirstBill
'        fpList1.InsertRow = CStr(PTaxBill.BillNumber) + Chr(9) + QPTrim$(PTaxBill.CustName) + Chr(9) + Using$("$###,###,##0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt)) + Chr(9) + CStr(x)
        fpList1.InsertRow = CStr(PTaxBill.BillNumber) + Chr(9) + QPTrim$(PTaxBill.CustName) + Chr(9) + Using$("$###,###,##0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt)) + Chr(9) + CStr(BillIdx(x)) '8/6/08
      End If
    Next x
    
    For x = BillArrCnt To 1 Step -1 '8/6/08
      Get THandle, BillIdx(x), PTaxBill '8/6/08
'    For x = NumOfTBRecs To 1 Step -1
'      Get THandle, x, PTaxBill
      If PTaxBill.BillNumber >= 0 Then
        LastBill = CStr(PTaxBill.BillNumber)
        LastNum = LastBill
        Exit For
      End If
    Next x
  
    Close
    
    fpDblSnglFirstBill = CLng(FirstBill)
    fpDblSnglLastBill = CLng(LastBill)
  
  End If
  
  Exit Sub
  
FillIdxR: '8/6/08
  BigNum = 0
  For x = 1 To NumOfTBRecs
    Get THandle, x, RTaxBill
    If RTaxBill.BillNumber > 0 Then
      BillArrCnt = BillArrCnt + 1
      ReDim Preserve BillArr(1 To BillArrCnt) As Long
      ReDim Preserve BillNum(1 To BillArrCnt) As Long
      BillArr(BillArrCnt) = x
      BillNum(BillArrCnt) = RTaxBill.BillNumber
      If RTaxBill.BillNumber > BigNum Then
        BigNum = RTaxBill.BillNumber
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
  
FillIdxP: '8/6/08
  BigNum = 0
  For x = 1 To NumOfTBRecs
    Get THandle, x, PTaxBill
    If PTaxBill.BillNumber > 0 Then
      BillArrCnt = BillArrCnt + 1
      ReDim Preserve BillArr(1 To BillArrCnt) As Long
      ReDim Preserve BillNum(1 To BillArrCnt) As Long
      BillArr(BillArrCnt) = x
      BillNum(BillArrCnt) = PTaxBill.BillNumber
      If PTaxBill.BillNumber > BigNum Then
        BigNum = PTaxBill.BillNumber
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "LoadList", Erl)
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

Private Sub PrintLaser1()
'  Dim ToPrint As String
'  Dim TaxRptT As Integer
'  Dim ReportFile As String
'  Dim RptHandle As Integer
'  Dim TBRec As VARETaxBillType
'  Dim TBHandle As Integer
'  Dim NumOfTBRecs As Long
'  Dim x As Long
'  Dim dlm$
'  Dim TBDRec As TxBillLaser1DefaultsType
'  Dim TBDHandle As Integer
'  Dim FBill&, PrnCnt&
'  Dim LBill&
'  Dim PCnt As Integer
'  Dim NCnt As Integer
'  Dim ThisRate As Double
'  Dim ThisYear$
'  Dim PrintCnt As Long
'
'  On Error GoTo ERRORSTUFF
'
'  ThisYear$ = QPTrim$(fpcmbFile.Text)
'  dlm$ = "~"
'  FBill = fpDblSnglFirstBill.Value
'  LBill = fpDblSnglLastBill.Value
'  ReportFile$ = StartPath$ + "/TaxBil1.PRN"
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'  OpenTxBillRealFile TBDHandle
'  Get #TBDHandle, 1, TBDRec
'  Close TBDHandle
'  ARptVATempTaxBill.Head1 = QPTrim(TBDRec.TxtHead1)
'  ARptVATempTaxBill.Head2 = QPTrim(TBDRec.TxtHead2)
'  ARptVATempTaxBill.LblOpt1 = QPTrim(TBDRec.txtOpt1)
'  ARptVATempTaxBill.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
'  ARptVATempTaxBill.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
'  ARptVATempTaxBill.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
'  ARptVATempTaxBill.LblPgph1 = QPTrim(TBDRec.txtPgph0)
'  ARptVATempTaxBill.LblPgph2 = QPTrim(TBDRec.txtPgph1)
'  ARptVATempTaxBill.LblPgph3 = QPTrim(TBDRec.txtPgph2)
'  ARptVATempTaxBill.LblPgph4 = QPTrim(TBDRec.txtPgph3)
'  ARptVATempTaxBill.LblPgph5 = QPTrim(TBDRec.txtPgph4)
'  ARptVATempTaxBill.LblPgph6 = QPTrim(TBDRec.txtPgph5)
'  ARptVATempTaxBill.LblPgph7 = QPTrim(TBDRec.txtPgph6)
'  ARptVATempTaxBill.LblPgph8 = QPTrim(TBDRec.txtPgph7)
'  ARptVATempTaxBill.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
'  ARptVATempTaxBill.LblHead4 = QPTrim(TBDRec.txtHead4)
'  ARptVATempTaxBill.LblHead5 = QPTrim(TBDRec.txtHead5)
'  ARptVATempTaxBill.LblHead6 = QPTrim(TBDRec.txtHead6)
'  ARptVATempTaxBill.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
'  ARptVATempTaxBill.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
'  If TBDRec.dologo = 1 Then
'    If Exist("towntaxlogo.bmp") Then
'      ARptVATempTaxBill.Image1.Picture = LoadPicture("towntaxlogo.bmp")
'      ARptVATempTaxBill.Image1.Visible = True
'    End If
'  End If
'
'  frmVATaxShowPctComp.Label1 = "Printing Tax Bills"
'  frmVATaxShowPctComp.cmdCancel.Visible = False
'  frmVATaxShowPctComp.Show, Me
'  cmdProcess.Enabled = False
'  cmdExit.Enabled = False
'  EnableCloseButton Me.hwnd, False
'
'  OpenRealPostedReprintFile TBHandle, NumOfTBRecs, ThisYear
'  If OptMulti.Value = True Then
'    ReDim PrintThis(1 To 1) As Long
'    For x = 0 To fpList1.ListCount - 1
'      fpList1.Row = x
'      If fpList1.Selected = True Then
'        fpList1.ListIndex = x
'        PrintCnt = PrintCnt + 1
'        ReDim Preserve PrintThis(1 To PrintCnt) As Long
'        fpList1.Col = 3
'        PrintThis(PrintCnt) = CLng(fpList1.ColText)
'      End If
'    Next x
'    NumOfTBRecs = PrintCnt
'  End If
'
'  For x = 1 To NumOfTBRecs
'    If OptMulti.Value = False Then
'      Get TBHandle, x, TBRec
'    Else
'      Get TBHandle, PrintThis(x), TBRec
'    End If
'      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
'        If TBRec.TotalBillDue > 0 Then
'          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
'          Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
'          Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
'          Print #RptHandle, QPTrim$(TBRec.RDesc2); dlm; QPTrim$(TBRec.RDesc1); dlm;
'          Print #RptHandle, TBRec.RealValue; dlm; TBRec.PersValue; dlm; TBRec.ExptValue; dlm;
'          If TBRec.RealTaxDue > 0 And TBRec.PersTaxDue > 0 Then
'            ThisRate = TBRec.RealTaxRate
'          ElseIf TBRec.RealTaxDue <= 0 And TBRec.PersTaxDue > 0 Then
'            ThisRate = TBRec.PersTaxRate
'          ElseIf TBRec.RealTaxDue > 0 And TBRec.PersTaxDue <= 0 Then
'            ThisRate = TBRec.RealTaxRate
'          Else
'            ThisRate = 0
'          End If
'          Print #RptHandle, OldRound(TBRec.RealValue + TBRec.PersValue); dlm; ThisRate; dlm;
'          Print #RptHandle, TBRec.TotalBillDue; dlm; TBDRec.dologo
'          PCnt = PCnt + 1
'        End If
'      End If
'    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      cmdProcess.Enabled = True
'      cmdExit.Enabled = True
'      EnableCloseButton Me.hwnd, True
'      Exit Sub
'    End If
'  Next x
'  Close TBHandle
'
'  Unload frmVATaxShowPctComp
'  cmdProcess.Enabled = True
'  cmdExit.Enabled = True
'  EnableCloseButton Me.hwnd, True
'
'  Close
'  ARptVATempTaxBill.GetName ReportFile$
'  ARptVATempTaxBill.Show
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "PrintLaser1", Erl)
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
 
End Sub

Private Sub PrintStandard(ByVal RptHandle As Integer, TBHandle As Integer, TaxBill As VARETaxBillType, CustName$, PINTemp$, PrnCnt As Long)
'  Dim TAXRATE#
'  Dim NetTaxVal#
'
'  On Error GoTo ERRORSTUFF
'
'  TAXRATE# = TaxBill.RealTaxRate
'  If TAXRATE# = 0 Then
'    TAXRATE# = TaxBill.PersTaxRate
'  End If
'
'  Print #RptHandle, "~"; Tab(50); Using("###0", PrnCnt)
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, TaxBill.TaxYear;
'  Print #RptHandle, Tab(7); Using("######", TaxBill.BillNumber);
'  Print #RptHandle, Tab(16); Using("#####", TaxBill.CustRec);
'  Print #RptHandle, Tab(23); QPTrim$(PINTemp); Tab(37); Using("####", TaxBill.TaxYear);
'
'  Print #RptHandle, Tab(42); Using("########", TaxBill.CustRec);
'  'PRINT #RptHandle, TAB(49); USING "######,#.##"; TaxBill.TotalBillDue + TaxBill.PriorYrBalance
'  Print #RptHandle, Tab(51); Using("######", TaxBill.CustRec)
'  Print #RptHandle,
'  Print #RptHandle, Tab(11); Left$(QPTrim$(CustName$), 21)
'  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc1), 21)
'  Print #RptHandle, Tab(11); Left$(QPTrim$(TaxBill.RDesc2), 21)
'  'v line 12
'  Print #RptHandle,
'
'  Print #RptHandle, Using("###,###,##0", TaxBill.RealValue); Tab(15); TaxBill.PersValue;
'  NetTaxVal# = OldRound#(TaxBill.RealValue + TaxBill.PersValue)
'  NetTaxVal# = OldRound#(NetTaxVal# - TaxBill.ExptValue)
'  Print #RptHandle, Tab(25); Using("#,###,##0", TaxBill.ExptValue);
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(36); Left$(QPTrim$(CustName$), 25)
'  Print #RptHandle, Using("###,###,##0", NetTaxVal#);
'  Print #RptHandle, Tab(16); Using("#.##", TAXRATE#);
'  Print #RptHandle, Tab(21); Using("##,###,##0.00", OldRound#(TaxBill.TotalBillDue - TaxBill.LateTaxDue));
'  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd1), 25)
'  Print #RptHandle, Tab(36); Left$(QPTrim$(TaxBill.CustAdd2), 25)
'  Print #RptHandle, Tab(36); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
'  Print #RptHandle, Tab(21); Using("##,###,##0.00", TaxBill.LateTaxDue)
'  Print #RptHandle,
'  Print #RptHandle, Tab(2); Using("###,##0.0", TaxBill.PriorYrBalance);
'  If TaxBill.OverPayAmt > 0 Then
'    Print #RptHandle, Tab(21); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance); Tab(47); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance - TaxBill.OverPayAmt)
'  Else
'    Print #RptHandle, Tab(21); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance); Tab(47); Using("#,###,##0.00", TaxBill.TotalBillDue + TaxBill.PriorYrBalance)
'  End If
'  If TaxBill.OverPayAmt > 0 Then
'    Print #RptHandle, "Credit of " + QPTrim$(Using$("$###,##0.00", TaxBill.OverPayAmt)) + " has been applied."
'  Else
'    Print #RptHandle,
'  End If
'  Print #RptHandle, "~"
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "PrintStandard", Erl)
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
'
End Sub
Private Sub cmdProcess_Click()
  Dim RptHandle As Integer
  Dim RTaxBill As VARETaxBillType
  Dim PTaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long, BillNo&
  Dim WhatRec&, PrnCnt&
  Dim PINTemp As String * 18
  Dim CustName$
  Dim RptFile$, FBill&, LBill&
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisFile$
  Dim PrintCnt As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
'  On Error GoTo ERRORSTUFF
  
  PrnCnt = 0
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
  
  If TaxMasterRec.TaxForm = 20003 Then
    Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
    If Real = 1 Then
      Call PrintMdltwnReal
      Exit Sub
    ElseIf Real = 2 Then
      Call TaxMsg(800, "Itemization is not possible with these posted reprints because property owned when these bills were posted may have changed.")
      Call PrintMdltwnPers
      Exit Sub
    End If
  End If
  
  If TaxMasterRec.TaxForm = 20004 Then
    Call TaxMsg(900, "Pitch 12 is recommended for this bill.")
    If Real = 1 Then
      Call PrintCdrBluffReal
      Exit Sub
    ElseIf Real = 2 Then
      Call TaxMsg(800, "Itemization is not possible with these posted reprints because property owned when these bills were posted may have changed.")
      Call PrintCdrBluffPers
      Exit Sub
    End If
  End If
  
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20002 Then
    If Real = 1 Then
      Call PrintLaserReal1
      Exit Sub
    ElseIf Real = 2 Then
      If TaxMasterRec.TaxForm = 16716 Then
        Call PrintLaserPers1
      Else
        Call PrintPersLaserItemized
      End If
      Exit Sub
    End If
  End If
  
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
  
  If Real = 1 Then
    RptHandle = FreeFile
    RptFile$ = "R" + Mid(fpcmbFile.Text, 2, 2) + "REPRN.PRN"
  
    Open RptFile For Output As RptHandle
  
    OpenRealPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
    NumOfTBRecs = BillCnt '8/6/08

    OpenTaxCustFile TCHandle, NumOfTCRecs
    PrnCnt = fpDblSnglFirstBill
  
    Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
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
      If PrintCnt = 0 Then
        Call TaxMsg(900, "Please make a selection from the list.")
        Close
        Exit Sub
      End If
      NumOfTBRecs = PrintCnt
    End If
      
    For x = 1 To NumOfTBRecs
      If OptMulti.Value = False Then
'        Get TBHandle, x, RTaxBill
        Get TBHandle, BillIdx(x), RTaxBill '8/6/08
      Else
        Get TBHandle, PrintThis(x), RTaxBill
        GoTo PrintIt '8/21/07
      End If
      If RTaxBill.BillNumber >= FBill And RTaxBill.BillNumber <= LBill Then
PrintIt: '8/21/07
        If QPTrim$(RTaxBill.RealPin) <> "" Then
          RSet PINTemp = QPTrim$(RTaxBill.RealPin)
        Else
          RSet PINTemp = "0"
        End If
        CustName$ = QPTrim$(RTaxBill.CustName)
        If InStr(TaxMasterRec.Name, "HALIFAX") Then
          Call PrintHalifaxStandardReal(RptHandle, TBHandle, RTaxBill, PrnCnt) ', TCHandle, TaxCust, PrnCnt)
        Else
          Call PrintRealVAStandard(RptHandle, TBHandle, RTaxBill, TCHandle, TaxCust, PrnCnt)
        End If
        PrnCnt = PrnCnt + 1
      End If
    Next x
    Close TBHandle
    Close RptHandle
  
    ViewPrint RptFile$, "Real Tax Bill Reprinting", True
  ElseIf Real = 2 Then
    RptHandle = FreeFile
    RptFile$ = "P" + Mid(fpcmbFile.Text, 2, 2) + "REPRN.PRN"
    Open RptFile For Output As RptHandle
  
    OpenPersPropFile PHandle, NumOfPRecs
    OpenTaxCustFile TCHandle, NumOfTCRecs
    OpenPersPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
    NumOfTBRecs = BillCnt '8/6/08

    PrnCnt = fpDblSnglFirstBill
  
    Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
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
'        Get TBHandle, x, PTaxBill
        Get TBHandle, BillIdx(x), PTaxBill '8/6/08
      Else
        Get TBHandle, PrintThis(x), PTaxBill
      End If
      If PTaxBill.BillNumber >= FBill And PTaxBill.BillNumber <= LBill Then
        If QPTrim$(PTaxBill.PersPin) <> "" Then
          RSet PINTemp = QPTrim$(PTaxBill.PersPin)
        Else
          RSet PINTemp = "0"
        End If
        CustName$ = QPTrim$(PTaxBill.CustName)
        If InStr(TaxMasterRec.Name, "HALIFAX") Then
          Call PrintHalifaxStandardPersonal(RptHandle, TBHandle, PTaxBill, PrnCnt)
        Else
          Call PrintPersVAStandard(RptHandle, TBHandle, PTaxBill, PrnCnt)
        End If
      End If
      PrnCnt = PrnCnt + 1
    Next x
    Close TBHandle
    Close RptHandle
    If PrnCnt = 0 Then
      Call TaxMsg(900, "Please make a selection from the list.")
      Close
      Exit Sub
    End If
  
    ViewPrint RptFile$, "Personal Tax Bill Reprinting", True
  End If
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "cmdProcess_Click", Erl)
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
End Sub

Private Sub OptRange_Click()
  fpList1.MultiSelect = MultiSelectNone
End Sub
Private Sub cmdAlign_Click()
  Dim Handle As Integer
  Dim TempHandle As Integer
  Dim cnt As Integer
  Dim TextLine$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  If InStr(lblType.Caption, "Real") Then
    Select Case TaxMasterRec.TaxForm
      Case 30000
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("VASTANDRMSK.TXT") Then
          alnRpt = "VASTANDRMSK.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'VASTANDRMSK.TXT'.")
          Exit Sub
        End If
      Case 20003
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("MdltwnRMask.TXT") Then
          alnRpt = "MdltwnRMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'MdltwnRMask.TXT'.")
          Exit Sub
        End If
      Case 20004
        Call TaxMsg(900, "Pitch 12 is recommended for this bill.")
        If Exist("CdrBluffRMask.TXT") Then
          alnRpt = "CdrBluffRMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'CdrBluffRMask.TXT'.")
          Exit Sub
        End If
      Case Else
        Call TaxMsg(900, "No mask is available.")
        Exit Sub
    End Select
  End If
    
  If InStr(lblType.Caption, "Personal") Then
    Select Case TaxMasterRec.TaxForm
      Case 30000
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("VASTANDPMSK.TXT") Then
          alnRpt = "VASTANDPMSK.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'VASTANDPMSK.TXT'.")
          Exit Sub
        End If
      Case 20003
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("MdltwnPMask.TXT") Then
          alnRpt = "MdltwnPMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'MdltwnPMask.TXT'.")
          Exit Sub
        End If
      Case 20004
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("CdrBluffPMask.TXT") Then
          alnRpt = "CdrBluffPMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'CdrBluffPMask.TXT'.")
          Exit Sub
        End If
      Case Else
        Call TaxMsg(900, "No mask is available.")
        Exit Sub
    End Select
  End If
  
  Handle = FreeFile
'  alnRpt = "TAXPPMSK.DAT"
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
  frmVATaxPrint.Show vbModal
  alnRpt = ""
  doAlign = False

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "cmdAlign_Click", Erl)
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

Private Sub PrintPersVAStandard(RptFile As Integer, TBHandle As Integer, PTaxBill As VAPPTaxBillType, PrnCnt As Long) ', PersRec As PersonalRecType, PHandle As Integer)
  'checked OK against mask (taxppmsk.dat) on 10/21/2005
  Dim x As Long, PYearStr$
  Dim File$, LC As Integer, CustName$
  Dim WhatYear As Integer, WhatPers&
  Dim CarCount As Integer
  Dim PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim PYear As Integer
  Dim TotOth As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  WhatYear = CInt(Mid(fpcmbFile.Text, Len(fpcmbFile.Text) - 7, 4))
  
  PERC! = PTaxBill.PPTRADiscnt
  
  CustName$ = QPTrim$(PTaxBill.CustName)
  Print #RptFile, "~"
  Print #RptFile, Tab(63); "TAX YEAR: "; CStr(PTaxBill.TaxYear) 'WhatYear
  Print #RptFile, Tab(75); Using$("#####", PTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); TownName$
  Print #RptFile, Tab(5); Add1$
  Print #RptFile, Tab(5); Add2$
  Print #RptFile, Tab(5); Add3$
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " " 'added
  Print #RptFile, Tab(5); "Acct # "; Using$("#####0", PTaxBill.CustRec)
  Print #RptFile, Tab(5); CustName$
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd1)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd2)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd3) + " " + QPTrim$(PTaxBill.CustZip)
  For LC = 18 To 21
   Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 24 Starts Here
  Print #RptFile, "Personal Property"; Tab(32); Using$("#.00", PTaxBill.PersTaxRate);
  Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.PersValue);
  Print #RptFile, Tab(51); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue)); ' - PTaxBill.OverPayAmt));
  If PPTRARmvl = True Then
    Print #RptFile, Tab(63); Using$("####0.00", 0);
    Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue)) ' - PTaxBill.OverPayAmt))
  Else
    Print #RptFile, Tab(63); Using$("####0.00", PTaxBill.PPTRADiscnt);
    Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue - PTaxBill.PPTRADiscnt)) ' - PTaxBill.OverPayAmt))
  End If
  Print #RptFile, "Machinery/Tools"; Tab(32); Using$("#.00", PTaxBill.MTTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MTValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MTTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MTTaxDue)
  Print #RptFile, "Farm Equipment";
   Print #RptFile, Tab(32); Using("#.00", PTaxBill.FETaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.FEValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.FETaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.FETaxDue)
  Print #RptFile, "Mobile Homes";
   Print #RptFile, Tab(32); Using$("#.00", PTaxBill.MHTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MHValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MHTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MHTaxDue)
  Print #RptFile, "Merchant Capital";
   Print #RptFile, Tab(32); Using$("#.00", PTaxBill.MCTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MCValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MCTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MCTaxDue)
   TotOth = OldRound(PTaxBill.OptRevTax1 + PTaxBill.OptRevTax2 + PTaxBill.OptRevTax3)
  If PTaxBill.OverPayAmt > 0 And TotOth = 0 Then '6/22/06
    Print #RptFile, "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " **"
  ElseIf PTaxBill.OverPayAmt > 0 And TotOth > 0 Then '6/22/06
    Print #RptFile, "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " **"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
  ElseIf PTaxBill.OverPayAmt = 0 And TotOth > 0 Then '6/22/06
    Print #RptFile, "Other Taxes: " + Using$("#,##0.00", TotOth)
  End If
'  For LCnt = CarCount To 6: Print #RptFile, "": Next LCntTab(43);Tab(43);Tab(60);
  For LCnt = CarCount To 5: Print #RptFile, "": Next LCnt
  Print #RptFile,
  Print #RptFile,
'  Print #RptFile,'10/24/06
'  Print #RptFile,
'  Print #RptFile,
  Print #RptFile, Tab(48); "Total Tax Due ";
  If PPTRARmvl = True Then
    Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt + PTaxBill.PPTRADiscnt))
  Else
    Print #RptFile, Using$("$#######0.00", PTaxBill.TotalBillDue - PTaxBill.OverPayAmt)
  End If
  Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(PTaxBill.DueDate) ' DueDate$;
  Print #RptFile, ""
  Print #RptFile, ""
  Print #RptFile, ""
  Print #RptFile, "BN"; Using$("####0", PrnCnt)
  Print #RptFile, "~"
  
End Sub

Private Sub PrintRealVAStandard(RptFile As Integer, TBHandle As Integer, RTaxBill As VARETaxBillType, TCHandle As Integer, TaxCust As TaxCustType, PrnCnt As Long)
 'checked OK against mask (TAXREMSK.DAT) on 10/21/2005
 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
 'TAXRESTD.BI
  Dim LC As Long, RealTaxRate#
  Dim CustName As String * 45, WhatYear As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim ThisDesc As String * 28
  Dim TotOth As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  RealTaxRate# = RTaxBill.RealTaxRate
  WhatYear = CInt(Mid(fpcmbFile.Text, Len(fpcmbFile.Text) - 8, 4))
  CustName$ = RTaxBill.CustName
  Print #RptFile, "~"
  Print #RptFile, Tab(64); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", RTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " " 'added
  Print #RptFile, Tab(5); TownName$
  Print #RptFile, Tab(5); Add1$
  Print #RptFile, Tab(5); Add2$
  Print #RptFile, Tab(5); Add3$
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " " 'added
  Print #RptFile, " "
  Print #RptFile, Tab(5); "PIN:  " + QPTrim$(RTaxBill.RealPin)
  Print #RptFile, Tab(5); "ACCT: " + Using$("#####", RTaxBill.CustRec)
  Print #RptFile, Tab(5); CustName$
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd1, 35)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd2, 35)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustAdd3) + " " + RTaxBill.CustZip

  For LC = 19 To 20 'made 18 = 19
    Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(61); "NET TOTAL"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 23 Starts Here
  ThisDesc = QPTrim$(RTaxBill.RDesc1)
  Print #RptFile, ThisDesc; 'QPTrim$(RTaxBill.RDesc1);
  Print #RptFile, Tab(30); Using("#0.00", RealTaxRate#);
  If RTaxBill.RealValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", (RTaxBill.RealValue - RTaxBill.ExptValue)); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  ElseIf RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", (RTaxBill.BldgValue - RTaxBill.ExptValue));
  ElseIf RTaxBill.RealValue + RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue - (RTaxBill.ExptValue * (RTaxBill.RealValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue - (RTaxBill.ExptValue * (RTaxBill.BldgValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RTaxBill.PersValue));
  Else
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  End If
  Print #RptFile, Tab(61); Using("#####0.00", OldRound(RTaxBill.RealValue + RTaxBill.BldgValue - RTaxBill.ExptValue));
  Print #RptFile, Tab(71); Using("######0.00", OldRound(RTaxBill.TotalBillDue)) ' - RTaxBill.OverPayAmt))
  Print #RptFile, QPTrim$(RTaxBill.RDesc2)
  TotOth = OldRound(RTaxBill.OptRevTax1 + RTaxBill.OptRevTax2 + RTaxBill.OptRevTax3 + RTaxBill.LateTaxDue)
  If RTaxBill.OverPayAmt > 0 And TotOth = 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt > 0 And TotOth > 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"; Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt = 0 And TotOth > 0 Then
    Print #RptFile, Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  Else
    For LCnt = 25 To 36: Print #RptFile, "": Next LCnt
  End If
 'Lines 25 to 36 are blank
'Line 37 for Totals
'       Print #RptFile, ""
'  Print #RptFile,
  Print #RptFile, Tab(48); "Total Tax Due ";
  Print #RptFile, Using$("$######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
  Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(RTaxBill.DueDate) ' DueDate$
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, "BN"; Using$("#####", PrnCnt)
  Print #RptFile, "~"
  
End Sub

Private Sub PrintLaserPers1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaser1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim TotValue As Double
  Dim ThisOpt1Desc As String * 15
  Dim ThisOpt2Desc As String * 15
  Dim ThisOpt3Desc As String * 15
  Dim FBill&
  Dim LBill&
  Dim ThisFile$
  Dim PrintCnt As Long
  
  On Error GoTo ERRORSTUFF
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  
  dlm$ = "~"
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  ReportFile$ = StartPath$ + "/TaxPBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBillPersFile TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  
  arVATaxBillPersLaser.Head1 = QPTrim(TBDRec.TxtHead1)
  arVATaxBillPersLaser.Head2 = QPTrim(TBDRec.TxtHead2)
  arVATaxBillPersLaser.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  arVATaxBillPersLaser.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  arVATaxBillPersLaser.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  arVATaxBillPersLaser.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  arVATaxBillPersLaser.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  arVATaxBillPersLaser.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  arVATaxBillPersLaser.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  arVATaxBillPersLaser.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  arVATaxBillPersLaser.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  arVATaxBillPersLaser.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  arVATaxBillPersLaser.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  arVATaxBillPersLaser.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  arVATaxBillPersLaser.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  arVATaxBillPersLaser.LblHead4 = QPTrim(TBDRec.txtHead4)
  arVATaxBillPersLaser.LblHead5 = QPTrim(TBDRec.txtHead5)
  arVATaxBillPersLaser.LblHead6 = QPTrim(TBDRec.txtHead6)
  arVATaxBillPersLaser.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  arVATaxBillPersLaser.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxBillPersLaser.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxBillPersLaser.Image1.Visible = True
    End If
  End If
  
  OpenPersPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
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
      Get TBHandle, BillIdx(x), TBRec
    Else
      Get TBHandle, PrintThis(x), TBRec
    End If
    If OptMulti.Value = True Then GoTo MultiOpt
    If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
MultiOpt:
      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
        If TBRec.PPTRAValue > 0 Then
          If PPTRARmvl = False Then
            TotValue = OldRound(TBRec.PPTRAValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
          ElseIf PPTRARmvl = True Then
            TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
          End If
        Else
          TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        End If
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                      5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                     6               7
        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
        '                        8                   9                    10
        Print #RptHandle, TBRec.PersValue; dlm; TBRec.FEValue; dlm; TBRec.ExptValue; dlm;
        If PPTRARmvl = False Then
          '                    11                      12                                     13
          Print #RptHandle, TBRec.PPTRAValue; dlm; TBRec.PPTRADiscnt; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
        Else
          '                    11                  12                                           13
          Print #RptHandle, TBRec.PPTRAValue; dlm; 0; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt + TBRec.PPTRADiscnt); dlm;
        End If
        '                       14                 15                   16
        Print #RptHandle, TBDRec.dologo; dlm; TBRec.MHValue; dlm; TBRec.MCValue; dlm;
        '                        17                     18                          19
        Print #RptHandle, TBRec.MTValue; dlm; OldRound(TBRec.PersTaxDue); dlm; TBRec.PersTaxNet; dlm;
        '                         20                    21                    22
        Print #RptHandle, TBRec.PersTaxRate; dlm; TBRec.FETaxDue; dlm; TBRec.FETaxRate; dlm;
        '                       23                    24                    25
        Print #RptHandle, TBRec.MCTaxDue; dlm; TBRec.MCTaxRate; dlm; TBRec.MHTaxDue; dlm;
        '                       26                    27                     28
        Print #RptHandle, TBRec.MHTaxRate; dlm; TBRec.MTTaxDue; dlm; TBRec.MTTaxRate; dlm;
        '                        29                     30                     31
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        
        ThisOpt1Desc = QPTrim$(TBRec.OptRevDesc1)
        ThisOpt2Desc = QPTrim$(TBRec.OptRevDesc2)
        ThisOpt3Desc = QPTrim$(TBRec.OptRevDesc3)
        '                        32              33                 34              35             36                  37                  38
        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; " "; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm; TBRec.OverPayAmt; dlm;
        '                          39                      40
        Print #RptHandle, TBRec.PriorYrBalance; dlm; TBRec.PrintPrior
        
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  arVATaxBillPersLaser.GetName ReportFile$
  arVATaxBillPersLaser.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "PrintLaserPers1", Erl)
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

Private Sub PrintLaserReal1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaser1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim FBill&
  Dim LBill&
  Dim ThisFile$
  Dim PrintCnt As Long
  Dim dlm$
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  
  ReportFile$ = StartPath$ + "/TaxRBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Real Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBillRealFile TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  
  ARptVATempTaxBill.Head1 = QPTrim(TBDRec.TxtHead1)
  ARptVATempTaxBill.Head2 = QPTrim(TBDRec.TxtHead2)
  ARptVATempTaxBill.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  ARptVATempTaxBill.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  ARptVATempTaxBill.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  ARptVATempTaxBill.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  ARptVATempTaxBill.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  ARptVATempTaxBill.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  ARptVATempTaxBill.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  ARptVATempTaxBill.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  ARptVATempTaxBill.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  ARptVATempTaxBill.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  ARptVATempTaxBill.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  ARptVATempTaxBill.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  ARptVATempTaxBill.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  ARptVATempTaxBill.LblHead4 = QPTrim(TBDRec.txtHead4)
  ARptVATempTaxBill.LblHead5 = QPTrim(TBDRec.txtHead5)
  ARptVATempTaxBill.LblHead6 = QPTrim(TBDRec.txtHead6)
  ARptVATempTaxBill.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  ARptVATempTaxBill.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      ARptVATempTaxBill.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      ARptVATempTaxBill.Image1.Visible = True
    End If
  End If
  
  OpenRealPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
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
      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
PrintIt: '8/21/07
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                          5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                            6                            7
        Print #RptHandle, QPTrim$(TBRec.RealPin); dlm; QPTrim$(TBRec.RDesc1); dlm;
        '                        8                     9                    10
        Print #RptHandle, TBRec.RealValue; dlm; TBRec.BldgValue; dlm; TBRec.ExptValue; dlm;
        If OldRound(TBRec.RealTaxDue - TBRec.OverPayAmt) > 0 Then
          ThisRate = TBRec.RealTaxRate
        Else
          ThisRate = 0
        End If
        '                                                   11                                   12
        Print #RptHandle, OldRound(TBRec.RealValue + TBRec.BldgValue - TBRec.ExptValue); dlm; ThisRate; dlm;
        '                                   13                                  14                 15                       16
        Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm; " "; dlm; QPTrim$(TBRec.CustZip); dlm; TBDRec.dologo; dlm;
        '                        17                     18                     19
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        '                       20                   21                  22                    23                     24
        Print #RptHandle, TBRec.Opt1Desc; dlm; TBRec.Opt2Desc; dlm; TBRec.Opt3Desc; dlm; TBRec.OverPayAmt; dlm; TBRec.LateTaxDue; dlm;
        '                          25                      26
        Print #RptHandle, TBRec.PriorYrBalance; dlm; TBRec.PrintPrior
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  
  ARptVATempTaxBill.GetName ReportFile$
  ARptVATempTaxBill.startrpt
  
  ARptVATempTaxBill.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "PrintLaserReal1", Erl)
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

'Private Sub PrintExpPers()
'  Dim x As Long, y As Integer
'  Dim TaxXRec As TaxBillExportPersType
'  Dim TXHandle As Integer
'  Dim TBRec As VAPPTaxBillType
'  Dim TBHandle As Integer
'  Dim NumOfTBRecs As Long
'  Dim ReportFile As String
'  Dim RptHandle As Integer
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim TaxYear$
'  Dim FF11$
'  Dim Map$
'  Dim PersRec As PersonalRecType
'  Dim PHandle As Integer
'  Dim NumOfPRecs As Long
'  Dim BillNo&
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'  Dim ThisSSN As String * 11
'  Dim PersExemp As Double
'  Dim NextRec As Long
'  Dim ExempAmt As Double
'  Dim PersDue As Double
'  Dim TotPersVal As Double
'  Dim FF9$
'  Dim PrnCnt As Long
'
'  On Error GoTo ERRORSTUFF
'
'  PrnCnt = 0
'  FF11$ = "########.#0"
'  FF9$ = "######.#0"
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'  TaxYear$ = CStr(TaxMasterRec.TaxYear)
'  ReportFile$ = "LCPP" + TaxYear + ".TXT"
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'
''  BillNo& = fpDblSnglStartPersBill.Value
'  OpenPersTaxBillFile TBHandle, NumOfTBRecs
'
''  OpenRealTaxBillFile TBHandle, NumOfTBRecs
'  OpenPersPropFile PHandle, NumOfPRecs
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  For x = 1 To NumOfTBRecs
'    Get TBHandle, x, TBRec
'    If TBRec.BillNumber > 0 Then
'      If TBRec.TotalBillDue > 0 Then 'assign and save bill number
'        TBRec.BillNumber = BillNo&
'        TBRec.BillPrinted = True
'        Put TBHandle, x, TBRec
'        BillNo& = BillNo& + 1
'      End If
'      RSet TaxXRec.CustName = QPTrim$(TBRec.CustName)
'      RSet TaxXRec.Add1 = QPTrim$(TBRec.CustAdd1)
'      RSet TaxXRec.Add2 = QPTrim$(TBRec.CustAdd2)
'      Get TCHandle, TBRec.CustRec, TaxCust
'      RSet TaxXRec.City = QPTrim$(TaxCust.City)
'      RSet TaxXRec.State = QPTrim$(TaxCust.State)
'      RSet TaxXRec.Zip = QPTrim$(TaxCust.Zip)
'      RSet TaxXRec.CustAcct = Using$("#####0", TBRec.CustRec)
'      ThisSSN = QPTrim$(TaxCust.CSSN)
'      ThisSSN = ReplaceString(ThisSSN, "-", "")
'      Call InsertSSNDashes(ThisSSN)
'      RSet TaxXRec.SSN1 = QPTrim$(ThisSSN)
'      ThisSSN = QPTrim$(TaxCust.OSSN)
'      ThisSSN = ReplaceString(ThisSSN, "-", "")
'      Call InsertSSNDashes(ThisSSN)
'      RSet TaxXRec.SSN2 = QPTrim$(ThisSSN)
''      RSet TaxXRec.DueDate = fptxtPersDueDate.Text
'      NextRec = TaxCust.FirstPersRec
'      PersExemp = 0
'      If NextRec = 0 Then
'        RSet TaxXRec.LessRelief = Using$(FF11$, 0)
'        RSet TaxXRec.NetDue = Using$(FF11$, 0)
'        RSet TaxXRec.RepeatDesc = ""
'        RSet TaxXRec.RepeatID = ""
'        RSet TaxXRec.RepeatValue = Using$(FF11$, 0)
'        RSet TaxXRec.RepeatTaxRate = Using$("#.#0", 0)
'        RSet TaxXRec.RepeatTotTax = Using$(FF9$, 0)
'        RSet TaxXRec.RepeatTaxRelief = Using$(FF9, 0)
'        RSet TaxXRec.RepeatTotTax = Using$(FF9$, 0)
'      Else
'        Get PHandle, NextRec, PersRec
'        PersExemp = OldRound(PersExemp + PersRec.EXMPOTHR + PersRec.EXMPSENI)
'        ExempAmt = OldRound(TBRec.PersTaxRate * PersExemp)
'        PersDue = OldRound(ExempAmt + TBRec.PersTaxDue)
'        RSet TaxXRec.LessRelief = Using$(FF11$, ExempAmt)
'        RSet TaxXRec.NetDue = Using$(FF11$, TBRec.PersTaxDue)
'        RSet TaxXRec.RepeatDesc = QPTrim$(PersRec.DESC1)
'        RSet TaxXRec.RepeatID = QPTrim$(PersRec.PropPin)
'        TotPersVal = OldRound(PersRec.CVALUE + PersRec.PersVal + PersRec.MCValue + PersRec.MHValue + PersRec.MTValue)
'        RSet TaxXRec.RepeatValue = Using$(FF11$, TotPersVal)
'        RSet TaxXRec.RepeatTaxRate = Using$("#.#0", TBRec.PersTaxRate)
'        RSet TaxXRec.RepeatTotTax = Using$(FF9$, PersDue)
'        RSet TaxXRec.RepeatTaxRelief = Using$(FF9, ExempAmt)
''        RSet TaxXRec.RepeatTotTax = Using$(FF9$, TBRec.PersTaxDue)
'        RSet TaxXRec.RepeatTotTax = Using$(FF9$, OldRound(TBRec.PersTaxDue + TBRec.MCTaxDue + TBRec.MHTaxDue + TBRec.FETaxDue + TBRec.MTTaxDue + TBRec.OptRevTax1 + TBRec.OptRevTax2 + TBRec.OptRevTax3 - TBRec.PPTRADiscnt))
''        If InStr(TBRec.CustAdd1, "991") > 0 Then Stop
'      End If
'      PrnCnt = PrnCnt + 1
'      Print #RptHandle, TaxXRec.CustName; TaxXRec.Add1; TaxXRec.Add2;
'      Print #RptHandle, TaxXRec.City; TaxXRec.State; TaxXRec.Zip;
'      Print #RptHandle, TaxXRec.CustAcct; TaxXRec.SSN1; TaxXRec.SSN2;
'      Print #RptHandle, TaxXRec.DueDate; TaxXRec.TotDue; TaxXRec.LessRelief;
'      Print #RptHandle, TaxXRec.NetDue;
'      For y = 1 To 75
'        Print #RptHandle, TaxXRec.RepeatDesc; TaxXRec.RepeatID;
'        Print #RptHandle, TaxXRec.RepeatValue; TaxXRec.RepeatTaxRate;
'        Print #RptHandle, TaxXRec.RepeatTotTax; TaxXRec.RepeatTaxRelief;
'        Print #RptHandle, TaxXRec.RepeatNetTax;
'      Next y
'      Print #RptHandle, TaxXRec.RepeatDesc; TaxXRec.RepeatID;
'      Print #RptHandle, TaxXRec.RepeatValue; TaxXRec.RepeatTaxRate;
'      Print #RptHandle, TaxXRec.RepeatTotTax; TaxXRec.RepeatTaxRelief;
'      Print #RptHandle, TaxXRec.RepeatNetTax
'    End If
'  Next x
'  Close
'
''  If PrnCnt > 0 Then
''    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
''  End If
'
'  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCPP" + TaxYear + ".TXT' located in the Citipak folder.")
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintExpPers", Erl)
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
'
'End Sub

Private Sub PrintExpReal()
  Dim x As Long
  Dim TaxXRec As TaxBillExportRealType
  Dim TXHandle As Integer
  Dim TBRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear$
  Dim FF10$
  Dim Map$
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim BillNo&
  Dim PrnCnt As Long
  
  On Error GoTo ERRORSTUFF
  PrnCnt = 0
  FF10$ = "#######.#0"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TaxYear$ = CStr(TaxMasterRec.RTaxYear)
  ReportFile$ = "LCRE" + TaxYear + ".TXT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
'  BillNo& = fpDblSnglStartRealBill.Value
  
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealPropFile RHandle, NumOfRRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 Then
      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
        RSet TaxXRec.TaxBillNum = Using$("######0", TBRec.BillNumber)
        RSet TaxXRec.CustName = QPTrim$(TBRec.CustName)
        RSet TaxXRec.Add1 = QPTrim$(TBRec.CustAdd1)
        RSet TaxXRec.Add2 = QPTrim$(TBRec.CustAdd2)
        RSet TaxXRec.Add3 = QPTrim$(TBRec.CustAdd3)
        RSet TaxXRec.TaxYear = Using$("###0", TBRec.TaxYear)
        RSet TaxXRec.CustAcct = Using$("####0", TBRec.CustPin)
        If TBRec.RealPropRecord > 0 Then
          Get RHandle, TBRec.RealPropRecord, RealRec
          RSet TaxXRec.MapNum = QPTrim$(RealRec.Map)
        Else
          RSet TaxXRec.MapNum = "NA"
        End If
        RSet TaxXRec.PropDesc1 = QPTrim$(TBRec.RDesc1)
        RSet TaxXRec.TAXRATE = Using$("#.#0", TBRec.RealTaxRate)
        RSet TaxXRec.LandVal = Using$(FF10, 0) 'see Va. code
        RSet TaxXRec.BldgVal = Using$(FF10, 0) 'see Va. code
        RSet TaxXRec.RealVal = Using$(FF10$, TBRec.RealValue)
        RSet TaxXRec.CurrTaxAmt = Using$(FF10$, OldRound(TBRec.RealTaxDue - TBRec.OverPayAmt))
        RSet TaxXRec.PropDesc2 = QPTrim$(TBRec.RDesc2)
        RSet TaxXRec.PropDesc3 = QPTrim$(TBRec.RDesc3)
        RSet TaxXRec.TotTaxAmt = Using$(FF10$, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt))
        PrnCnt = PrnCnt + 1
        Print #RptHandle, TaxXRec.TaxBillNum; TaxXRec.CustName;
        Print #RptHandle, TaxXRec.Add1; TaxXRec.Add2; TaxXRec.Add3;
        Print #RptHandle, TaxXRec.TaxYear; TaxXRec.CustAcct;
        Print #RptHandle, TaxXRec.MapNum; TaxXRec.PropDesc1;
        Print #RptHandle, TaxXRec.TAXRATE; TaxXRec.LandVal;
        Print #RptHandle, TaxXRec.BldgVal; TaxXRec.RealVal;
        Print #RptHandle, TaxXRec.CurrTaxAmt; TaxXRec.PropDesc2;
        Print #RptHandle, TaxXRec.PropDesc3; TaxXRec.TotTaxAmt
      End If
    End If
  Next x
  Close
  
'  If PrnCnt > 0 Then
'    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
'  End If
  
'  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCRE" + TaxYear + ".TXT' located in the Citipak folder.")
        
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "PrintExpReal", Erl)
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

Private Sub PrintPersLaserItemized()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaserItemized
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim TotValue As Double
  Dim ThisOpt1Desc As String * 15
  Dim ThisOpt2Desc As String * 15
  Dim ThisOpt3Desc As String * 15
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim NumOfPers As Integer
  Dim NextRec As Long
  Dim thisVin As String
  Dim y As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NumOfOpts As Integer
  Dim FBill&
  Dim LBill&
  Dim PrintCnt As Long
  Dim ThisFile$
  
  On Error GoTo ERRORSTUFF
  
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  NumOfOpts = 0
  If QPTrim$(TaxMasterRec.POptRev1) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(TaxMasterRec.POptRev2) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(TaxMasterRec.POptRev3) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxPLsrItem.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  BillNo& = fpDblSnglFirstBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenLaserPersItemized TBDHandle
  Get #TBDHandle, 1, TBDRec
  Close TBDHandle
  
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxLaserPersItemized.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxLaserPersItemized.Image1.Visible = True
    End If
  End If
  OpenTaxCustFile TCHandle, NumOfTCRecs
  ReDim VinArray(1 To 1) As String
  OpenPersPropFile PHandle, NumOfPRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)

  OpenPersPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
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
    NumOfPers = 0
    If OptMulti.Value = False Then
'      Get TBHandle, x, TBRec
      Get TBHandle, BillIdx(x), TBRec '8/6/08
    Else
      Get TBHandle, PrintThis(x), TBRec
    End If
    Get TCHandle, TBRec.CustPin, TaxCust
    If OptMulti.Value = True Then GoTo MultiOpt
'    If TBRec.BillNumber > 0 Then
      If TBRec.BillNumber >= FBill And TBRec.BillNumber <= LBill Then
MultiOpt:
        If TBRec.PPTRAValue > 0 Then
          TotValue = OldRound(TBRec.PPTRAValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        Else
          TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        End If
        GoSub GetBarCodeData
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                      5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                     6               7
        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
        '                        8                   9                    10
        Print #RptHandle, TBRec.PersValue; dlm; TBRec.FEValue; dlm; TBRec.ExptValue; dlm;
        '                    11                      12                        13
        Print #RptHandle, TBRec.PPTRAValue; dlm; TBRec.PPTRADiscnt; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
        '                       14                 15                   16
        Print #RptHandle, TBDRec.dologo; dlm; TBRec.MHValue; dlm; TBRec.MCValue; dlm;
        '                      17                            18                       19
        Print #RptHandle, TBRec.MTValue; dlm; OldRound(TBRec.PersTaxDue); dlm; TBRec.PersTaxNet; dlm;
        '                         20                    21                    22
        Print #RptHandle, TBRec.PersTaxRate; dlm; TBRec.FETaxDue; dlm; TBRec.FETaxRate; dlm;
        '                       23                    24                    25
        Print #RptHandle, TBRec.MCTaxDue; dlm; TBRec.MCTaxRate; dlm; TBRec.MHTaxDue; dlm;
        '                       26                    27                     28
        Print #RptHandle, TBRec.MHTaxRate; dlm; TBRec.MTTaxDue; dlm; TBRec.MTTaxRate; dlm;
        '                        29                     30                     31
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        ThisOpt1Desc = QPTrim$(TBRec.OptRevDesc1)
        ThisOpt2Desc = QPTrim$(TBRec.OptRevDesc2)
        ThisOpt3Desc = QPTrim$(TBRec.OptRevDesc3)
        '                        32              33                 34              35             36                  37
        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; BZip; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm;
        '                           38                           39                            40
        Print #RptHandle, QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
        '                           41                           42                 43
        Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm; ""; dlm;
        '                           44                           45                             46
        Print #RptHandle, QPTrim(TBDRec.txtPgph0); dlm; QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm;
        '                           47                           48                            49
        Print #RptHandle, QPTrim(TBDRec.txtPgph3); dlm; QPTrim(TBDRec.txtPgph4); dlm; "                     "; dlm;
        '                           50               51       52
        Print #RptHandle, "                  "; dlm; ""; dlm; ""; dlm;
        '                           53                           54                            55
        Print #RptHandle, QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
        '                 56       57
        Print #RptHandle, ""; dlm; ""; dlm;
        NextRec = TaxCust.FirstPersRec
        If NextRec > 0 Then
          NumOfPers = NumOfPers + 1
          Get PHandle, NextRec, PersRec
          '                           58                     59                    60
          Print #RptHandle, QPTrim$(PersRec.Vin); dlm; QPTrim$(PersRec.MakeMod); dlm; PersRec.PersVal; dlm;
          '                       61                     62                   63                  64                  65                 66                      67
          Print #RptHandle, PersRec.MTValue; dlm; PersRec.MCValue; dlm; PersRec.CVALUE; dlm; PersRec.MHValue; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
          NextRec = PersRec.NextRec
        Else
          '                 58       59       60       61       62       63       64          65                    66                   67
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
        End If
        Do While NextRec > 0
          Get PHandle, NextRec, PersRec
          NumOfPers = NumOfPers + 1
          '                         0                         1                           2                                3
          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm; QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
          '                             4                        5               6            7        8        9        10
          Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm; TotValue; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 11       12                          13
          Print #RptHandle, ""; dlm; ""; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
          '                 14       15       16       17       18
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 19       20       21       22       23       24       25       26
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 27       28       29       30       31           32                  33                34
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm;
          '                 35       36       37                38                            39                            40
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
          '                            41                           42
          Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm;
          '                 43                 44                           45                            46
          Print #RptHandle, ""; dlm; QPTrim(TBDRec.txtPgph0); dlm; QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm;
          '                            47                 48       49       50
          Print #RptHandle, QPTrim(TBDRec.txtPgph3); dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 51       52                53                               54                        55
          Print #RptHandle, ""; dlm; ""; dlm; QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
          '                 56       57              58
          Print #RptHandle, ""; dlm; ""; dlm; QPTrim$(PersRec.Vin); dlm;
          '                         59                           60                   61                       62
          Print #RptHandle, QPTrim$(PersRec.MakeMod); dlm; PersRec.PersVal; dlm; PersRec.MTValue; dlm; PersRec.MCValue; dlm;
          '                        63                   64                65                  66                     67
          Print #RptHandle, PersRec.CVALUE; dlm; PersRec.MHValue; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
          
          If NumOfOpts = 1 Then
            If NumOfPers = 13 Then NumOfPers = 0
          ElseIf NumOfOpts = 2 Then
            If NumOfPers = 12 Then NumOfPers = 0
          ElseIf NumOfOpts = 3 Then
            If NumOfPers = 11 Then NumOfPers = 0
          Else
            If NumOfPers = 14 Then NumOfPers = 0
          End If
          NextRec = PersRec.NextRec
        Loop
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
'    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  arVATaxLaserPersItemized.Show
  
  Exit Sub
  
GetBarCodeData:
  If TBDRec.UseBarCode = False Then
    BZip = ""
    Return
  ElseIf TBDRec.UseBarCode = True Then
    Get TCHandle, TBRec.CustPin, TaxCust
    If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
      BZip = ""
    Else
      BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
    End If
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxReprintPosted", "PrintPersLaserItemized", Erl)
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

Private Sub PrintMdltwnReal()
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim x As Long, RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&, TotOpt As Double
  Dim LBill&
  Dim ThisFile$
  Dim PrintCnt As Long

  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenRealPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
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
  
  File$ = StartPath$ + "/TxBMdltwnRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  GoSub LoadHeaders
  
  Tab1 = 44 - Len(TownName) / 2
  Tab2 = 44 - Len(Add1) / 2
  Tab3 = 44 - Len(Add3) / 2
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoTo PrintIt '8/21/07
    End If
    If OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt) > 0 Then
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
PrintIt: '8/21/07
'        Get TCHandle, TaxBill.CustRec, TaxCust
        CustName$ = QPTrim$(TaxBill.CustName)
        DueDate = MakeRegDate(TaxBill.DueDate)
        RealTaxRate# = TaxBill.RealTaxRate
        TotOpt = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
        Print #RptFile, "                                R E A L   E S T A T E"
        Print #RptFile, "                                 T A X   N O T I C E"
        Print #RptFile, Tab(Tab1); TownName
        Print #RptFile, Tab(Tab2); Add1
        Print #RptFile, Tab(Tab3); Add3
        Print #RptFile,
        Print #RptFile, "            VALUATION AMOUNT: "; Using$("$##,###,###.00", OldRound#(TaxBill.RealValue + TaxBill.BldgValue));
        Print #RptFile, Tab(50); "ACCT. #: "; CStr(TaxBill.CustRec)
        Print #RptFile, "                   EXEMPTION: "; Using$("$##,###,###.00", TaxBill.ExptValue);
        Print #RptFile, Tab(50); "PIN. #: "; QPTrim$(TaxBill.RealPin)
        Print #RptFile, "         LATE PENALTY AMOUNT: "; Using$("$##,###,###.00", TaxBill.LateTaxDue);
        Print #RptFile, Tab(50); "RECPT #: "; Using$("#####0", TaxBill.BillNumber)
'        Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue);
        If TotOpt > 0 Then
          Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue) + "*";
        Else
          Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue);
        End If
        Print #RptFile, Tab(50); "TAX RATE %: "; Using$("#0.0000", RealTaxRate#)
        Print #RptFile, Tab(50); "TAX YEAR: "; CStr(TaxBill.TaxYear)
        Print #RptFile, Tab(50); "DUE DATE: "; DueDate
'        Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45)
        If TotOpt > 0 Then
          Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45); Tab(50); "*PLUS " + QPTrim$(Using$("$##,##0.00", TotOpt)) + " IN ADDED TAXES."
        Else
          Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45)
        End If
        Print #RptFile, Tab(11); Left$(QPTrim$(TaxBill.CustAdd1), 35)
        Print #RptFile, Tab(11); Left$(QPTrim$(TaxBill.CustAdd2), 35)
        Print #RptFile, Tab(11); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
        Print #RptFile,
        Print #RptFile,
        Print #RptFile, Tab(31); "T H A N K   Y O U ! ! !"
        Print #RptFile,
        Print #RptFile, "~"
      End If
    End If
NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  ViewPrint File$, "Real Property Tax Bills", True
  Exit Sub
  
LoadHeaders:
  WorkName = ""
  WordLen = Len(TownName)
  For y = 1 To WordLen
    ch = Mid(TownName, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  TownName = WorkName
  
  WorkName = ""
  WordLen = Len(Add1)
  For y = 1 To WordLen
    ch = Mid(Add1, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add1 = WorkName
  
  WorkName = ""
  WordLen = Len(Add2)
  For y = 1 To WordLen
    ch = Mid(Add2, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add2 = WorkName
  
  WorkName = ""
  WordLen = Len(Add3)
  For y = 1 To WordLen
    ch = Mid(Add3, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add3 = WorkName
  
  Return

End Sub

Private Sub PrintMdltwnPers()
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
  Dim x As Long, PersTaxRate#
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim TownName$, Add1$, Add2$, Add3$
  Dim VehDesc$, PrnCnt As Integer
  Dim TaxAmt#, TotOpt As Double
  Dim Tab1 As Integer, Tab2 As Integer, Tab3 As Integer, Tab4 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&, ThisFile$
  Dim LBill&, PrintCnt As Integer
  
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
'  OpenPersPropFile PHandle, NumOfPRecs
'  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
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
  
  File$ = StartPath$ + "/TxBStandPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  Tab1 = 40 - Len(TownName) / 2
  Tab2 = 40 - Len(Add1) / 2
  Tab3 = 40 - Len(Add3) / 2
  
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill '8/6/08
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoTo PrintThis '8/21/07
    End If
'    If TaxBill.BillPrinted = False Then GoTo Natta
'    If TaxBill.BillNumber > 0 Then
    If OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt) > 0 Then
      If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
PrintThis:
'        Get TCHandle, TaxBill.CustRec, TaxCust
        WhatYear = TaxBill.TaxYear
        DueDate$ = MakeRegDate(TaxBill.DueDate)
        GoSub PrintIt
      End If
    End If
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  
  ViewPrint File$, "Personal Property Tax Bills", True
  Exit Sub

PrintIt:
  CustName$ = RTrim$(TaxBill.CustName)
  CustName$ = LTrim$(CustName$)
  PersTaxRate# = TaxBill.PersTaxRate
  TotOpt = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
  
  Print #RptFile,
  Print #RptFile, Tab(Tab1); TownName
  Print #RptFile, Tab(Tab2); Add1
  Print #RptFile, Tab(Tab3); Add3
  Print #RptFile, Tab(27); "PERSONAL PROPERTY TAX BILL"
  Print #RptFile, Tab(30);
  For LC = 6 To 8
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(10); "ACCT # "; Using$("######0", TaxBill.CustRec);
  Print #RptFile, Tab(63); "BILL # "; Using$("######0", TaxBill.BillNumber)
  Print #RptFile, Tab(10); Left$(CustName$, 25);
  Print #RptFile, Tab(63); "TAX YEAR: "; CStr(WhatYear)
  Print #RptFile, Tab(10); Left$(QPTrim$(TaxBill.CustAdd1), 25);
  Print #RptFile, Tab(63); "TAX RATE: "; Using("#0.##0", PersTaxRate#) + "%"
  Print #RptFile, Tab(10); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptFile, Tab(10); QPTrim$(TaxBill.CustAdd3); " "; QPTrim(TaxBill.CustZip)
  For LC = 14 To 17
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(37); "PROPERTY"; Tab(51); "   TAX"; Tab(61); "   PPTRA"
  Print #RptFile, Tab(38); "  VALUE"; Tab(51); "AMOUNT"; Tab(61); "DISCOUNT"; Tab(71); "TOTAL DUE"
  'Line 23 Starts Here
  Print #RptFile, Tab(2); "Personal Property";
  Print #RptFile, Tab(38); Using$("###,##0", TaxBill.PersValue);
  Print #RptFile, Tab(47); Using$("###,##0.00", TaxBill.PersTaxDue);
  Print #RptFile, Tab(59); Using("###,##0.00", TaxBill.PPTRADiscnt);
  If TotOpt > 0 Then
    Print #RptFile, Tab(70); Using("###,##0.00", TaxBill.TotalBillDue) + "*"
  Else
    Print #RptFile, Tab(70); Using("###,##0.00", TaxBill.TotalBillDue)
  End If

  Print #RptFile, ""
  Print #RptFile, ""
  If TotOpt > 0 Then
    Print #RptFile, "*Includes " + QPTrim$(Using("$##,##0.00", TotOpt)) + " in added taxes."; Tab(40); "Total Tax Due by "; DueDate$;
    Print #RptFile, Tab(69); Using("$###,##0.00", TaxBill.TotalBillDue) + "*"
  Else
    Print #RptFile, Tab(40); "Total Tax Due by "; DueDate$;
    Print #RptFile, Tab(69); Using("$###,##0.00", TaxBill.TotalBillDue)
  End If
   
  Print #RptFile,
  Print #RptFile,

  Return

End Sub
Private Sub PrintCdrBluffReal()
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer, PrnCnt As Integer
  Dim NumOfRRecs As Long, WhatYear As Integer
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&, ThisFile$, OptTot As Double
  Dim LBill&, PrintCnt As Integer
  
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)

  OpenRealPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
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
  
  File$ = StartPath$ + "/TxBCdrBluffRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile

  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False

  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill '8/6/08
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoTo PrintIt '8/21/07
    End If
    DueDate$ = MakeRegDate(TaxBill.DueDate)
    RealTaxRate# = TaxBill.RealTaxRate
    CustName$ = QPTrim$(TaxBill.CustName)
    If TaxBill.BillNumber < 0 Then GoTo NotThisOne
    If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
PrintIt: '8/21/07
      CustName$ = QPTrim$(TaxBill.CustName)
      WhatYear = TaxBill.TaxYear
  'Must Calc Late Fee Here
      OptTot = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
      Print #RptFile, "~"
      Print #RptFile, Tab(50); CStr(WhatYear); Tab(78); Using("########", TaxBill.BillNumber)
      Print #RptFile,
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(28); Using("#.##", RealTaxRate#);
      Print #RptFile, Tab(36); Using("###,###,###", TaxBill.RealValue);
      Print #RptFile, Tab(48); Using("##,###,###", TaxBill.BldgValue);
      Print #RptFile, Tab(61); Using("##,###,###", TaxBill.RealValue + TaxBill.BldgValue);
      Print #RptFile, Tab(75); Using("###,###.##", TaxBill.RealTaxDue);
      Print #RptFile, Tab(90); QPTrim$(TaxBill.Comment) + "%"
      Print #RptFile, " "
      Print #RptFile, Tab(68); DueDate$
      Print #RptFile, QPTrim$(TaxBill.RDesc1)
      Print #RptFile, QPTrim$(TaxBill.RDesc2)
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, Tab(7); "ACCT # "; CStr(TaxBill.CustRec)
      Print #RptFile, Tab(7); Left$(CustName$, 45)
      Print #RptFile, Tab(7); Left$(TaxBill.CustAdd1, 35)
      Print #RptFile, Tab(7); Left$(TaxBill.CustAdd2, 35)
      Print #RptFile, Tab(7); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
      Print #RptFile,
'      Print #RptFile, "BN"; Using("#####", PrnCnt)
      If OptTot = 0 Then
        Print #RptFile, "BN"; Using("#####", PrnCnt)
      Else
        Print #RptFile, "BN"; Using("#####", PrnCnt); Tab(54); "Tax Due includes " + QPTrim$(Using$("$##,##0.00", OptTot)) + " in other taxes."
      End If
      Print #RptFile, "~"
    End If
NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True

  Close
  
  ViewPrint File$, "Real Property Tax Bills", True

End Sub
Private Sub PrintCdrBluffPers()
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
'  Dim BillInfo As VAPPTaxBillInfoType
'  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, PersTaxRate#
'  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile# ', WhatPers&
'  Dim CarCount As Integer
  Dim TownName$ ', Add1$, Add2$, Add3$
'  Dim PHandle As Integer, PPTRAVal#
'  Dim NumOfPRecs As Long, PPTRADiscount#
'  Dim PersRec As PersonalRecType,VehDesc$
'  Dim PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
'  Dim TotOth As Double
  Dim DueDate$, ThisFile$
  Dim FBill&, OptTot As Double
  Dim LBill&, PrintCnt As Integer
  
  ThisFile$ = MyPath + QPTrim$(fpcmbFile.Text)
  FBill = fpDblSnglFirstBill.Value
  LBill = fpDblSnglLastBill.Value
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
'  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
'  OpenPersTaxBillFile TBHandle, NumOfTBRecs
'  OpenPersBillInfoFile BIHandle
  OpenPersPostedReprintFile TBHandle, NumOfTBRecs, ThisFile
  NumOfTBRecs = BillCnt '8/6/08
'  Get BIHandle, 1, BillInfo
'  Close BIHandle
  
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
  
  File$ = StartPath$ + "/TxCdrBluffPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
'  PrnCnt = fpDblSnglPersFirstBill
  For x = 1 To NumOfTBRecs
    If OptMulti.Value = False Then
'      Get TBHandle, x, TaxBill
      Get TBHandle, BillIdx(x), TaxBill
    Else
      Get TBHandle, PrintThis(x), TaxBill
      GoSub PrintIt '8/21/07
    End If
    If TaxBill.BillNumber < 0 Then GoTo Natta
    If TaxBill.BillNumber >= FBill And TaxBill.BillNumber <= LBill Then
PrintThis: '8/21/07
      Get TCHandle, TaxBill.CustRec, TaxCust
      WhatYear = TaxBill.TaxYear
      DueDate$ = MakeRegDate(TaxBill.DueDate)
      GoSub PrintIt
'      PrnCnt = PrnCnt + 1
    End If
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  
  ViewPrint File$, "Personal Property Tax Bills", True
  Exit Sub

PrintIt:
   CustName$ = QPTrim$(TaxBill.CustName)
   PersTaxRate# = TaxBill.PersTaxRate
   Print #RptFile, "~"; Tab(34); Str$(WhatYear); " PERSONAL PROPERTY"
   Print #RptFile, Tab(5); TownName$
   Print #RptFile, Tab(5); Add1$
   Print #RptFile, Tab(5); Add3$
   Print #RptFile, Tab(5); "   "
   Call InsertSSNDashes(TaxCust.CSSN)
   Print #RptFile, Tab(10); TaxCust.CSSN; Tab(65); "PP"; Using("##.###", PersTaxRate#)
   Print #RptFile, " "
   Print #RptFile, " "

  'Line 30 to 35 here to print vehicles
'    CarCount = 0
'    WhatPers& = TaxCust.FirstPersRec
'    Do
'      Get PHandle, WhatPers&, PersRec
''      PYear$ = Right$(PersRec.Desc5, 4)
'      PYear = PersRec.TaxBillYear
'      If Left$(PersRec.Desc5, 1) = "Y" Then
'        If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
'          PPTRAVal# = TaxMasterRec.MaxVehTaxVal
'        Else
'          PPTRAVal# = OldRound#(PersRec.PersVal)
'        End If
'        If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
'          PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
'        Else
'          PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (TaxMasterRec.PPTRADisc / 100)) * PersTaxRate#))
'        End If
'      Else
'        PPTRADiscount# = 0
'      End If
'      VehDesc$ = QPTrim$(PersRec.Desc4) + " " + Left$(PersRec.DESC2, 22) + "   " + Left$(PersRec.Desc5, 1)
'      VehDesc$ = QPTrim$(VehDesc$)
'      CarCount = CarCount + 1
'      If PersRec.PersVal <= 0 Then
'        TaxAmt# = (MHTaxRate# / 100) * PersRec.MHValue
'        Print #RptFile, VehDesc$;
'        Print #RptFile, Tab(33); Using("##,###,###", PersRec.MHValue);
'        Print #RptFile, Tab(44); Using("###,###.##", TaxAmt#);
'        Print #RptFile, Tab(54); Using("##,###.##", PPTRADiscount#);
'        Print #RptFile, Tab(64); Using("##,###.##", TaxAmt#)
'      Else
'        TaxAmt# = (PersTaxRate# / 100) * PersRec.PersVal
'        Print #RptFile, " "
'        Print #RptFile, " "
'        Print #RptFile, " "
'        Print #RptFile, Tab(54); Using("#####.##", PPTRADiscount#);
'        Print #RptFile, Tab(64); Using("##,###.##", OldRound#(TaxAmt# - PPTRADiscount#))
'      End If
'      If (CarCount >= 5) And (PersRec.NextRec > 0) Then
'        Print #RptFile, " "
'        Print #RptFile, ""
'        Print #RptFile, Tab(10); Using("#####", PrnCnt);
'        Print #RptFile, Tab(36); DueDate$; Tab(66); "CONT'D"                 '; USING "$$#####,#.##"; TaxBill.TotalBillDue
'        Print #RptFile,
'        Print #RptFile, Tab(9); CustName$
'        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd1)
'        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd2)
'        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
'        Print #RptFile,
'        Print #RptFile,
'        Print #RptFile, "~"
'        Print #RptFile, "~"
'        Print #RptFile, Tab(5); TownName$
'        Print #RptFile, Tab(5); Add1$
'        Print #RptFile, Tab(5); Add2$
'        Print #RptFile, Tab(5); Add3$;
'        Print #RptFile, Tab(10); QPTrim$(TaxCust.CSSN)
'        Print #RptFile, " "
'        Print #RptFile, '"Vehicle Listing Continued ..."
'        CarCount = 0
'      ElseIf (CarCount >= 5) And (PersRec.NextRec <= 0) Then
'        Print #RptFile, ""
'      '  CarCount = 0
'      End If
'      WhatPers& = PersRec.NextRec
'    Loop While WhatPers& > 0

  ' Finish the bill up here
'    If CarCount < 5 Then
     OptTot = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
     If OptTot > 0 Then
       Print #RptFile, "Other Taxes:"; Tab(64); Using("##,###.##", OptTot#)
     End If
     
     If OptTot > 0 Then
       For LCnt = 1 To 5
         Print #RptFile, ""
       Next
     Else
       For LCnt = 1 To 6
         Print #RptFile, ""
       Next
     End If
'
    Print #RptFile, ""
    Print #RptFile, Tab(10); Using("#####", TaxBill.BillNumber);
    Print #RptFile, Tab(36); DueDate$; Tab(62); Using("$###,###.##", TaxBill.TotalBillDue)
    Print #RptFile,
    Print #RptFile, Tab(9); CustName$
    Print #RptFile, Tab(9); TaxBill.CustAdd1
    Print #RptFile, Tab(9); TaxBill.CustAdd2
    Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
    Print #RptFile,
    Print #RptFile,
    Print #RptFile, "~"
  Return

End Sub
Private Sub PrintHalifaxStandardPersonal(RptFile As Integer, TBHandle As Integer, PTaxBill As VAPPTaxBillType, PrnCnt As Long)
 'TAXPPSTD.BI
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PPTRAVal#
  Dim PPTRADiscount#
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim MultiYear As Integer
  Dim TotOth As Double
  Dim PrintComments As String
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  WhatYear = CInt(Mid(fpcmbFile.Text, Len(fpcmbFile.Text) - 7, 4))
  MultiYear = TaxMasterRec.MultiYear
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  If WhatYear = 1999 Then PERC! = 27.5
  If WhatYear = 2000 Then PERC! = 47.5
  If WhatYear >= 2001 Then PERC! = TaxMasterRec.PPTRADisc
  
  CustName$ = QPTrim$(PTaxBill.CustName)
  Print #RptFile, "~"
  Print #RptFile, Tab(63); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", PTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); "Acct # "; Using$("#####0", PTaxBill.CustRec)
  Print #RptFile, Tab(5); CustName$
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd1)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd2)
  Print #RptFile, Tab(5); QPTrim$(PTaxBill.CustAdd3) + " " + QPTrim$(PTaxBill.CustZip)
'  For LC = 18 To 21
  For LC = 19 To 21 'added
   Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 24 Starts Here
  Print #RptFile, "Personal Property"; Tab(32); Using$("#.00", PTaxBill.PersTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.PersValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.PersTaxDue);
   Print #RptFile, Tab(63); Using$("####0.00", PTaxBill.PPTRADiscnt);
   Print #RptFile, Tab(72); Using$("#####0.00", OldRound(PTaxBill.PersTaxDue - PTaxBill.PPTRADiscnt)) ' - PTaxBill.OverPayAmt))
   
  Print #RptFile, "Machinery/Tools"; Tab(32); Using$("#.00", PTaxBill.MTTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MTValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MTTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MTTaxDue)
  Print #RptFile, "Farm Equipment";
   Print #RptFile, Tab(32); Using("#.00", PTaxBill.FETaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.FEValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.FETaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.FETaxDue)
  Print #RptFile, "Mobile Homes";
   Print #RptFile, Tab(32); Using$("#.00", PTaxBill.MHTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MHValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MHTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MHTaxDue)
  Print #RptFile, "Merchant Capital";
   Print #RptFile, Tab(32); Using$("#.00", PTaxBill.MHTaxRate);
   Print #RptFile, Tab(37); Using$("#####0.00", PTaxBill.MCValue);
   Print #RptFile, Tab(51); Using$("#####0.00", PTaxBill.MCTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", PTaxBill.MCTaxDue)
   TotOth = OldRound(PTaxBill.OptRevTax1 + PTaxBill.OptRevTax2 + PTaxBill.OptRevTax3)
   If PTaxBill.OverPayAmt > 0 And TotOth = 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(43); "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " **"
   ElseIf PTaxBill.OverPayAmt > 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(30); "* Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", PTaxBill.OverPayAmt)) + " *"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
   ElseIf PTaxBill.OverPayAmt = 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
   Else
     Print #RptFile, " PPTRA Vehicle Information"
   End If
 'Line 30 to 35 here to print vehicles
  For LCnt = CarCount To 5: Print #RptFile, "": Next LCnt

   ' Finish the bill up here
  Print #RptFile,
  Print #RptFile, Tab(48); "Total Tax Due ";
  Print #RptFile, Using$("$#######0.00", OldRound(PTaxBill.TotalBillDue - PTaxBill.OverPayAmt))
  Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(PTaxBill.DueDate)
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, " " 'added
  Print #RptFile, " " 'added
  Print #RptFile, "BN"; Using$("####0", PrnCnt)
  Print #RptFile, "~"

End Sub

Private Sub PrintHalifaxStandardReal(RptFile As Integer, TBHandle As Integer, RTaxBill As VARETaxBillType, PrnCnt As Long)
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim RealTaxRate#
  Dim RYear As Integer, RYearStr$
  Dim LC As Integer
  Dim CustName As String * 45, WhatYear As Integer
  Dim WhatReal&
  Dim TownName$, Add1$, Add2$, Add3$
  Dim TaxAmt#, LCnt As Integer
  Dim ThisDesc As String * 28
  Dim TotOth As Double
  
  RealTaxRate# = RTaxBill.RealTaxRate
  WhatYear = RTaxBill.TaxYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  'Must Calc Late Fee Here
  Print #RptFile, "~"
  Print #RptFile, Tab(64); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", RTaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); "PIN:  " + QPTrim$(RTaxBill.RealPin)
  Print #RptFile, Tab(5); "ACCT: " + Using$("#####", RTaxBill.CustRec)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustName)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd1, 35)
  Print #RptFile, Tab(5); Left$(RTaxBill.CustAdd2, 35)
  Print #RptFile, Tab(5); QPTrim$(RTaxBill.CustAdd3) + " " + RTaxBill.CustZip

  For LC = 19 To 20 'made 18 = 19
    Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(61); "NET TOTAL"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
  'Line 23 Starts Here
  ThisDesc = QPTrim$(RTaxBill.RDesc1)
  Print #RptFile, ThisDesc; 'QPTrim$(RTaxBill.RDesc1);
  Print #RptFile, Tab(30); Using("#0.00", RealTaxRate#);
  If RTaxBill.RealValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", (RTaxBill.RealValue - RTaxBill.ExptValue)); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  ElseIf RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", (RTaxBill.BldgValue - RTaxBill.ExptValue));
  ElseIf RTaxBill.RealValue + RTaxBill.BldgValue > RTaxBill.ExptValue Then
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue - (RTaxBill.ExptValue * (RTaxBill.RealValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue - (RTaxBill.ExptValue * (RTaxBill.BldgValue / (RTaxBill.RealValue + RTaxBill.BldgValue)))); ' - RRTaxBill.PersValue));
  Else
    Print #RptFile, Tab(37); Using("######0.00", RTaxBill.RealValue); ' - RRTaxBill.PersValue));
    Print #RptFile, Tab(50); Using("#####0.00", RTaxBill.BldgValue);
  End If
  Print #RptFile, Tab(61); Using("#####0.00", OldRound(RTaxBill.RealValue + RTaxBill.BldgValue - RTaxBill.ExptValue));
  Print #RptFile, Tab(71); Using("######0.00", OldRound(RTaxBill.TotalBillDue)) ' - RTaxBill.OverPayAmt))
  Print #RptFile, QPTrim$(RTaxBill.RDesc2)
  TotOth = OldRound(RTaxBill.OptRevTax1 + RTaxBill.OptRevTax2 + RTaxBill.OptRevTax3 + RTaxBill.LateTaxDue)
  If RTaxBill.OverPayAmt > 0 And TotOth > 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"; Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt > 0 And TotOth = 0 Then
    Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", RTaxBill.OverPayAmt)) + " **"
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  ElseIf RTaxBill.OverPayAmt = 0 And TotOth > 0 Then
    Print #RptFile, Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
    For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
  Else
    For LCnt = 25 To 36: Print #RptFile, "": Next LCnt
  End If
 'Lines 25 to 36 are blank
'Line 37 for Totals
  Print #RptFile, Tab(48); "Total Tax Due ";
  Print #RptFile, Using$("$#######0.00", OldRound(RTaxBill.TotalBillDue - RTaxBill.OverPayAmt))
  Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(RTaxBill.DueDate)
  Print #RptFile,
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, "BN"; Using$("#####", PrnCnt)
  Print #RptFile, "~"

End Sub

