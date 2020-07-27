VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLChangeAcctMeth 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Accounting Method"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBLChangeAcctMeth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbAcctMeth 
      Height          =   375
      Left            =   5415
      TabIndex        =   0
      Tag             =   $"frmBLChangeAcctMeth.frx":08CA
      Top             =   1245
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
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
      Style           =   0
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
      ColDesigner     =   "frmBLChangeAcctMeth.frx":0A1F
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   540
      Left            =   5760
      TabIndex        =   28
      Top             =   7344
      Width           =   876
      _Version        =   131072
      _ExtentX        =   1545
      _ExtentY        =   952
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
      MaxWidth        =   2160
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
      HideOnInactiveApp=   -1  'True
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin EditLib.fpText fptxtPenRevGL 
      Height          =   300
      Left            =   3972
      TabIndex        =   1
      Tag             =   $"frmBLChangeAcctMeth.frx":0CA6
      Top             =   2304
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   529
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
   Begin VB.Frame Frame2 
      BackColor       =   &H008F8265&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2412
      Left            =   10704
      TabIndex        =   15
      Top             =   3504
      Width           =   684
      Begin fpBtnAtlLibCtl.fpBtn cmdClearX 
         Height          =   1080
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "Press this button to clear all rows of 'X's. "
         Top             =   1290
         Width           =   585
         _Version        =   131072
         _ExtentX        =   1032
         _ExtentY        =   1905
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
         ButtonDesigner  =   "frmBLChangeAcctMeth.frx":0D42
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdAllX 
         Height          =   1080
         Left            =   45
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "Press this button to insert an 'X' in all rows with a category listed."
         Top             =   90
         Width           =   585
         _Version        =   131072
         _ExtentX        =   1032
         _ExtentY        =   1905
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
         ButtonDesigner  =   "frmBLChangeAcctMeth.frx":0F21
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008F8265&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5424
      TabIndex        =   11
      Top             =   2976
      Width           =   4716
      Begin fpBtnAtlLibCtl.fpBtn cmdClearCash 
         Height          =   360
         Left            =   3270
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   $"frmBLChangeAcctMeth.frx":10FE
         Top             =   90
         Width           =   1260
         _Version        =   131072
         _ExtentX        =   2222
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmBLChangeAcctMeth.frx":11B1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClearAR 
         Height          =   360
         Left            =   1680
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   $"frmBLChangeAcctMeth.frx":138F
         Top             =   90
         Width           =   1305
         _Version        =   131072
         _ExtentX        =   2302
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmBLChangeAcctMeth.frx":1451
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClearRev 
         Height          =   360
         Left            =   48
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   $"frmBLChangeAcctMeth.frx":162D
         Top             =   96
         Width           =   1404
         _Version        =   131072
         _ExtentX        =   2476
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmBLChangeAcctMeth.frx":16E3
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3648
      Left            =   288
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   $"frmBLChangeAcctMeth.frx":18C0
      Top             =   3504
      Width           =   10344
      _Version        =   196613
      _ExtentX        =   18246
      _ExtentY        =   6435
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   13684944
      MaxCols         =   6
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmBLChangeAcctMeth.frx":19FA
      VisibleCols     =   6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   7104
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "Press this button to exit this screen. The program does not trap for unsaved changes."
      Top             =   8016
      Width           =   2940
      _Version        =   131072
      _ExtentX        =   5186
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":3401
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSaveGLNums 
      Height          =   540
      Left            =   7104
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   $"frmBLChangeAcctMeth.frx":35DF
      Top             =   7392
      Width           =   2940
      _Version        =   131072
      _ExtentX        =   5186
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":3679
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGL 
      Height          =   540
      Left            =   3360
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   $"frmBLChangeAcctMeth.frx":3860
      Top             =   7392
      Width           =   1596
      _Version        =   131072
      _ExtentX        =   2815
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":3909
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   870
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "Click on this button to activate informational balloons for each field."
      ToolTipText     =   "Click on this button to activate help balloons."
      Top             =   8010
      Width           =   2355
      _Version        =   131072
      _ExtentX        =   4154
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":3AE7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSaveAcctMeth 
      Height          =   444
      Left            =   7668
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   $"frmBLChangeAcctMeth.frx":3CCA
      Top             =   1200
      Width           =   3132
      _Version        =   131072
      _ExtentX        =   5524
      _ExtentY        =   783
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":3D65
   End
   Begin EditLib.fpText fptxtPenARGL 
      Height          =   300
      Left            =   5556
      TabIndex        =   2
      Tag             =   $"frmBLChangeAcctMeth.frx":3F53
      Top             =   2304
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   529
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
   Begin EditLib.fpText fptxtPenCashGL 
      Height          =   300
      Left            =   7140
      TabIndex        =   3
      Tag             =   $"frmBLChangeAcctMeth.frx":3FE7
      Top             =   2304
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
      _ExtentY        =   529
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSavePenGLs 
      Height          =   645
      Left            =   9150
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   $"frmBLChangeAcctMeth.frx":4083
      Top             =   2010
      Width           =   1650
      _Version        =   131072
      _ExtentX        =   2910
      _ExtentY        =   1138
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":410E
   End
   Begin EditLib.fpText fptxtPenX 
      Height          =   300
      Left            =   8772
      TabIndex        =   4
      Tag             =   $"frmBLChangeAcctMeth.frx":42F9
      Top             =   2304
      Width           =   348
      _Version        =   196608
      _ExtentX        =   614
      _ExtentY        =   529
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
      CharValidationText=   "X"
      MaxLength       =   1
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGLVerify 
      Height          =   540
      Left            =   870
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   $"frmBLChangeAcctMeth.frx":4404
      Top             =   7395
      Width           =   2355
      _Version        =   131072
      _ExtentX        =   4154
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
      ButtonDesigner  =   "frmBLChangeAcctMeth.frx":44B8
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   3312
      TabIndex        =   29
      Top             =   8160
      Width           =   2076
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   191.951
      X2              =   11520.03
      Y1              =   7062.092
      Y2              =   7062.092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   5663.542
      X2              =   5663.542
      Y1              =   7062.092
      Y2              =   8464.569
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Second: Save Penalty GL Numbers"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   720
      TabIndex        =   20
      Top             =   1872
      Width           =   3132
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5772
      Left            =   192
      Top             =   2928
      Width           =   11340
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8820
      TabIndex        =   26
      Top             =   2016
      Width           =   252
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GL Cash"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   7524
      TabIndex        =   24
      Top             =   2016
      Width           =   828
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GL Accts Rec"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   5796
      TabIndex        =   23
      Top             =   2016
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GL Revenue"
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   4260
      TabIndex        =   22
      Top             =   2016
      Width           =   1116
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Third: Save Category Code GL Numbers"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   192
      TabIndex        =   21
      Top             =   2928
      Width           =   4044
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   924
      Left            =   708
      Top             =   1872
      Width           =   10236
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "First: Select and save the new accounting method"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   720
      TabIndex        =   18
      Top             =   1104
      Width           =   4572
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   684
      Left            =   708
      Top             =   1104
      Width           =   10236
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   192
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Accounting Method"
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
      Left            =   3420
      TabIndex        =   8
      Top             =   384
      Width           =   4812
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   144
      Width           =   8652
   End
End
Attribute VB_Name = "frmBLChangeAcctMeth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdAllX_Click()
  Dim x As Integer
  
  For x = 1 To 500
    vaSpread1.Row = x
    vaSpread1.Col = 1
    If vaSpread1.Text <> "" Then
      vaSpread1.Col = 6
      vaSpread1.Text = "X"
    End If
  Next x

End Sub

Private Sub cmdClearAR_Click()
  Dim x As Integer
  
  For x = 1 To 500
    vaSpread1.Col = 4
    vaSpread1.Row = x
    vaSpread1.Text = ""
  Next x

End Sub

Private Sub cmdClearCash_Click()
  Dim x As Integer
  
  For x = 1 To 500
    vaSpread1.Col = 5
    vaSpread1.Row = x
    vaSpread1.Text = ""
  Next x

End Sub

Private Sub cmdClearRev_Click()
  Dim x As Integer
  
  For x = 1 To 500
    vaSpread1.Col = 3
    vaSpread1.Row = x
    vaSpread1.Text = ""
  Next x
  
End Sub

Private Sub cmdClearX_Click()
  Dim x As Integer
  
  For x = 1 To 500
    vaSpread1.Col = 6
    vaSpread1.Row = x
    vaSpread1.Text = ""
  Next x

End Sub

Private Sub cmdExit_Click()
  KillFile "changeaccmeth.dat"
  frmSoSoftMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdGLVerify_Click()
  If GLNumsValid = False Then Exit Sub
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F3 Turn Off Help"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdHelp.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F3 Turn On Help"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
    cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
  End If
End Sub

Private Sub cmdSaveGLNums_Click()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CHandle As Integer
  Dim CatCodeCnt As Integer
  Dim x As Integer
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim TotalAccts As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AcctMethod$
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  AcctMethod = QPTrim$(TownRec.AcctMeth)
  
  Select Case AcctMethod
    Case "A"
      For x = 1 To 500
        vaSpread1.Row = x
        vaSpread1.Col = 1
        If vaSpread1.Text = "" Then GoTo EmptyRowA
        vaSpread1.Col = 3
        If vaSpread1.Text = "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Accrual', currently saved, requires a GL Revenue number for all categories. Please include this number on row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 3, x
          Exit Sub
        End If
        vaSpread1.Col = 4
        If vaSpread1.Text = "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Accrual', currently saved, requires a GL Accounts Receivable number for all categories. Please include this number on row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 4, x
          Exit Sub
        End If
        vaSpread1.Col = 5
        If vaSpread1.Text = "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Accrual', currently saved, requires a GL Cash number for all categories. Please include this number on row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 5, x
          Exit Sub
        End If
EmptyRowA:
      Next x
    Case "C"
      For x = 1 To 500
        vaSpread1.Row = x
        vaSpread1.Col = 1
        If vaSpread1.Text = "" Then GoTo EmptyRowC
        vaSpread1.Col = 3
        If vaSpread1.Text = "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Cash', currently saved, requires a GL Revenue number for all categories. Please include this number on row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 3, x
          Exit Sub
        End If
        vaSpread1.Col = 4
        If vaSpread1.Text <> "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Cash', currently saved, does not allow a GL Accounts Receivable number for any category. Please delete this number from row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 4, x
          Exit Sub
        End If
        vaSpread1.Col = 5
        If vaSpread1.Text = "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Cash', currently saved, requires a GL Cash number for all categories. Please include this number on row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 5, x
          Exit Sub
        End If
EmptyRowC:
      Next x
    Case "N"
      For x = 1 To 500
        vaSpread1.Row = x
        vaSpread1.Col = 1
        If vaSpread1.Text = "" Then GoTo EmptyRowN
        vaSpread1.Col = 3
        If vaSpread1.Text <> "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'None', currently saved, does not allow a GL Revenue number for any category. Please delete this number from row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 3, x
          Exit Sub
        End If
        vaSpread1.Col = 4
        If vaSpread1.Text <> "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'None', currently saved, does not allow a GL Accounts Receivable number for any category. Please delete this number from row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 4, x
          Exit Sub
        End If
        vaSpread1.Col = 5
        If vaSpread1.Text <> "" Then
          frmBLMessageBoxJr.Label1.Caption = "The accounting method 'None', currently saved, does not a GL Cash number for any category. Please declude this number from row " + CStr(x) + "."
          frmBLMessageBoxJr.Label1.Top = 700
          frmBLMessageBoxJr.Show vbModal
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 5, x
          Exit Sub
        End If
EmptyRowN:
      Next x
    Case Else
      frmBLMessageBoxJr.Label1.Caption = "Error: The program cannot determine the accounting method now saved. Please correct this problem before continuing."
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
      
    End Select
  
  OpenCatCodeIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then 'file is there but there is nothing in it
    frmBLMessageBoxJr.Label1.Caption = "No Category Codes in index."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim CodeIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle
  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)
  
  For x = 1 To CatCodeCnt
    Get CHandle, CodeIdx(x), CodeRec
    vaSpread1.Row = x
    vaSpread1.Col = 3
    CodeRec.REVGLNUM = GetGLRecNum(vaSpread1.Text)
    vaSpread1.Col = 4
    CodeRec.ARGLACCT = GetGLRecNum(vaSpread1.Text)
    vaSpread1.Col = 5
    CodeRec.CASHACCT = GetGLRecNum(vaSpread1.Text)
    Put CHandle, CodeIdx(x), CodeRec
  Next x
  
  frmBLSucSave.Label1.Caption = "Your category code GL numbers have been saved successfully."
  frmBLSucSave.Show vbModal
  
End Sub

Private Sub cmdSaveAcctMeth_Click()
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  If Mid(fpcmbAcctMeth.Text, 1, 1) = QPTrim$(TownRec.AcctMeth) Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "The accounting method you are saving is the method that is currently saved. Press F10 to continue saving anyway, otherwise press ESC to abort."
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Save"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Close TownHandle
      Unload frmBLMessageBoxJrWOpts
      Exit Sub
    End If
  End If
  
  TownRec.AcctMeth = Mid(fpcmbAcctMeth.Text, 1, 1)
  Put TownHandle, 1, TownRec
  Close TownHandle
  
  Select Case QPTrim$(TownRec.AcctMeth)
    Case "N"
      cmdGL.Enabled = False
    Case Else
      cmdGL.Enabled = True
  End Select
  
  frmBLSucSave.Label1.Caption = "Accounting method saved successfully."
  frmBLSucSave.Show vbModal
 End Sub

Private Sub cmdSavePenGLs_Click()
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AcctMeth$
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  AcctMeth$ = QPTrim$(TownRec.AcctMeth)
  Select Case AcctMeth$
    Case "A"
      If QPTrim$(fptxtPenRevGL.Text) = "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Accrual', currently saved, requires a GL number for Penalty GL Revenue. Please add this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenRevGL.SetFocus
        Exit Sub
      End If
      If QPTrim$(fptxtPenARGL.Text) = "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Accrual', currently saved, requires a GL number for Penalty GL Accounts Receivable. Please add this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenARGL.SetFocus
        Exit Sub
      End If
      If QPTrim$(fptxtPenCashGL.Text) = "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Accrual', currently saved, requires a GL number for Penalty GL Cash. Please add this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenCashGL.SetFocus
        Exit Sub
      End If
    Case "C"
      If QPTrim$(fptxtPenRevGL.Text) = "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Cash', currently saved, requires a GL number for Penalty GL Revenue. Please add this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenRevGL.SetFocus
        Exit Sub
      End If
      If QPTrim$(fptxtPenARGL.Text) <> "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Cash', currently saved, does not allow a GL number for Penalty GL Accounts Receivable. Please delete this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenARGL.SetFocus
        Exit Sub
      End If
      If QPTrim$(fptxtPenCashGL.Text) = "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'Cash', currently saved, requires a GL number for Penalty GL Cash. Please add this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenCashGL.SetFocus
        Exit Sub
      End If
    Case "N"
      If QPTrim$(fptxtPenRevGL.Text) <> "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'None', currently saved, does not allow a GL number for Penalty GL Revenue. Please delete this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenRevGL.SetFocus
        Exit Sub
      End If
      If QPTrim$(fptxtPenARGL.Text) <> "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'None', currently saved, does not allow a GL number for Penalty GL Accounts Receivable. Please delete this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenARGL.SetFocus
        Exit Sub
      End If
      If QPTrim$(fptxtPenCashGL.Text) <> "" Then
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Label1.Caption = "The accounting method 'None', currently saved, does not allow a GL number for Penalty GL Cash. Please delete this number."
        frmBLMessageBoxJr.Show vbModal
        Close
        fptxtPenCashGL.SetFocus
        Exit Sub
      End If
    Case Else
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Label1.Caption = "ERROR: The program could not determine the current accounting method. Please correct this situation before continuing."
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
  End Select
      
  TownRec.PENREVGLNUM = GetGLRecNum(fptxtPenRevGL.Text)
  TownRec.PENRECGLNUM = GetGLRecNum(fptxtPenARGL.Text)
  TownRec.PENCASHACCT = GetGLRecNum(fptxtPenCashGL.Text)
  Put TownHandle, 1, TownRec
  Close TownHandle
  
  frmBLSucSave.Label1.Caption = "Your penalty GL numbers have been saved successfully."
  frmBLSucSave.Show vbModal
  
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%S"
      Call cmdSaveAcctMeth_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSaveGLNums_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%V"
      Call cmdGLVerify_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%G"
      Call cmdGL_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdGL_Click()
  frmBLGLList.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "changeaccmeth.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLChangeAcctMeth.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim CatCodeCnt As Integer
  Dim Nextx As Integer
  Dim One As Integer
  Dim DHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No Category Code Index has been saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  lblBalloon.Visible = False
  
  OpenCatCodeIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then 'file is there but there is nothing in it
    frmBLMessageBoxJr.Label1.Caption = "No Category Codes in index."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim CodeIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle
  
  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)
  
  If CatCodeCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No Category Codes on file."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
 
  For x = 1 To CatCodeCnt
    Get CHandle, CodeIdx(x), CodeRec
    vaSpread1.Row = x
    vaSpread1.Col = 1
    vaSpread1.Text = QPTrim$(CodeRec.CatCode)
    vaSpread1.Col = 2
    vaSpread1.Text = QPTrim$(CodeRec.CODEDESC)
    vaSpread1.Col = 3
    vaSpread1.Text = GetGLNum(CodeRec.REVGLNUM)
    vaSpread1.Col = 4
    vaSpread1.Text = GetGLNum(CodeRec.ARGLACCT)
    vaSpread1.Col = 5
    vaSpread1.Text = GetGLNum(CodeRec.CASHACCT)
  Next x
  
  Call FixSpread
  
  fpcmbAcctMeth.AddItem "None"
  fpcmbAcctMeth.AddItem "Cash"
  fpcmbAcctMeth.AddItem "Accrual"
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  fptxtPenRevGL.Text = GetGLNum(TownRec.PENREVGLNUM)
  fptxtPenARGL.Text = GetGLNum(TownRec.PENRECGLNUM)
  fptxtPenCashGL.Text = GetGLNum(TownRec.PENCASHACCT)
  
  Select Case QPTrim$(TownRec.AcctMeth)
    Case "A"
      fpcmbAcctMeth.Text = "Accrual"
    Case "N"
      fpcmbAcctMeth.Text = "None"
    Case "C"
      fpcmbAcctMeth.Text = "Cash"
  End Select
  
  One = 1
  DHandle = FreeFile
  Open "changeaccmeth.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
End Sub

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
'    GoTo SkipAdjust
    Select Case ScreenW
      Case 1280
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 5
          coladj = 10
          vaSpread1.FontSize = 18
          vaSpread1.RowHeight(-1) = 22
          vaSpread1.RowHeight(0) = 22
        Else
          COne = 13
          coladj = 4.5
          vaSpread1.RowHeight(-1) = 18
          vaSpread1.RowHeight(0) = 18
        End If
      Case 1152
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 14
          coladj = 7
          vaSpread1.FontSize = 14
          vaSpread1.RowHeight(0) = 18.5
          vaSpread1.RowHeight(-1) = 18.5
        Else
          COne = 6.65
          coladj = 2.25
          vaSpread1.RowHeight(0) = 16
          vaSpread1.RowHeight(-1) = 17
        End If
      Case 1024
        If Screen.TwipsPerPixelX <> 12 Then
          COne = 13.49
          coladj = 5.65
          vaSpread1.RowHeight(0) = 14
          vaSpread1.RowHeight(-1) = 14
        Else
          COne = 2.5
          coladj = 0.8
        End If
      Case 800
        COne = 0
        coladj = -0.5
        vaSpread1.Font.Size = 12
        vaSpread1.RowHeight(-1) = 14
      Case Else
    End Select
    
SkipAdjust:
    vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1)
    vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + COne
    vaSpread1.ColWidth(3) = vaSpread1.ColWidth(3) + coladj
    vaSpread1.ColWidth(4) = vaSpread1.ColWidth(4) + coladj
    vaSpread1.ColWidth(5) = vaSpread1.ColWidth(5) + coladj
'    vaSpread1.ColWidth(6) = vaSpread1.ColWidth(6)

End Sub

Private Sub fpcmbAcctMeth_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbAcctMeth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAcctMeth.ListIndex = -1
  End If
  If fpcmbAcctMeth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtPenRevGL.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtPenCashGL_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbAcctMeth.SetFocus
  End If
End Sub

Private Sub fptxtPenX_Click(Button As Integer)
  If QPTrim$(fptxtPenX.Text) = "" Then
    fptxtPenX.Text = "X"
  ElseIf QPTrim$(fptxtPenX.Text) = "X" Then
    fptxtPenX.Text = ""
  End If
    
End Sub

Private Sub fptxtPenX_KeyDown(KeyCode As Integer, Shift As Integer)
  If QPTrim$(fptxtPenX.Text) = "" Then
    fptxtPenX.Text = "X"
  ElseIf QPTrim$(fptxtPenX.Text) = "X" Then
    fptxtPenX.Text = ""
  End If
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
  If Col = 6 Then
    vaSpread1.Col = 1
    vaSpread1.Row = Row
    If vaSpread1.Text <> "" Then
      vaSpread1.Col = Col
      vaSpread1.Row = Row
      If vaSpread1.Text = "X" Then
        vaSpread1.Text = ""
      ElseIf vaSpread1.Text = "" Then
        vaSpread1.Text = "X"
      End If
    End If
  End If
     
End Sub

Private Function GLNumsValid() As Boolean
  Dim GLIdxRec As JGLAcctIdxType
  Dim IdxHandle As Integer
  Dim NumOfGLRecs As Integer
  Dim x As Integer, y As Integer
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  Dim RevNum$, Rev As Integer
  Dim AcctsRecNum$, Acct As Integer
  Dim CashRecNum$, Cash As Integer
  Dim ThisGLNum$, Nextx As Integer
  Dim EmptyRow As Integer
  Dim NumOfIdxRecs As Integer
  
  On Error GoTo ERRORSTUFF
  
  GLNumsValid = True
  If Not Exist("GLACCT.IDX") Or Not Exist("GLACCT.DAT") Then
    Exit Function
  End If
  
  Rev = 0
  Acct = 0
  Cash = 0
  OpenGLIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(GLIdxRec)
  If NumOfIdxRecs = 0 Then
    MsgBox "ERROR: No GL index can be found. Verification aborted."
    Close
    Exit Function
  End If
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer
  
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, GLIdxRec 'build GL number index
    IdxRec(x) = GLIdxRec.RecNo
  Next x
  Close IdxHandle
  
  OpenGLAcctFile AcctHandle
'  NumOfGLRecs = LOF(AcctHandle) / Len(GLAcctRec)
  For x = 1 To NumOfIdxRecs
    Get AcctHandle, IdxRec(x), GLAcctRec
      If Rev = 0 Then
        If QPTrim$(fptxtPenRevGL.Text) = "" Then Rev = 1
        If QPTrim$(fptxtPenRevGL.Text) = QPTrim$(GLAcctRec.Num) Then
          Rev = 1
          If Cash = 1 And Acct = 1 Then Exit For
        End If
      End If
      If Acct = 0 Then
        If QPTrim$(fptxtPenARGL.Text) = "" Then Acct = 1
        If QPTrim$(fptxtPenARGL.Text) = QPTrim$(GLAcctRec.Num) Then
          Acct = 1
          If Rev = 1 And Cash = 1 Then Exit For
        End If
      End If
      If Cash = 0 Then
        If QPTrim$(fptxtPenCashGL.Text) = "" Then Cash = 1
        If QPTrim$(fptxtPenCashGL.Text) = QPTrim$(GLAcctRec.Num) Then
          Cash = 1
          If Rev = 1 And Acct = 1 Then Exit For
        End If
      End If
  Next x
  
  If Rev = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "The Penalty GL Revenue number is not a valid GL number. Please correct that number."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtPenRevGL.SetFocus
    GLNumsValid = False
    Exit Function
  End If
  If Acct = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "The Penalty GL Accounts Receivable number is not a valid GL number. Please correct that number."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtPenARGL.SetFocus
    GLNumsValid = False
    Exit Function
  End If
  If Cash = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "The Penalty GL Cash number is not a valid GL number. Please correct that number."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtPenCashGL.SetFocus
    GLNumsValid = False
    Exit Function
  End If
  
  Rev = 0
  Acct = 0
  Cash = 0
  Nextx = 1
  
  Do
    vaSpread1.Col = 1
    vaSpread1.Row = Nextx
    
    If vaSpread1.Text = "" Then
      EmptyRow = EmptyRow + 1
      GoTo NewLoop
    End If
    
    For x = 1 To NumOfIdxRecs
      vaSpread1.Row = Nextx
      Get AcctHandle, IdxRec(x), GLAcctRec
      If GLAcctRec.Deleted Then GoTo NotThisOne
        If Rev = 0 Then
          vaSpread1.Col = 3
          If QPTrim$(vaSpread1.Text) = "" Then Rev = 1
          If QPTrim$(vaSpread1.Text) = QPTrim$(GLAcctRec.Num) Then
            Rev = 1
            If Acct = 1 And Cash = 1 Then
              Nextx = Nextx + 1
              GoTo NewLoop
            End If
          End If
        End If
        If Acct = 0 Then
          vaSpread1.Col = 4
          If QPTrim$(vaSpread1.Text) = "" Then Acct = 1
          If QPTrim$(vaSpread1.Text) = QPTrim$(GLAcctRec.Num) Then
            Acct = 1
            If Rev = 1 And Cash = 1 Then
              Nextx = Nextx + 1
              GoTo NewLoop
            End If
          End If
        End If
        If Cash = 0 Then
          vaSpread1.Col = 5
          If QPTrim$(vaSpread1.Text) = "" Then Cash = 1
          If QPTrim$(vaSpread1.Text) = QPTrim$(GLAcctRec.Num) Then
            Cash = 1
            If Rev = 1 And Acct = 1 Then
              Nextx = Nextx + 1
              GoTo NewLoop
            End If
          End If
        End If
NotThisOne:
    Next x
              
    If Rev = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "The GL Revenue number on row " + CStr(Nextx) + " is not valid. Please correct this number."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      GLNumsValid = False
      vaSpread1.SetFocus
      vaSpread1.SetActiveCell 3, Nextx
      Exit Function
    End If
    If Acct = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "The GL Accounts Receivable number on row " + CStr(Nextx) + " is not valid. Please correct this number."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      GLNumsValid = False
      vaSpread1.SetFocus
      vaSpread1.SetActiveCell 4, Nextx
      Exit Function
    End If
    If Cash = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "The GL Cash number on row " + CStr(Nextx) + " is not valid. Please correct this number."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Close
      GLNumsValid = False
      vaSpread1.SetFocus
      vaSpread1.SetActiveCell 5, Nextx
      Exit Function
    End If
               
    Nextx = Nextx + 1
NewLoop:
    Rev = 0
    Acct = 0
    Cash = 0
    If EmptyRow = 25 Then Exit Do
    
  Loop Until Nextx = 500
    
  frmBLMessageBoxJr.Label1.Caption = "All General Ledger numbers are valid."
  frmBLMessageBoxJr.Label1.Top = 700
  frmBLMessageBoxJr.Show vbModal
  
  
  Exit Function
  
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatEdit", "GLNumsValid", Erl)
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

