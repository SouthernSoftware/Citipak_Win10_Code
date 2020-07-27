VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCatCodeLookup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Code Lookup"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCatCodeLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbType 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Tag             =   $"frmBLCatCodeLookup.frx":08CA
      Top             =   2640
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
      ColDesigner     =   "frmBLCatCodeLookup.frx":0A89
   End
   Begin LpLib.fpList fpList1 
      Height          =   3540
      Left            =   2295
      TabIndex        =   3
      Tag             =   $"frmBLCatCodeLookup.frx":0E37
      Top             =   4830
      Width           =   7065
      _Version        =   196608
      _ExtentX        =   12462
      _ExtentY        =   6244
      TextAlias       =   ""
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
      Columns         =   2
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
      ColDesigner     =   "frmBLCatCodeLookup.frx":0F9F
   End
   Begin EditLib.fpText fptxtCatNum 
      Height          =   390
      Left            =   4920
      TabIndex        =   0
      Tag             =   $"frmBLCatCodeLookup.frx":12D8
      Top             =   1515
      Width           =   1305
      _Version        =   196608
      _ExtentX        =   2302
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
      MaxLength       =   5
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
   Begin EditLib.fpText fptxtCatDesc 
      Height          =   390
      Left            =   4920
      TabIndex        =   1
      Tag             =   $"frmBLCatCodeLookup.frx":14AF
      Top             =   2077
      Width           =   4230
      _Version        =   196608
      _ExtentX        =   7461
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   4875
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "Press the 'Exit' button to leave this screen and return to the main Category Maintenance menu."
      Top             =   3450
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmBLCatCodeLookup.frx":16BD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSearch 
      Height          =   540
      Left            =   7485
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   $"frmBLCatCodeLookup.frx":189B
      Top             =   3450
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
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
      ButtonDesigner  =   "frmBLCatCodeLookup.frx":1A8A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   $"frmBLCatCodeLookup.frx":1C68
      Top             =   3450
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
      ButtonDesigner  =   "frmBLCatCodeLookup.frx":1CF9
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   0
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3345
      Left            =   720
      Top             =   1200
      Width           =   10200
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1500
      Top             =   180
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Category LookUp"
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
      Left            =   2940
      TabIndex        =   11
      Top             =   315
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category Type:"
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
      Left            =   3015
      TabIndex        =   10
      Top             =   2730
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category Description:"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   2355
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category Number:"
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
      Left            =   2820
      TabIndex        =   8
      Top             =   1600
      Width           =   1980
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
      Left            =   2145
      TabIndex        =   7
      Top             =   4035
      Width           =   2100
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmBLCatCodeLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmBLCategoryMaintMenu.Show
  DoEvents
  Unload frmBLCatCodeLookup
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

Private Sub cmdSearch_Click()
   Dim CodeRec As ARNewCatCodeRecType
   Dim CodeHandle As Integer
   Dim CodeIdxRec As CatCodeIdxType
   Dim IdxHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim cnt As Integer
   Dim CatCnt As Integer
   Dim TempCatDesc$
   Dim TempCatNum$
   Dim TempCatType$
   Dim NumFlag As Boolean
   Dim DescFlag As Boolean
   Dim TypeFlag As Boolean
   Dim Found As Boolean
   Dim MatchCnt As Integer
   Dim FoundCnt As Integer
   Dim OnlyOneFound$
   
   On Error GoTo ERRORSTUFF
   fpList1.Clear
   
   NumFlag = False
   DescFlag = False
   TypeFlag = False
   
   If Not Exist("arcode.dat") Then 'no file there
     frmBLMessageBoxJr.Label1.Caption = "No Category files have been saved."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Close
     Exit Sub
   End If
   
   OpenCatCodeIdxFile IdxHandle
   CatCnt = LOF(IdxHandle) / Len(CodeIdxRec)
   ReDim IdxRec(1 To CatCnt) As Integer
   For x = 1 To CatCnt
     Get IdxHandle, x, CodeIdxRec
       IdxRec(x) = CodeIdxRec.CatCodeRec
   Next x
   Close IdxHandle
   
   If QPTrim$(fptxtCatNum.Text) <> "" Then
     TempCatNum$ = QPTrim$(fptxtCatNum.Text)
     NumFlag = True
   End If
   
   If QPTrim$(fptxtCatDesc.Text) <> "" Then
     TempCatDesc = QPTrim$(fptxtCatDesc.Text)
     DescFlag = True
   End If
   
   If QPTrim$(fpcmbType.Text) <> "" Then
     TempCatType = Mid(fpcmbType.Text, 1, 1)
     TypeFlag = True
   End If
   
   OpenCatCodeFile CodeHandle
   
   For x = 1 To CatCnt
     Get CodeHandle, IdxRec(x), CodeRec
     Found = True
     If NumFlag = True Then
       If InStr(CodeRec.CatCode, TempCatNum) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If DescFlag = True Then
       If InStr(UCase$(CodeRec.CODEDESC), TempCatDesc) > 0 Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
     If TypeFlag = True Then
       If Mid(CodeRec.CodeType, 1, 1) = Mid(TempCatType, 1, 1) Then
         Found = True
       Else
         Found = False
         GoTo NotAMatch
       End If
     End If
   
    If Found Then
      FoundCnt = FoundCnt + 1
      fpList1.Row = -1
      MatchCnt = MatchCnt + 1
      GCatNum = IdxRec(x)
      fpList1.InsertRow = QPTrim$(CodeRec.CatCode) & Chr$(9) & " " & QPTrim$(CodeRec.CODEDESC)
      'only used if no more than one found
      OnlyOneFound = QPTrim$(CodeRec.CatCode)
    End If
NotAMatch:
  Next x
  
  If FoundCnt > 1 Then
    fpList1.SetFocus
  End If
   
  If MatchCnt <= 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No match found."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
    Close
  End If
  
  If FoundCnt > 1 Then
    fpList1.ListIndex = 0
  End If
  
  If FoundCnt = 1 Then
    For x = 1 To CatCnt
      Get CodeHandle, IdxRec(x), CodeRec
      If OnlyOneFound = QPTrim$(CodeRec.CatCode) Then
        GCatNum = IdxRec(x)
        Exit For
      Else
        Found = False
        GoTo NotThisTime
      End If
   
NotThisTime:
    Next x
    
    fptxtCatNum.Text = ""
    fptxtCatDesc.Text = ""
    fpcmbType.Text = ""
    
    fpList1.Clear
    FoundCnt = 0
    
    frmBLCustEdit.Caption = "Business License Edit Category"
    frmBLCatEdit.Show
    DoEvents
'    Unload frmBLCatCodeLookup
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  lblBalloon.Visible = False
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
    If QPTrim$(fptxtCatNum.Text) <> "" Or QPTrim$(fptxtCatDesc.Text) <> "" _
      Or QPTrim$(fpcmbType.Text) <> "" Then
      Call cmdSearch_Click
      KeyCode = 0
      Exit Sub
    ElseIf fpList1.ListIndex <> -1 Then
      GoTo CatAlreadySelected '8/6
    Else
      SendKeys "{Tab}"
      KeyCode = 0
      Exit Sub
    End If
CatAlreadySelected:
    fpList1.Col = 1
    If QPTrim$(fpList1.ColText) = "" Then
      frmBLMessageBoxJr.Label1.Caption = "No Category Code has been selected"
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
    Else
      Call fpList1_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
    
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdSearch_Click
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%T"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%C"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCatCodeLookup.")
      Call Terminate
      End
    End If
  End If
End Sub
Public Sub LoadMe()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CodeRecNo As Integer
  Dim CatCodeCnt As Integer
  Dim ThisCodeNum As Integer
  Dim CatCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No Category Code Index has been saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
   
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
  
  If Not Exist("ARCODE.DAT") Then
    frmBLMessageBoxJr.Label1.Caption = "Path to ARCODE.DAT could not be found."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)
  
  If CatCodeCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No Category Codes on file."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To CodeIdxRecNum 'CatCodeCnt
    Get CHandle, CodeIdx(x), CodeRec
    If Len(QPTrim(CodeRec.CatCode)) = 0 Then GoTo BadCode
    fpList1.InsertRow = QPTrim$(CodeRec.CatCode) & " " & Chr$(9) & QPTrim$(CodeRec.CODEDESC)
    CatCnt = CatCnt + 1
    If GCatNum = CodeIdx(x) Then ThisCodeNum = x 'used to return to
    'the list after a save
BadCode:
  Next x
  Close CHandle
  
  If CatCnt > 1 Then
    fpList1.ListIndex = 0
  End If
  'the following code was put here so that the row selected
  'always shows up on the screen...
  If GCatNum > 0 Then '0 always goes to top
    fpList1.Selected(ThisCodeNum - 1) = True 'go back to where we left it after a save
     If CatCodeCnt - 13 < ThisCodeNum - 1 Then '13 is the most this list can show at once
       fpList1.TopIndex = CatCodeCnt - 13 'this is the index that allows
       'the last item to still be seen
     Else
       fpList1.TopIndex = ThisCodeNum - 1
     End If
  End If
  
ZeroText:
  Exit Sub
   

ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatCodeLookup", "LoadMe", Erl)
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


Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtCatNum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpList1_DblClick()
   Dim CodeRec As ARNewCatCodeRecType
   Dim CHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim DESC$
   Dim Code$
   Dim Found As Boolean
   
   fpList1.Col = 0 'assign variables from the user selected row
   If QPTrim$(fpList1.ColText) = "" Then
     frmBLMessageBoxJr.Label1.Caption = "No item has been selected"
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Exit Sub
   End If
   Code$ = QPTrim$(fpList1.ColText)
   fpList1.Col = 1
   DESC$ = QPTrim$(fpList1.ColText)
   
   OpenCatCodeFile CHandle
   TotalAccts = LOF(CHandle) \ Len(CodeRec)
   
   If TotalAccts = 0 Then Exit Sub
   
   For x = 1 To TotalAccts
     Get CHandle, x, CodeRec
     If Code$ = QPTrim$(CodeRec.CatCode) And DESC$ = QPTrim$(CodeRec.CODEDESC) Then 'match the selected
     'row with the right code
       Found = True
       fpList1.Row = -1
       GCatNum = x 'now you can assign the correct global
       Exit For
     Else
       Found = False
       GoTo NotAMatch
     End If
      
NotAMatch:
   Next x
  Close CHandle
  
  If Found = True Then
    frmBLCatEdit.Show
    DoEvents
'    Me.Hide
  Else
    frmBLMessageBoxJr.Label1.Caption = "No match found."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
End Sub
  
Public Sub RefreshSearchList()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim CHandle As Integer
  Dim x As Integer
  Dim NumOfCodes As Integer
  Dim ThisRow As Integer
  Dim CatCodeCnt As Integer
  Dim ThisCodeNum As Integer
  
  On Error GoTo ERRORSTUFF
  'this sub is only activated from the category edit screen and
  'is designed to rebuild the category list with the latest
  'saved data and it places the focus on the line corresponding
  'to whichever category was last on the edit screen
  fpList1.Action = ActionClear
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
  For x = 1 To CodeIdxRecNum
    Get CHandle, CodeIdx(x), CodeRec
    If Len(QPTrim(CodeRec.CatCode)) = 0 Then GoTo BadCode
    fpList1.InsertRow = QPTrim$(CodeRec.CatCode) & " " & Chr$(9) & QPTrim$(CodeRec.CODEDESC)
    If GCatNum = CodeIdx(x) Then ThisCodeNum = x 'used to return to
    'the list after a save
BadCode:
  Next x
  Close CHandle
  If GCatNum > 0 Then '0 always goes to top
    fpList1.Selected(ThisCodeNum - 1) = True 'go back to where we left it after a save
     If CatCodeCnt - 13 < ThisCodeNum - 1 Then '13 is the most this list can show at once
       fpList1.TopIndex = CatCodeCnt - 13 'this is the index that allows
       'the last item to still be seen
     Else
       fpList1.TopIndex = ThisCodeNum - 1
     End If
  End If
  
ZeroText:
  Exit Sub
   

ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCatCodeLookup", "LoadMe", Erl)
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
  
  

