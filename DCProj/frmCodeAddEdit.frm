VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCodeAddEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decal Code Add/Edit"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodeAddEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboInactive 
      Height          =   384
      Left            =   4506
      TabIndex        =   7
      Top             =   5208
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   677
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
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCodeAddEdit.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcctNumNa2 
      Height          =   384
      Left            =   4512
      TabIndex        =   5
      Top             =   4236
      Width           =   5856
      _Version        =   196608
      _ExtentX        =   10329
      _ExtentY        =   677
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
      Columns         =   4
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
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
      ScrollBarH      =   3
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCodeAddEdit.frx":0C30
   End
   Begin LpLib.fpCombo fpcboAcctNumNa1 
      Height          =   384
      Left            =   4512
      TabIndex        =   2
      Top             =   3744
      Width           =   5856
      _Version        =   196608
      _ExtentX        =   10329
      _ExtentY        =   677
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
      Columns         =   4
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
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
      ScrollBarH      =   3
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCodeAddEdit.frx":1033
   End
   Begin EditLib.fpLongInteger fpCodeRecNo 
      Height          =   300
      Left            =   768
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   144
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   96
      Top             =   144
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   8328
      TabIndex        =   8
      Top             =   6864
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
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
      ButtonDesigner  =   "frmCodeAddEdit.frx":1436
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   384
      Left            =   9936
      TabIndex        =   9
      Top             =   6864
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
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
      ButtonDesigner  =   "frmCodeAddEdit.frx":1612
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPrintInfo 
      Height          =   384
      Left            =   6696
      TabIndex        =   10
      Top             =   6864
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
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
      ButtonDesigner  =   "frmCodeAddEdit.frx":17EE
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   11
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   593
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "1:48 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/1/2006"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpText fpCodeDesc 
      CausesValidation=   0   'False
      Height          =   372
      Left            =   4506
      TabIndex        =   1
      Top             =   3252
      Width           =   4884
      _Version        =   196608
      _ExtentX        =   8615
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AlignTextV      =   1
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
   Begin EditLib.fpText fpGLAcct1 
      Height          =   372
      Left            =   4506
      TabIndex        =   3
      Top             =   3744
      Width           =   3312
      _Version        =   196608
      _ExtentX        =   5842
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   14
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
   Begin EditLib.fpText fpGLAcct2 
      Height          =   372
      Left            =   4506
      TabIndex        =   4
      Top             =   4236
      Width           =   3312
      _Version        =   196608
      _ExtentX        =   5842
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   14
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
   Begin EditLib.fpDoubleSingle fpStandFee 
      Height          =   348
      Left            =   4506
      TabIndex        =   6
      Top             =   4740
      Width           =   1380
      _Version        =   196608
      _ExtentX        =   2434
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpText fpCode 
      CausesValidation=   0   'False
      Height          =   372
      Left            =   4512
      TabIndex        =   0
      Top             =   2784
      Width           =   852
      _Version        =   196608
      _ExtentX        =   1503
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   1
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483631
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
      MaxLength       =   3
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Edit Decal Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3330
      TabIndex        =   19
      Top             =   1008
      Width           =   5652
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      Height          =   612
      Left            =   3222
      Top             =   888
      Width           =   5772
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   3516
      Left            =   1554
      Top             =   2376
      Width           =   9180
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Category Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   1458
      TabIndex        =   18
      Top             =   2880
      Width           =   2856
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inactive (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2094
      TabIndex        =   17
      Top             =   5256
      Width           =   2220
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Fee Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2094
      TabIndex        =   16
      Top             =   4780
      Width           =   2220
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1734
      TabIndex        =   15
      Top             =   3360
      Width           =   2580
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue G/L Acct#:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1638
      TabIndex        =   14
      Top             =   3830
      Width           =   2676
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipt G/L Acct#:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1494
      TabIndex        =   13
      Top             =   4305
      Width           =   2820
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   732
      Left            =   3222
      Top             =   768
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      Height          =   3564
      Left            =   1530
      Top             =   2352
      Width           =   9228
   End
End
Attribute VB_Name = "frmCodeAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim CodeRecNo As Long, CntL As Long
Dim TransRec As Long, MsgRec As Long
Dim UBSetupLen As Integer, cnt As Integer
Dim OldBook As String, NBook As String
Dim FinalFlag As Boolean, UpDateOwner As Boolean
Dim BeenDone As Boolean, LinkGL As Boolean
Dim BtnFnt As Double, EditFlag As Boolean
Public Endedit As Boolean
Dim fromform As Form, toform As Form, codeopt As Integer
Dim dontdoit As Boolean, TempAcct2 As String
Dim TempAcct1 As String, DupFound As Boolean, ValidCode As Boolean
Dim Ms0 As String, Ms1 As String, Ms2 As String
Dim Ms3 As String, Ms4 As String, Ms5 As String
Dim ChkOKFlag As Boolean, stuff As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
    EditFlag = True
    Endedit = False
  Else
    codeopt = 0
    EditFlag = False
    Endedit = True
  End If
End Sub
'Public Sub SetScreen()
'  If EditFlag Then
'    Load frmCodeDisplayList
'    frmCodeDisplayList.Show vbModal
'  Else
'    NewCodeDefaults
'  End If
'End Sub

Private Sub Form_Load()
  Dim dcSetUpRec(1) As DCSetupType
  Dim RecLen As Integer
  'BlockInput True
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  LoadDCSetUpFile dcSetUpRec(), RecLen
  If QPTrim$(dcSetUpRec(1).GLInterface) = "Y" Then
    LinkGL = True
  Else
    LinkGL = False
  End If
  If LinkGL = True Then
    fpcboAcctNumNa1.Visible = True
    fpcboAcctNumNa2.Visible = True
    fpGLAcct1.Visible = False
    fpGLAcct1.Enabled = False
    fpGLAcct2.Visible = False
    fpGLAcct2.Enabled = False
    FillAcctNumName fpcboAcctNumNa1
    FillAcctNumName fpcboAcctNumNa2
  Else
    fpcboAcctNumNa1.Visible = False
    fpcboAcctNumNa1.Enabled = False
    fpGLAcct1.Visible = True
    fpcboAcctNumNa2.Visible = False
    fpcboAcctNumNa2.Enabled = False
    fpGLAcct2.Visible = True
  End If
  fpcboInactive.AddItem "No"
  fpcboInactive.AddItem "Yes"
  fpcboInactive.ListIndex = -1
  If EditFlag Then
    Label1.Caption = "Edit Decal Code"
    frmCodeAddEdit.Caption = "Edit Decal Code"
    Me.HelpContextID = hlpEditExistingDecal
  Else
    Label1.Caption = "Add Decal Code"
    frmCodeAddEdit.Caption = "Add Decal Code"
    Me.HelpContextID = hlpAddADecalCode
  End If
  Refresh
  DoEvents
  dontdoit = False
 ' BlockInput False
  'SetScreen
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via CodeAddEdit by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub ClearScn()
  fpCode = ""
  fpCodeDesc = ""
  fpStandFee = 0
  If LinkGL = True Then
    fpcboAcctNumNa1.ListIndex = -1
    fpcboAcctNumNa2.ListIndex = -1
  Else
    fpGLAcct1 = ""
    fpGLAcct2 = ""
  End If
  'fpcboInactive.ListIndex = -1
  CodeRecNo = 0
  'EditFlag = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    'DoEvents
    Temp_Class.ResizeControls Me
   ' DoEvents
   ' Me.Visible = True
   ' Me.AutoRedraw = False
   ' DoEvents
  End If
  DoEvents
End Sub


Private Sub fpCmdExit_Click()
    Select Case CheckSaveCodeFile
    Case True:  '-1 save chenges
      If SaveCode = True Then
        Call ExitCodes
      End If
    Case False:
      Call ExitCodes
    Case Else     '1 is review
      'continue editing
    End Select
End Sub
Private Function CheckSaveCodeFile%()
  Dim CodeFile As Integer, CodeRecNo As Integer
  Dim Changed As Boolean
  Dim DCCodeRecLen As Integer, cnt As Integer
  Dim DCCodeRec As DCCatCodeRecType
  Dim TText As String
  
  Changed = False
  DCCodeRecLen = Len(DCCodeRec)
  
  If fpCodeRecNo > 0 Then
    CodeFile = FreeFile
    Open DCPath + "DCCode.DAT" For Random Shared As CodeFile Len = DCCodeRecLen
    Get CodeFile, fpCodeRecNo, DCCodeRec
    Close CodeFile
    If QPTrim$(DCCodeRec.CODEDESC) <> QPTrim$(fpCodeDesc.Text) Then
      Changed = True
      GoTo ExitCheck
    End If
    If DCCodeRec.Fee <> fpStandFee.Value Then
      Changed = True
      GoTo ExitCheck
    End If
    If LinkGL Then
      If fpcboAcctNumNa1.ListIndex > 0 Then
        fpcboAcctNumNa1.col = 1
        If DCCodeRec.REVGLNUM <> QPTrim$(fpcboAcctNumNa1.ColText) Then
          Changed = True
          GoTo ExitCheck
        End If
      End If
      If fpcboAcctNumNa2.ListIndex > 0 Then
        fpcboAcctNumNa2.col = 1
        If DCCodeRec.CASHACCT <> QPTrim$(fpcboAcctNumNa2.ColText) Then
          Changed = True
          GoTo ExitCheck
        End If
      End If
    Else
      If QPTrim$(DCCodeRec.REVGLNUM) <> QPTrim$(fpGLAcct1) Then
        Changed = True
        GoTo ExitCheck
      End If
      If QPTrim$(DCCodeRec.CASHACCT) <> QPTrim$(fpGLAcct2) Then
        Changed = True
        GoTo ExitCheck
      End If
    End If
    If fpcboInactive.ListIndex = 1 Then
      If DCCodeRec.InactiveFlag = "N" Then
        Changed = True
        GoTo ExitCheck
      End If
    ElseIf fpcboInactive.ListIndex = 0 Then
      If DCCodeRec.InactiveFlag = "Y" Then
        Changed = True
        GoTo ExitCheck
      End If
    End If
    If QPTrim(fpCode) <> QPTrim(DCCodeRec.CATCODE) Then
      Changed = True
      GoTo ExitCheck
    End If
  End If

ExitCheck:
  If Changed Then
    Load frmChangedWarning
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      CheckSaveCodeFile = False
    Case True
      CheckSaveCodeFile = True
    Case 1
      CheckSaveCodeFile = 1
    End Select
  Else
    CheckSaveCodeFile = False
  End If
End Function

Private Sub ExitCodes()
  On Local Error Resume Next
  BeenDone = False
'  If WhatCode > 0 Then
'    EditFlag = True
'    WhatCode = 0
'  End If
'  If fpCodeRecNo = 0 And WhatCode = 0 Then
'    If EditFlag Then
'      EditFlag = False
'    End If
'  End If
'  ClearScn
  If EditFlag And Not Endedit Then
'    EditFlag = False
'    fpCodeRecNo = -1
'    Call LoadCode2form
'   ' Call Form_Activate
      Load frmCodeDisplayList
      frmCodeDisplayList.Show vbModal
  Else
    Load frmDCCodeMenu
    DoEvents
    frmDCCodeMenu.Show
    Unload Me
  End If
End Sub
Private Sub Form_Activate()
  If EditFlag Then
    If Not Endedit Then
      If Val(fpCodeRecNo) > 0 Then
        EditFlag = True
        Call LoadCode2form
         DoEvents
      Else
        Load frmCodeDisplayList
        frmCodeDisplayList.Show vbModal
      End If
    Else
      ExitCodes
    End If
  Else
    NewCodeDefaults
  End If

'    If Val(fpCodeRecNo) = -1 And Not BeenDone Then
'      BeenDone = True
'      Load frmCodeDisplayList
'      frmCodeDisplayList.Show vbModal
'      DoEvents
'      If Val(fpCodeRecNo) > 0 Then
'        EditFlag = True
'        Call LoadCode2form
'        DoEvents
'      Else
'        Call ExitCodes
'        'Call fpCmdExit_Click
'      End If
'      DoEvents
'    Else
'      BeenDone = True
'    End If
End Sub
Private Sub NewCodeDefaults()
  fpCodeRecNo = 0
  fpCode = ""
  fpCode.Enabled = True
  fpCodeDesc = ""
    If LinkGL Then
      fpcboAcctNumNa1.ListIndex = 0
      fpcboAcctNumNa2.ListIndex = 0
    Else
      fpGLAcct1 = ""
      fpGLAcct2 = ""
    End If
  fpStandFee = 0
  fpcboInactive.ListIndex = 0
' ' LblInfo.Caption = "New"
' UBLog PWUser + " New Cust Entry"
End Sub
Private Sub LoadCode2form()
  ReDim DCCatCodeRec(1) As DCCatCodeRecType
  Dim CatCodeRecLen As Integer, CatFile As Integer, NumOfCatRecs As Long
  CatCodeRecLen = Len(DCCatCodeRec(1))
  CatFile = FreeFile
  Open "DCCODE.DAT" For Random As CatFile Len = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
  If fpCodeRecNo > 0 Then
    Get CatFile, fpCodeRecNo, DCCatCodeRec(1)
    fpCode = QPTrim(DCCatCodeRec(1).CATCODE)
   ' WhatCode = fpCodeRecNo
    fpCode.Enabled = False
    fpCodeDesc = QPTrim$(DCCatCodeRec(1).CODEDESC)
    fpStandFee.Value = DCCatCodeRec(1).Fee
    If LinkGL Then
      fpcboAcctNumNa1.SearchText = QPStrip(DCCatCodeRec(1).REVGLNUM)
      fpcboAcctNumNa2.SearchText = QPStrip(DCCatCodeRec(1).CASHACCT)
      fpcboAcctNumNa1.Action = 0
      fpcboAcctNumNa2.Action = 0
      If fpcboAcctNumNa1.SearchIndex <> -1 Then
        fpcboAcctNumNa1.ListIndex = fpcboAcctNumNa1.SearchIndex
      Else
        TempAcct1$ = QPTrim(DCCatCodeRec(1).REVGLNUM)
        fpcboAcctNumNa1.ListIndex = 0
      End If
      If fpcboAcctNumNa2.SearchIndex <> -1 Then
        fpcboAcctNumNa2.ListIndex = fpcboAcctNumNa2.SearchIndex
      Else
        TempAcct2$ = QPTrim(DCCatCodeRec(1).CASHACCT)
        fpcboAcctNumNa2.ListIndex = 0
      End If
    Else
      fpGLAcct1 = QPTrim$(DCCatCodeRec(1).REVGLNUM)
      fpGLAcct2 = QPTrim$(DCCatCodeRec(1).CASHACCT)
    End If
    If DCCatCodeRec(1).InactiveFlag = "Y" Then
      fpcboInactive.ListIndex = 1
    Else
      fpcboInactive.ListIndex = 0
    End If

  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF2
      KeyCode = 0
      Call fpcmdPrintinfo_Click
    Case vbKeyF10
      KeyCode = 0
      DoEvents
      fpCmdSave_Click
'032003
'    Case vbKeyReturn
'      KeyCode = 0
'      SendKeys "{tab}", True   ' Set the focus to the next control.
'      DoEvents
    Case Else:
  End Select
End Sub


Private Sub fpCmdSave_Click()
Dim cnt As Integer
  If SaveCode = True Then
    ClearScn
    If EditFlag Then
      Load frmCodeDisplayList
      frmCodeDisplayList.Show vbModal
    End If
  Else
  'stay right where you are
    If EditFlag = True Then
      fpCodeDesc.SetFocus
    Else
      fpCode.SetFocus
    End If
  End If
  DoEvents

End Sub

Private Sub fpCode_LostFocus()
  
  If fpCode.Enabled = True Then
  If EditFlag = False Then
    CheckCodes
    If DupFound Then
      Ms0 = "ERROR:"
      Ms1 = ""
      Ms2 = "DUPLICATE CODE!"
      Ms3 = ""
      Ms4 = ""
      Ms5 = "Please enter a new code."
      Dothemsg
      ChkOKFlag = False
      fpCode = ""
      fpCode.SetFocus
      Exit Sub
'    Else
'      fpCodeDesc.SetFocus
    End If
  End If
  End If
End Sub
Private Function ChkOK2Save()
    DupFound = False
    ChkOKFlag = True
    If EditFlag = False Then
      CheckCodes
      If DupFound = True Then
        Ms0 = "DUPLICATE CODE"
        Ms1 = ""
        Ms2 = "PLEASE CHECK YOUR CODE"
        Ms3 = ""
        Ms4 = "Enter a Different Code"
        Ms5 = ""
        Dothemsg
        ChkOKFlag = False
        Exit Function
      End If
    End If
    If Len(fpCode.Text) = 0 Then
      Ms0 = "INVALID CODE"
      Ms1 = ""
      Ms2 = "PLEASE CHECK YOUR CODE"
      Ms3 = ""
      Ms4 = "Enter a Valid Code"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Function
    End If
    If Len(fpCodeDesc.Text) = 0 Then
      Ms0 = "INVALID DESCRIPTION"
      Ms1 = ""
      Ms2 = "PLEASE CHECK DESCRIPTION"
      Ms3 = ""
      Ms4 = "Enter a Valid Description"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Function
    End If
    If LinkGL Then
      If fpcboAcctNumNa1.ListIndex = -1 Or fpcboAcctNumNa2.ListIndex = -1 Then
        Ms0 = "INVALID ACCOUNT NUMBER"
        Ms1 = ""
        Ms2 = "PLEASE CHECK YOUR ACCT"
        Ms3 = ""
        Ms4 = "Enter a Valid GL Account"
        Ms5 = ""
        Dothemsg
        ChkOKFlag = False
        Exit Function
      End If
    End If
    If fpcboInactive.ListIndex = -1 Then
      Ms0 = "INVALID ENTRY"
      Ms1 = ""
      Ms2 = "PLEASE CHECK INACTIVE FLAG"
      Ms3 = ""
      Ms4 = "Enter a Valid Inactive Option"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Function
    End If
    If fpStandFee < 0 Then
      Ms0 = "INVALID ENTRY"
      Ms1 = ""
      Ms2 = "PLEASE CHECK STANDARD FEE"
      Ms3 = ""
      Ms4 = "Enter a Valid Amount"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Function
    End If
  ChkOK2Save = ChkOKFlag
End Function
Private Function SaveCode()
  Dim CodeFile As Integer, CodeRecNo As Integer
  Dim DCCodeRecLen As Integer, cnt As Integer
  Dim DCCodeRec As DCCatCodeRecType
  Dim ChkFlag As String * 1
  SaveCode = False
  If ChkOK2Save = False Then
    Exit Function
  End If
  ChkFlag$ = Chr$(1)
  DCCodeRecLen = Len(DCCodeRec)
  If fpCodeRecNo > 0 Then
    CodeRecNo = fpCodeRecNo
    CodeFile = FreeFile
    Open DCPath + "DCCode.DAT" For Random Shared As CodeFile Len = DCCodeRecLen
    Get CodeFile, CodeRecNo, DCCodeRec
    Close CodeFile
  Else
    CheckCodes
    If DupFound = True Then
      Exit Function
    End If
    CodeRecNo = GetNumCodeRecs%
    CodeRecNo = CodeRecNo + 1
  End If
  If fpCodeRecNo <= 0 Then
    DCCodeRec.CATCODE = QPTrim$(fpCode.Text)
  End If
  DCCodeRec.CODEDESC = QPTrim$(fpCodeDesc.Text)
  If fpStandFee.Value >= 0 Then
    DCCodeRec.Fee = fpStandFee.Value
  Else
    DCCodeRec.Fee = 0
  End If
    If LinkGL Then
      If fpcboAcctNumNa1.ListIndex > 0 Then
        fpcboAcctNumNa1.col = 1
        DCCodeRec.REVGLNUM = QPTrim$(fpcboAcctNumNa1.ColText)
      Else
        DCCodeRec.REVGLNUM = TempAcct1$
      End If
      If fpcboAcctNumNa2.ListIndex > 0 Then
        fpcboAcctNumNa2.col = 1
        DCCodeRec.CASHACCT = QPTrim$(fpcboAcctNumNa2.ColText)
      Else
        DCCodeRec.CASHACCT = TempAcct2$
      End If
    Else
      DCCodeRec.REVGLNUM = QPTrim$(fpGLAcct1)
      DCCodeRec.CASHACCT = QPTrim$(fpGLAcct2)
    End If
    If fpcboInactive.ListIndex = 1 Then
      DCCodeRec.InactiveFlag = "Y"
    Else
      DCCodeRec.InactiveFlag = "N"
    End If
  CodeFile = FreeFile
  Open DCPath + "DCCode.DAT" For Random Shared As CodeFile Len = DCCodeRecLen
  Put CodeFile, CodeRecNo, DCCodeRec
  SaveCode = True
  Close CodeFile
  DCLog PWUser$ + " Saved --Code - " + QPTrim$(fpCode.Text) + "," + QPTrim$(fpCodeDesc.Text)
  DCLog "GLRev - " + QPTrim$(DCCodeRec.REVGLNUM)
  DCLog "GLCash - " + QPTrim$(DCCodeRec.CASHACCT)
  DCLog "Fee - " + Str(fpStandFee.Value)
  MsgBox "Decal Code Saved", vbOKOnly, "Saved"
End Function

Private Sub StuffBeenEntered()
  Dim cnt As Integer
  cnt = 0
  stuff = False
  If Len(fpCode.Text) <> 0 Then cnt = cnt + 1
  If Len(fpCodeDesc.Text) <> 0 Then cnt = cnt + 1
  If LinkGL Then
    If fpcboAcctNumNa1.ListIndex > -1 Then cnt = cnt + 1
    If fpcboAcctNumNa2.ListIndex > -1 Then cnt = cnt + 1
  End If
  If fpcboInactive.ListIndex > -1 Then cnt = cnt + 1
  If fpStandFee <= 0 Then cnt = cnt + 1
  If cnt > 0 Then
    stuff = True
  End If
End Sub

Private Sub CheckCodes()
  Dim NewCode As String, Thiscode As String
  ReDim DCCatCodeRec(1) As DCCatCodeRecType
  Dim CatCodeRecLen As Integer, CatFile As Integer, NumOfCatRecs As Long
  CatCodeRecLen = Len(DCCatCodeRec(1))
  CatFile = FreeFile
  Open "DCCODE.DAT" For Random As CatFile Len = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen

  DupFound = False 'assume this not dup
  ValidCode = True 'assume this code is ok
  NewCode$ = QPTrim$(fpCode)
  For cnt = 1 To NumOfCatRecs
    Get CatFile, cnt, DCCatCodeRec(1)
    Thiscode$ = QPTrim$(DCCatCodeRec(1).CATCODE)
    If Thiscode$ = NewCode$ Then
      ValidCode = False
      DupFound = True
      Exit For
    End If
  Next cnt
  Close
End Sub
Private Sub Dothemsg()
    Dim FntSize As Integer
    ReDim MsgText(0 To 5) As String
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = Ms0
    MsgText(1) = Ms1
    MsgText(2) = Ms2
    MsgText(3) = Ms3
    MsgText(4) = Ms4
    MsgText(5) = Ms5
    GetOKorNot MsgText(), True
End Sub

Private Sub fpcmdPrintinfo_Click()
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
    'do the graphics
      Printonecode 1
    ElseIf rptopt = 2 Then
    'do the text
      Printonecode 2
    End If
   ActivateControls Me

End Sub
Private Sub Printonecode(rptopt)
  Dim graphicflag As Boolean
  Dim RPTFile As Integer, Tmp1 As String, Tmp2 As String
  Dim ReportFile As String
  Dim Dash80 As String * 78
  ReportFile$ = DCPath + "CodeIndv.RPT"
  If rptopt = 1 Then
    graphicflag = True
  Else
    graphicflag = False
  End If
  Dash80$ = String$(78, "-")
  RPTFile = FreeFile
  Open ReportFile$ For Output As RPTFile
    Print #RPTFile, ""
    Print #RPTFile, ""
    Print #RPTFile, TOWNNAME$
    Print #RPTFile, "Vehicle Decal Code"
    Print #RPTFile, Tab(2); Now
    Print #RPTFile, ""
    Print #RPTFile, Dash80$
    Print #RPTFile, Tab(20); "  Decal Code:  "; QPTrim$(fpCode.Text)
    Print #RPTFile, Tab(20); " Description:  "; QPTrim$(fpCodeDesc.Text)
    Print #RPTFile, Tab(20); "         Fee:  "; Using$("######.##", fpStandFee.Value)
    If LinkGL Then
      fpcboAcctNumNa1.col = 1
      Tmp1$ = QPTrim$(fpcboAcctNumNa1.ColText)
      fpcboAcctNumNa2.col = 1
      Tmp2$ = QPTrim$(fpcboAcctNumNa2.ColText)
    Else
      Tmp1$ = QPTrim$(fpGLAcct1)
      Tmp2$ = QPTrim$(fpGLAcct2)
    End If
    Print #RPTFile, Tab(20); " Rev GL Acct:  "; Tmp1$
    Print #RPTFile, Tab(20); "Cash GL Acct:  "; Tmp2$
    If fpcboInactive.ListIndex = 1 Then
      Print #RPTFile, Tab(20); "    Inactive:  "; "Y"
    Else
      Print #RPTFile, Tab(20); "    Inactive:  "; "N"
    End If
    Print #RPTFile, Dash80$
    Print #RPTFile, Chr$(12)
    Close
 If graphicflag Then
    ARptLineRpt.GetName ReportFile$
    ARptLineRpt.startrpt
  Else
    ViewPrint ReportFile$, "Decal Code Report"
    KillFile "CodeIndv.RPT"
  End If
End Sub
Private Sub fpCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpCodeDesc.SetFocus
  End If
  If KeyCode = vbKeyUp Then
    fpCmdExit.SetFocus
  End If
End Sub

Private Sub fpCodeDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpcboAcctNumNa1.Enabled = True Then
      fpcboAcctNumNa1.SetFocus
    Else
      fpGLAcct1.SetFocus
    End If
  End If
End Sub
Private Sub fpcboAcctNumNa1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa1.ListDown = True
  End If
  If fpcboAcctNumNa1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboAcctNumNa2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCodeDesc.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboAcctNumNa2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa2.ListDown = True
  End If
  If fpcboAcctNumNa2.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpStandFee.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboAcctNumNa1.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpGLAcct1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpGLAcct2.SetFocus
  End If
End Sub
Private Sub fpGLAcct2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpStandFee.SetFocus
  End If
End Sub

Private Sub fpStandFee_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboInactive.SetFocus
  End If
End Sub
Private Sub fpcboInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboInactive.ListDown = True
  End If
  If fpcboInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpCmdSave.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpStandFee.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

