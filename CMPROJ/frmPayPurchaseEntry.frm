VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPayDecalEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Management Decal Purchase Entry"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmPayPurchaseEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpDecalCat 
      Height          =   375
      Left            =   2310
      TabIndex        =   2
      Top             =   3285
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      Columns         =   3
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
      BorderDropShadowWidth=   1
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
      AutoSearchFillDelay=   500
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPayPurchaseEntry.frx":08CA
   End
   Begin LpLib.fpCombo fpcboTenderType 
      Height          =   345
      Left            =   9645
      TabIndex        =   12
      Top             =   4080
      Width           =   2250
      _Version        =   196608
      _ExtentX        =   3969
      _ExtentY        =   609
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
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
      ColDesigner     =   "frmPayPurchaseEntry.frx":0C81
   End
   Begin LpLib.fpCombo fpBusPers 
      Height          =   375
      Left            =   2310
      TabIndex        =   9
      Top             =   6120
      Width           =   600
      _Version        =   196608
      _ExtentX        =   1058
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
      ColDesigner     =   "frmPayPurchaseEntry.frx":0FE0
   End
   Begin EditLib.fpText fpPaymentRec 
      Height          =   300
      Left            =   552
      TabIndex        =   66
      Top             =   48
      Visible         =   0   'False
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
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
      ThreeDInsideStyle=   0
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
      ControlType     =   1
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdCharge 
      Height          =   375
      Left            =   6180
      TabIndex        =   22
      Top             =   7635
      Width           =   1290
      _Version        =   131072
      _ExtentX        =   2275
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":134A
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdCheck 
      Height          =   375
      Left            =   4650
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":1527
   End
   Begin EditLib.fpLongInteger fpAcct 
      Height          =   324
      Left            =   1704
      TabIndex        =   0
      Top             =   1392
      Width           =   1872
      _Version        =   196608
      _ExtentX        =   3302
      _ExtentY        =   572
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
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/14/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtPaymentDate 
      Height          =   324
      Left            =   10080
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1416
      Width           =   1548
      _Version        =   196608
      _ExtentX        =   2730
      _ExtentY        =   572
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
      ControlType     =   1
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   300
      Left            =   8160
      TabIndex        =   15
      Top             =   6936
      Width           =   2448
      _Version        =   196608
      _ExtentX        =   4318
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
      MaxLength       =   10
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
   Begin EditLib.fpCurrency fpChangeDue 
      Height          =   312
      Left            =   9648
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5892
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpTotReceived 
      Height          =   312
      Left            =   9648
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5316
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpTAmtOwed 
      Height          =   324
      Left            =   9648
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3744
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   572
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   "."
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   375
      Left            =   9180
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":27F9
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   375
      Left            =   10695
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1230
      _Version        =   131072
      _ExtentX        =   2170
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":29D5
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdCash 
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1140
      _Version        =   131072
      _ExtentX        =   2011
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":2BB1
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdInfo 
      Height          =   375
      Left            =   1785
      TabIndex        =   19
      Top             =   7635
      Width           =   1200
      _Version        =   131072
      _ExtentX        =   2117
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":3E82
   End
   Begin EditLib.fpText fpCustRecNo 
      Height          =   324
      Left            =   528
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   624
      Visible         =   0   'False
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      NoSpecialKeys   =   3
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
      Text            =   "fpText1"
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
   Begin fpBtnAtlLibCtl.fpBtn fpcmdFind 
      Height          =   375
      Left            =   7740
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7635
      Width           =   1200
      _Version        =   131072
      _ExtentX        =   2117
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":405D
   End
   Begin EditLib.fpText fpVehRecNo 
      Height          =   300
      Left            =   432
      TabIndex        =   49
      Top             =   336
      Visible         =   0   'False
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
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
      ThreeDInsideStyle=   0
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
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpChkAmt 
      Height          =   324
      Left            =   9648
      TabIndex        =   14
      Top             =   4728
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   572
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin EditLib.fpDoubleSingle fpCashAmt 
      Height          =   324
      Left            =   9648
      TabIndex        =   13
      Top             =   4392
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   572
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDrawer 
      Height          =   375
      Left            =   300
      TabIndex        =   18
      Top             =   7635
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPayPurchaseEntry.frx":532E
   End
   Begin EditLib.fpText fpStateLic 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2304
      TabIndex        =   5
      Top             =   4344
      Width           =   3324
      _Version        =   196608
      _ExtentX        =   5863
      _ExtentY        =   572
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
      MarginTop       =   0
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
   Begin EditLib.fpDateTime fpExpireDate 
      Height          =   324
      Left            =   2304
      TabIndex        =   8
      Top             =   5400
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   572
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
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
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
   Begin EditLib.fpDoubleSingle fpFee 
      Height          =   348
      Left            =   2304
      TabIndex        =   3
      Top             =   3648
      Width           =   1380
      _Version        =   196608
      _ExtentX        =   2434
      _ExtentY        =   614
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
   Begin EditLib.fpText fpstickernum 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3168
      TabIndex        =   1
      Top             =   2040
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
      _ExtentY        =   572
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
      MarginTop       =   0
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
      MaxLength       =   12
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
   Begin EditLib.fpText fpMakeModl 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   2304
      TabIndex        =   4
      Top             =   4008
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
      _ExtentY        =   572
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
      MarginTop       =   0
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
      MaxLength       =   25
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
   Begin EditLib.fpText fpVINDesc 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   2304
      TabIndex        =   6
      Top             =   4680
      Width           =   4356
      _Version        =   196608
      _ExtentX        =   7683
      _ExtentY        =   614
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
      MarginTop       =   0
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
      MaxLength       =   40
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
   Begin EditLib.fpText fpNotes 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   2304
      TabIndex        =   7
      Top             =   5040
      Width           =   4356
      _Version        =   196608
      _ExtentX        =   7683
      _ExtentY        =   614
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
      MarginTop       =   0
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
      MaxLength       =   40
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
   Begin EditLib.fpText fpResident 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   6000
      TabIndex        =   10
      Top             =   6264
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   572
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
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "~-0123456789"
      MaxLength       =   1
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
   Begin EditLib.fpText fpOwner 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   6000
      TabIndex        =   11
      Top             =   6600
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   572
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
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "~-0123456789"
      MaxLength       =   1
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "As of January 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4344
      TabIndex        =   65
      Top             =   5976
      Width           =   2220
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Residential (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3648
      TabIndex        =   64
      Top             =   6288
      Width           =   2220
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Owner (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3648
      TabIndex        =   63
      Top             =   6624
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Category:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   528
      TabIndex        =   62
      Top             =   3336
      Width           =   1692
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VIN#/Desc:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   744
      TabIndex        =   61
      Top             =   4740
      Width           =   1476
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1296
      TabIndex        =   60
      Top             =   5088
      Width           =   924
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Expires:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   720
      TabIndex        =   59
      Top             =   5424
      Width           =   1500
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fee Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   456
      TabIndex        =   58
      Top             =   3696
      Width           =   1740
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Veh Make/Mdl:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   528
      TabIndex        =   57
      Top             =   4044
      Width           =   1692
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State License #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   480
      TabIndex        =   56
      Top             =   4392
      Width           =   1740
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business/Personal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   55
      Top             =   6144
      Width           =   2100
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6816
      Y1              =   2592
      Y2              =   2592
   End
   Begin VB.Label fptxtCity 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   7872
      TabIndex        =   54
      Top             =   2712
      Width           =   3924
   End
   Begin VB.Label fptxtAddress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   7872
      TabIndex        =   53
      Top             =   2376
      Width           =   3924
   End
   Begin VB.Label fptxtName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   7872
      TabIndex        =   52
      Top             =   2040
      Width           =   3924
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Left            =   4272
      TabIndex        =   51
      Top             =   1512
      Width           =   1824
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6192
      TabIndex        =   50
      Top             =   1464
      Width           =   1860
   End
   Begin VB.Shape Shape3 
      Height          =   612
      Left            =   216
      Top             =   7464
      Width           =   11796
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   828
      Left            =   216
      Top             =   1032
      Width           =   11796
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H8000000E&
      Height          =   5604
      Left            =   216
      Top             =   1848
      Width           =   11796
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Decal Sticker Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   528
      TabIndex        =   44
      Top             =   2064
      Width           =   2532
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Decal Payment/Purchase Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3468
      TabIndex        =   43
      Top             =   516
      Width           =   5268
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2580
      Top             =   432
      Width           =   7020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Left            =   4272
      TabIndex        =   42
      Top             =   1176
      Width           =   1824
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   360
      TabIndex        =   41
      Top             =   1128
      Width           =   2856
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   6804
      TabIndex        =   40
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Index           =   0
      Left            =   7728
      TabIndex        =   39
      Top             =   3768
      Width           =   1728
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   8352
      TabIndex        =   38
      Top             =   1488
      Width           =   1584
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Left            =   7644
      TabIndex        =   37
      Top             =   5352
      Width           =   1812
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   6840
      TabIndex        =   36
      Top             =   3216
      Width           =   5184
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Paid:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   7188
      TabIndex        =   35
      Top             =   4416
      Width           =   2268
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   7872
      TabIndex        =   34
      Top             =   4080
      Width           =   1584
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   6528
      TabIndex        =   33
      Top             =   2496
      Width           =   1248
   End
   Begin VB.Label lblchange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   7608
      TabIndex        =   32
      Top             =   5904
      Width           =   1872
   End
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check/Charge Amt Paid:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6984
      TabIndex        =   31
      Top             =   4752
      Width           =   2472
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6840
      TabIndex        =   30
      Top             =   6960
      Width           =   1224
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   9312
      X2              =   12012
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      X1              =   6816
      X2              =   6816
      Y1              =   2592
      Y2              =   7464
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Vehicle Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   240
      TabIndex        =   29
      Top             =   2592
      Width           =   6564
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   312
      Left            =   8280
      TabIndex        =   28
      Top             =   1152
      Width           =   1656
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6192
      TabIndex        =   27
      Top             =   1128
      Width           =   732
   End
   Begin VB.Label lblSource 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   10080
      TabIndex        =   26
      Top             =   1128
      Width           =   1560
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2592
      Top             =   312
      Width           =   7020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPayDecalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CashFlag As Boolean, uselook As Boolean, CustAcct As Long
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean, PayListCnt As Long
Dim fromform As Form, toform As Form, codeopt As Integer, noreset As Boolean
Dim Oper As String, PayListRec As Long, RecpPort As String, DefPayDate As String
Dim RctValidate As Boolean, NewVeh As Boolean, VehAcct As Long
Dim PayList() As PayListType
Dim DefExpDate As String, CmNum As Long, Txchk As Boolean, Didchk As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional DDate As String)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  uselook = True
  If DDate <> "" Then
    DefPayDate = DDate
  End If
  If codeopt = 1 Then
    Unload frmDCEditLookUP
  ElseIf codeopt = 2 Then
    Unload frmDCDisplayList
  End If

End Sub
Private Sub Form_Activate()
  If Val(fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    fpAcct = fpCustRecNo
    GetCustinfo
    DoEvents
  End If
End Sub

Private Sub cmdExit_Click()
  ChkEmptyAcct
  noreset = True
  Chk4Change
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    CheckInfo
    If ChkOKFlag Then
      fpCmdSave_Click
    Else
      Exit Sub
    End If
  End If
  CustAcct = 0
  fpCustRecNo = 0
  fpVehRecNo = 0
  fpPaymentRec = 0
'  If codeopt = 1 Then
'    ActivateControls frmDCEditLookUP
'  ElseIf codeopt = 2 Then
'    ActivateControls frmDCDisplayList
'  End If
'  If codeopt = 0 Then
'    Load frmDCPurchaseMenu
'    DoEvents
'    frmDCPurchaseMenu.Show
'  End If
'  Erase PayList
'  EditFlag = False
'  NewVeh = False
'  BeenDone = False
    Load frmCMPaySource
    DoEvents
    frmCMPaySource.Show
  
  CMLog "OUT: CMDecal Payment" + " Oper:" + Oper$
  DCLog "OUT: CMDecal Payment" + " Oper:" + Oper$
  Unload Me
  DoEvents
End Sub

Private Sub fpAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp, vbKeyTab
      KeyCode = 0
    If Len(fpAcct) > 0 Then
      fpstickernum.SetFocus
'      If fpcboTenderType.Enabled = True Then
'        fpcboTenderType.SetFocus
'      End If
    End If
  End Select
End Sub
Private Sub ChkEmptyAcct()
  If Len(fpAcct) <= 0 Then
    ClearScn
  End If
End Sub


Private Sub fpBusPers_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpBusPers.ListDown = True
  End If
  If fpBusPers.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpResident.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpExpireDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpCmdInfo_Click()
  If Len(fpAcct) > 0 Then
    'DeActivateControls Me
    frmInfo.Label1 = "Loading. . ."
    frmInfo.Show
    DoEvents
    'here
    frmRptDCCustInq.fpCustRecNo = Me.fpCustRecNo
    'frmRptCustInq.Wheretogo frmPayPurchaseEntry, frmRptCustInq, 0
    'Load frmRptCustInq
    frmRptDCCustInq.Show
    DoEvents
    Unload frmInfo
  End If
End Sub
Private Sub fpDecalCat_GotFocus()
  If Len(fpstickernum) <= 0 Then
    fpstickernum.SetFocus
  End If
End Sub

Private Sub fpExpireDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpExpireDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpBusPers.SetFocus
  End If
End Sub

Private Sub fpFee_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpFee_LostFocus()
  fpTAmtOwed = fpFee
End Sub

Private Sub fpFee_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyTab
      If fpFee > 0 Then
        'fpTAmtOwed = fpFee
        fpMakeModl.SetFocus
      Else
        fpFee.SetFocus
      End If
    Case vbKeyUp
      If fpFee > 0 Then
        'fpTAmtOwed = fpFee
        fpDecalCat.SetFocus
      Else
        'fpFee.SetFocus
      End If
    Case Else
  End Select
End Sub

Private Sub fpCashAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCashAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpChkAmt.Enabled Then
      fpChkAmt.SetFocus
    Else
      fptxtDesc.SetFocus
    End If
  End If
End Sub


Private Sub fpChkAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpChkAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtDesc.SetFocus
  End If
End Sub

Private Sub Chk4Change()
  Dim cntout As Integer, cnt As Integer
  If Len(fpAcct) > 0 Then
    cntout = 0
    Answer = 0
    If fpTotReceived <> 0 Or fpFee <> 0 Then cntout = cntout + 1
    If cntout > 0 Then
      frmChangedWarning.Show vbModal, Me
      Select Case SaveFlag
      Case False
        Answer = 3
      Case True
        Answer = 2
      Case 1
        Answer = 1
      End Select
    Else
      Answer = 0
    End If
  End If
End Sub
Private Sub Chk4OKforNew()
  Dim FntSize As Integer
  Dim cntout As Integer, cnt As Integer
  If Len(fpAcct) > 0 Then
    cntout = 0
    Answer = 0
    If fpTotReceived <> 0 Or fpFee <> 0 Then cntout = cntout + 1
    If cntout > 0 Then
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      frmMsgDialog.Label(2).FontSize = (FntSize + 2)
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING:Payment In Progress"
      MsgText(1) = ""
      MsgText(2) = "Do You Want to Abandon this Payment?"
      MsgText(3) = "Ok to Abandon,"
      MsgText(4) = "Cancel to Remain on Current Payment."
      MsgText(5) = ""
      If GetOKorNot(MsgText()) Then
       DCLog "USER WANTS TO Abandon"
       Answer = 2
      Else
       DCLog "USER Canceled"
       Answer = 1
      End If
    Else
      Answer = 0
    End If
  End If
End Sub

Private Sub fpcmdDrawer_Click()
  Dim Port As String, PortFile As Integer
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  Port$ = QPTrim$(RecpPort)
  DCLog "Oper: " + Oper$ + "Decal Pay-Open Drawer"
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  Print #PortFile, Chr$(7)
  Close PortFile
End Sub

Private Sub fpcmdFind_Click()
  Chk4OKforNew
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    'continue on
  End If
  ClearScn
  frmDCEditLookUP.Caption = "Decal Customer Find"
  frmDCEditLookUP.Label1.Caption = "Decal Customer Find"
  frmDCEditLookUP.Wheretogo frmPayDecalEntry, frmPayDecalEntry
  Unload Me
  DoEvents
  frmDCEditLookUP.Show
  DoEvents
End Sub

Private Sub fpCmdSave_Click()
On Local Error GoTo ERRORSTUFF
  ChkEmptyAcct
  DoEvents
  If Len(fpAcct) <= 0 Then
    MsgBox "Invalid Account Information.", vbOKOnly, "Invalid Entry"
    Exit Sub
  End If
  If fpVehRecNo <= 0 And NewVeh = False Then
    MsgBox "Invalid Vehicle Selection. Please Select from List.", vbOKOnly, "Invalid Entry"
    Exit Sub
  End If

  CalcBALFlds
  CheckInfo
  If ChkOKFlag Then
    'DeActivateControls Me
    If fpcboTenderType.ListIndex = 1 Or fpcboTenderType.ListIndex = 2 Then
      frmPrintReceipt.setvallist = 1
    Else
      frmPrintReceipt.setvallist = 0
    End If

    frmPrintReceipt.Show 1
    CheckInfo
    If ChkOKFlag Then
      If SavePay = True Then
        SaveTransaction
      
        If PrnRecp = True Or PrnVali = True Then
          PrintReceipt
        End If
  
        MsgBox "Transaction Saved.", vbOKOnly, "Complete"
  
        ClearScn
  
      End If
    End If
  End If
Exit Sub
ERRORSTUFF:
  DCLog PWUser + " Error " + Str(Err.Number) + " CMPayEntry, cmdSave"
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "DC CMPayEntry", "cmdSave", Erl)
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
    Unload Me

End Sub

Private Sub fpcmdCash_Click()
If Len(fpAcct) > 0 Then
  fpcboTenderType.ListIndex = 0
  fpChkAmt.Enabled = False
  fpCashAmt.Enabled = True
  fpChkAmt = 0
  fpCashAmt = fpTAmtOwed.DoubleValue
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
   ' Autodist
  End If
  fptxtDesc.SetFocus
End If
End Sub

Private Sub fpcmdCheck_Click()
If Len(fpAcct) > 0 Then
  fpcboTenderType.ListIndex = 1
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt = fpTAmtOwed.DoubleValue
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
   ' Autodist
  End If
  fptxtDesc.SetFocus
End If
End Sub
Private Sub fpCmdCharge_Click()
If Len(fpAcct) > 0 Then
  fpcboTenderType.ListIndex = 3
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt = fpTAmtOwed.DoubleValue
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    TempAmtRecv = fpTotReceived
    '
  End If
  fpChangeDue.Enabled = False
  fptxtDesc.SetFocus
End If
End Sub

Private Sub fpCashAmt_LostFocus()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
If fpTotReceived > 0 Then
  If fpcboTenderType.ListIndex <> 3 Then
    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpFee.DoubleValue)
  End If
End If
End Sub

Private Sub fpChkAmt_LostFocus()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
If fpTotReceived > 0 Then
  If fpcboTenderType.ListIndex <> 3 Then
    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpFee.DoubleValue)
  End If
End If
End Sub
Private Sub fpcboTenderType_DropDown()
  ClrAmts
End Sub
'
Private Sub fpcboTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTenderType.ListDown = True
    'ClrAmts
    KeyCode = 0
  End If
  If KeyCode = vbKeyDelete Then
    fpcboTenderType.ListIndex = -1
    fpcboTenderType.Action = ActionClearSearchBuffer
    'ClrAmts
  End If
  If fpcboTenderType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fpCashAmt.Enabled = True Then
        fpCashAmt.SetFocus
      Else
        fpChkAmt.SetFocus
      End If
        KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpAcct.SetFocus
        KeyCode = 0
      End If
    End If
  End If

  DoEvents
End Sub
Private Sub ClrAmts()
  Dim cnt As Integer
  fpCashAmt = 0
  fpChkAmt = 0
  fpChangeDue.Enabled = True
  fpChangeDue = 0
  fpTotReceived = 0
End Sub
Private Sub fpcboTenderType_SelChange(ItemIndex As Long)
  If BeenDone Then
    fixamts
 End If
End Sub

Private Sub fixamts()

  fpcboTenderType.Action = ActionClearSearchBuffer
  If noreset = False Then
    If fpcboTenderType.ListIndex = 0 Then
      fpCashAmt.Enabled = True
      fpChkAmt = 0
      fpChkAmt.Enabled = False
      fpChangeDue.Enabled = True
      'ClrAmts
     ' fpCashAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 1 Then
      fpCashAmt.Enabled = False
      fpCashAmt = 0
      fpChkAmt.Enabled = True
      fpChangeDue.Enabled = True
      'ClrAmts
     ' fpChkAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 2 Then
      fpCashAmt.Enabled = True
      fpChkAmt.Enabled = True
      fpChangeDue.Enabled = True
     ' ClrAmts
     'fpCashAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 3 Then
      fpCashAmt.Enabled = False
      fpCashAmt = 0
      fpChkAmt.Enabled = True
      fpChangeDue = 0
      fpChangeDue.Enabled = False
     ' ClrAmts
      'fpChkAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = -1 Then
'      MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
'      fpcboTenderType.SetFocus
    End If
  End If
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    If fpcboTenderType.ListIndex <> 3 Then
      fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTAmtOwed.DoubleValue)
    End If
  End If
  DoEvents
  noreset = False
End Sub
Private Sub fpAcct_LostFocus()
'Dim Acct As Long
'    Acct = fpAcct
'    If Acct > 0 Then
'      If Acct > GetTaxCustCnt Then
'        MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
'        fplngAcct.SetFocus
'        Exit Sub
'      ElseIf IsCustDeleted(Acct) Then
'        MsgBox "Deleted Account.", vbOKOnly, "Deleted Account"
'        fplngAcct.SetFocus
'        Exit Sub
'      Else
'       'If DoesCustOwe(Acct) Then
'          Cust2Screen (Acct)
'       ' Else
'       '   MsgBox "This Customer Does Not Owe A Balance.", vbOKOnly, "No Balance"
'      End If
'    Else
'      MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
'      fplngAcct.SetFocus
'      Exit Sub
'    End If
    Chk4OKforNew
    If Answer = 1 Then
      fpAcct = fpCustRecNo
      Exit Sub
    ElseIf Answer = 2 Then
      'continue on
    ElseIf fptxtName.Caption <> "" And fpAcct = fpCustRecNo Then
      Exit Sub
    End If
     
  fpCustRecNo = fpAcct
  If Val(fpCustRecNo) > 0 Then
    Label8.Caption = "Vehicle Payment"
    txtPaymentDate = DefPayDate
    NewVeh = False
    fptxtName.Caption = ""
    fptxtAddress.Caption = ""
    fptxtCity.Caption = ""
    fptxtDesc = ""
    fpVehRecNo = 0
    fpPaymentRec = 0
    fpcboTenderType.ListIndex = -1
    fpCashAmt = 0
    fpChkAmt = 0
    fpChangeDue = 0
    'fpTotOwed = 0
    fpFee = 0
    fpTAmtOwed = 0
    fpTotReceived = 0
    fpstickernum = ""
    fpVinDesc = ""
    fpMakeModl = ""
    fpStateLic = ""
    fpNotes = ""
    fpBusPers.ListIndex = -1
    fpDecalCat.ListIndex = -1
    fpFee = 0
    fpExpireDate = DefExpDate
    fpResident = ""
    fpOwner = ""
    PayListRec& = 0
   ' EditFlag = False
   GetCustinfo
  'Else
    'MsgBox "NO", vbOKOnly
  '  fpAcct.SetFocus
  End If
End Sub


Private Sub fpMakeModl_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpMakeModl_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpStateLic.SetFocus
  End If
End Sub

Private Sub fpNotes_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpNotes_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpExpireDate.SetFocus
  End If
End Sub

Private Sub fpOwner_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpOwner_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboTenderType.SetFocus
  End If
End Sub

Private Sub fpResident_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpResident_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpOwner.SetFocus
  End If
End Sub

Private Sub fpStateLic_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpStateLic_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpVinDesc.SetFocus
  End If
End Sub



Private Sub fptxtDesc_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpCmdSave.SetFocus
  End If
End Sub

Private Sub fpVINDesc_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpVINDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpNotes.SetFocus
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via PaymentEntry by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      If cmdExit.Enabled Then
        Call cmdExit_Click
      End If
    Case vbKeyF2:
      KeyCode = 0
      DoEvents
      fpcmdDrawer_Click
    Case vbKeyF4:
      KeyCode = 0
      DoEvents
      If fpCmdInfo.Enabled Then
        Call fpCmdInfo_Click
      End If
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      If fpCmdCash.Enabled Then
        Call fpcmdCash_Click
      End If
    Case vbKeyF6:
      KeyCode = 0
      DoEvents
      If fpcmdCheck.Enabled Then
        Call fpcmdCheck_Click
      End If
    Case vbKeyF7:
      KeyCode = 0
      DoEvents
      If fpcmdFind.Enabled Then
        Call fpcmdFind_Click
      End If
    Case vbKeyF8:
      KeyCode = 0
      DoEvents
      If fpCmdCharge.Enabled Then
        Call fpCmdCharge_Click
      End If
'    Case vbKeyF9:
'      KeyCode = 0
'      DoEvents
'      If fpCmdDist.Enabled Then
'        Call fpCmdDist_Click
'      End If
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      If fpCmdSave.Enabled Then
        Call fpCmdSave_Click
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim DCSetuplen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  If Exist(UBPath$ + "DCSetup.DAT") Then
    ReDim DCSetup(1) As DCSetupType
    LoadDCSetUpFile DCSetup(), DCSetuplen
    If QPTrim(DCSetup(1).Taxbalchk) = "Y" Then
      Txchk = True
    Else
      Txchk = False
    End If
  Else
    Txchk = False
  End If
  StatusBar1.Panels.Item(1).Text = TownName$
  txtPaymentDate.Text = DefPayDate
  FillCatCMBO fpDecalCat
  noreset = False
  fpcboTenderType.AddItem "Cash"
  fpcboTenderType.AddItem "Check"
  fpcboTenderType.AddItem "Cash & Check"
  fpcboTenderType.AddItem "Charge"
  lblOperator = OperNum
  lblOperName.Caption = PWUser
  lblSource.Caption = "Decal"
  Oper$ = QPTrim(lblOperator.Caption)
  DCLog " IN Oper " + Oper$ + ": Decal Payment"
 ' LoadPayList
  GetRcpInfo
  Dim DCFile As Integer
  Dim DCDate(1) As DCExpireDate
  DCFile = FreeFile
  If Exist("DCEXPIRE.DAT") Then
    Open "DCEXPIRE.DAT" For Random Access Read Write Shared As DCFile Len = 10
    Get #DCFile, 1, DCDate(1)
    DefExpDate = DCDate(1).ExpireDate
    Close
  Else
    DefExpDate = DefPayDate
  End If
  fpExpireDate = DefExpDate
  fpBusPers.AddItem "P"
  fpBusPers.AddItem "B"
End Sub
Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
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
      ValiDef = 1
      RctValidate = True
      GetUBBankINfo
    Else
      ValiDef = 0
      RctValidate = False
    End If
  Close RP1
  Else
    ValiDef = 0
    RecpDef = 99
  End If
Exit Sub
nofound:
  ValiDef = 0
  RecpDef = 99
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
  '''' Me.Visible = False
    Temp_Class.ResizeControls Me
  '  Me.Visible = True
  '  Me.SetFocus
  End If
  DoEvents
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub
Private Sub GetCustinfo()
  Dim DCCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile As Integer, cnt As Integer, TotalBalance As Double
  Dim ListFile As Integer, SSNSrch As String
  Dim PayFileName As String, DCPayRecLen As Integer
  ReDim DCCustREc(1) As DCCustRecType
  ReDim DCPaymentRec(1) As DCEditPaymentRecType
  DCPayRecLen = Len(DCPaymentRec(1))
  DCCustRecLen = Len(DCCustREc(1))
  On Local Error GoTo errstuff
  CashFlag = False
  If uselook = True Then
    Unload frmDCEditLookUP
    Unload frmDCDisplayList
    uselook = False
  End If
10:
  If fpCustRecNo <> "" Then
       CustAcct = fpCustRecNo
  Else
    'MsgBox "You Must Enter An Account Number.", vbOKOnly, "Invalid Account"
    fpAcct.SetFocus
    Exit Sub
  End If
20:
  NumOfCustRecs& = FileSize(UBPath$ + "DCCUST.DAT") \ DCCustRecLen
  If CustAcct& > NumOfCustRecs& Or CustAcct& <= 0 Then
    DCLog "ERROR: Invalid Account:" + Str$(CustAcct&) + " Oper:" + Oper$
    CustAcct& = 0
    'LabelDel.Visible = True
    GoTo SkiptoHere
  End If
30:
  If IsDCDeleted(CustAcct&) Then
    DCLog "ERROR: Deleted Account:" + Str$(CustAcct&) + " Oper:" + Oper$
    CustAcct& = 0
    'LabelDel.Caption = "Deleted Account!"
    'LabelDel.Visible = True
    GoTo SkiptoHere
  End If
 ' GoSub ClearForm
40:
  CustFile = FreeFile
  Open UBPath$ + "DCCUST.DAT" For Random Shared As CustFile Len = DCCustRecLen
  Get CustFile, CustAcct&, DCCustREc(1)
  'FOR Cnt = 1 TO 15
  '  UBCustRec(1).CurrRevAmts(Cnt) = 0
  'NEXT
  'PUT CustFile, CUSTACCT&, UBCustRec(1)
  Close CustFile
50:
  SSNSrch = QPStrip$(DCCustREc(1).SOSEC)
  DCLog "Oper:" + Oper$ + " Payment for Account:" + Str$(CustAcct&)
  If DCCustREc(1).CASHONLY = "Y" Then
    CashFlag = True
  End If
60:
    fpcboTenderType.Clear
    fpcboTenderType.AddItem "Cash"
    If CashFlag Then
      fpcboTenderType.ListIndex = 0
      fpCmdCharge.Enabled = False
      fpcmdCheck.Enabled = False
    Else
      fpcboTenderType.AddItem "Check"
      fpcboTenderType.AddItem "Cash & Check"
      fpcboTenderType.AddItem "Charge"
      fpcboTenderType.ListIndex = -1
      fpCmdCharge.Enabled = True
      fpcmdCheck.Enabled = True
    End If
70:
    TotalBalance# = DCCustREc(1).AcctBal
    'LSet Form$(CustAcctFld, 0) = Str$(CustAcct&)
    txtPaymentDate = DefPayDate
    fptxtName.Caption = DCCustREc(1).BILLNAME
    fptxtAddress.Caption = DCCustREc(1).ADDRESS1
80:
    If CashFlag Then
      Label4.ForeColor = &HFFFF&
      Label4.Caption = "Cash Only"
    Else
      Label4.Caption = ""
      Label4.ForeColor = &H80000012
    End If
    fpTAmtOwed = 0
    fpCashAmt = 0
    fpChkAmt = 0
    fpTotReceived = 0
    fpChangeDue = 0
90:
  CustAcct& = Val(fpCustRecNo)
  fptxtCity.Caption = DCCustREc(1).City
  fpResident = DCCustREc(1).resident
  fpOwner = DCCustREc(1).Owner
  BeenDone = True
      fpDecalCat.ListIndex = -1
      fpFee = 0
      fpMakeModl = ""
      fpStateLic = ""
      fpExpireDate = ""
      fpstickernum = ""
      fpVinDesc = ""
      fpNotes = ""
      fpBusPers.ListIndex = -1
98:



  If Txchk And Didchk = False Then
    If Len(SSNSrch) > 0 Then
      If Chk4BalinTax(SSNSrch) Then
        CustAcct& = 0
        GoTo skiptonewone
      End If
    End If
  End If
  Exit Sub
SkiptoHere:
  BeenDone = True
  frmLookupError.Label.Caption = "Invalid Account Number"
  frmLookupError.Label1.Caption = "Please Enter A Valid Account Number."
  frmLookupError.Show 1
skiptonewone:
  BeenDone = True
  ClearScn
  ''Return
errstuff:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " ( Line:" & Erl & ")"
  End If
  Close
  Exit Sub
End Sub
Private Function Chk4BalinTax(SSN As String)
  Dim Chk4Bal As Boolean, TXBAL As Double
  Dim FntSize As Integer, cnt As Long
  Dim lurec As Long, vernum As Integer
  Dim RcptPrnFile As ReceiptPRNType
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim DispTaxfound As Integer
  Dim DispTaxfndNm As String
  vernum = 0
  Chk4Bal = False
  If Not Exist("Taxcust.dat") Then
    Exit Function
  End If
  OpenVATaxCustFile CHandle, NumOfCRecs
  'REM OUT 4/11/2011 - CODE TO MATCH THE CODE IN DECALS NOT TO CHECK VERSION, WAS NOT CHECKING IN APPALACHIA IN CM AS IT DID IN DECALS.
   '    For cnt = 1 To 10
'      Get CHandle, cnt, TaxCustRec
'      If TaxCustRec.FileVer = 205 Then
'        vernum = vernum + 1
'      End If
'    Next
vernum = 10
    If vernum = 10 Then
      For cnt = 1 To NumOfCRecs
        Get CHandle, cnt, TaxCustRec
        If QPTrim$(TaxCustRec.CSSN) = SSN Or QPTrim$(TaxCustRec.OSSN) = SSN Then
          'Chk4Bal = True
          If TaxCustRec.Deleted <> -1 Then
          lurec = cnt
          TXBAL = VAGetCustBalance#(lurec, -1)
          If TXBAL > 0 Then
          DispTaxfound = TaxCustRec.Acct
          DispTaxfndNm = TaxCustRec.CustName
          Exit For
          End If
          End If
        End If
      Next
    End If
  Close CHandle
    If TXBAL > 0 Then
      Chk4Bal = True
'      If Didchk = True Then  'this caused loop
'        Unload frmMsgDialog
'      End If
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(0).FontSize = (FntSize + 2)
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      frmMsgDialog.Label(2).FontSize = (FntSize + 2)
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      frmMsgDialog.Label(4).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING: Tax Balance Owed"
      MsgText(1) = "Matching Social Sec# " + QPTrim$(SSN)
      MsgText(2) = "Found On Tax Acct# - " + Str$(DispTaxfound)
      MsgText(3) = "For - " + QPTrim$(DispTaxfndNm)
      MsgText(4) = "that has a Tax Balance Due of " & Str$(TXBAL)
      MsgText(5) = "OK to continue Or CANCEL to select another"
      Didchk = True
      
       Dim zz As Integer, RetValue As Integer
        frmMsgDialog.Caption = MsgText(0)
        For zz = 1 To 5
          frmMsgDialog.Label(zz - 1) = MsgText(zz)
        Next
        frmMsgDialog.Show vbModal
        RetValue = Val(frmMsgDialog.RetLabel)
      If RetValue Then
        DCLog "USER WANTS TO CONTINUE!"
        Chk4Bal = False
        Didchk = False
        fpstickernum.SetFocus
      Else
        DCLog "USER ABORTED."
        Chk4Bal = True
        'Exit Function
        Didchk = False
      End If
      Unload frmMsgDialog
   
    Else
      Chk4Bal = False
    End If
  
  Chk4BalinTax = Chk4Bal
End Function

Private Sub fpDecalCat_Change()
  Dim lookrec As Integer
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer
  Dim NumOFDCCatRecs As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  If fpDecalCat.ListIndex <> -1 Then
    fpDecalCat.col = 0
    lookrec = QPTrim$(fpDecalCat.ColText)
    ghandle = FreeFile
    Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
    NumOFDCCatRecs = LOF(ghandle) \ DCCatCodeRecLen
    Get #ghandle, lookrec, DCCatCodeRec
      fpFee = DCCatCodeRec.Fee
      fpTAmtOwed = DCCatCodeRec.Fee
    Close ghandle
  End If
End Sub
Private Sub fpDecalCat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpDecalCat.ListDown = True
  End If
  If fpDecalCat.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpFee.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpstickernum.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpstickernum_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp, vbKeyTab
      KeyCode = 0
    If Len(fpstickernum) > 0 Then
      Label8.Caption = "Vehicle Payment"
      txtPaymentDate = DefPayDate
      NewVeh = False
      'EditFlag = False
      fpVehRecNo = 0
      fpPaymentRec = 0
      fpcboTenderType.ListIndex = -1
      fpCashAmt = 0
      fpChkAmt = 0
      fpChangeDue = 0
      'fpTotOwed = 0
      fpFee = 0
      fpTAmtOwed = 0
      fpTotReceived = 0
      fpVinDesc = ""
      fpMakeModl = ""
      fpStateLic = ""
      fpNotes = ""
      fpBusPers.ListIndex = -1
      fpDecalCat.ListIndex = -1
      fpFee = 0
      fpExpireDate = DefExpDate
      fpResident = ""
      fpOwner = ""
      GetVehList
      fpDecalCat.SetFocus
    End If
  End Select
End Sub


Public Sub GetVehInfo()
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  ReDim DCVRec(1) As DCVehType
  ReDim DCPaymentRec(1) As DCEditPaymentRecType
  Dim PayFileName As String, ListFile As Integer, DCPayRecLen As Integer
  DCLog "Oper:" + Oper$ + " Entering purchase for Veh:" + Str$(CustAcct&)
  If fpVehRecNo <> 0 Then
    VehAcct = fpVehRecNo
    'CheckPayList
  End If
'  If Not EditFlag Then
    'New Payment entry for existing vehicle
    If fpVehRecNo > 0 Then
      DCVehReclen = Len(DCVRec(1))
      DCvFile = FreeFile
      Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
      NumOfVRecs = LOF(DCvFile) \ DCVehReclen
      Get DCvFile, fpVehRecNo, DCVRec(1)
      Close DCvFile
      Label8.Caption = "Enter Vehicle Payment"
      fpDecalCat.ColumnSearch = 1
      fpDecalCat.SearchText = QPTrim$(DCVRec(1).DecalCat)
      fpDecalCat.Action = ActionSearch
      If fpDecalCat.SearchIndex <> -1 Then
        fpDecalCat.ListIndex = fpDecalCat.SearchIndex
      End If
      fpFee = DCVRec(1).Fee
      fpMakeModl = QPTrim$(DCVRec(1).makemodel)
      fpStateLic = QPTrim$(DCVRec(1).StateTag)
      'fpExpireDate = Num2Date$(DCVRec(1).ExpireDate)
      'fpstickernum = DCVRec(1).Sticker
      fpVinDesc = QPTrim$(DCVRec(1).Desc)
      fpNotes = QPTrim$(DCVRec(1).Notes)
      If DCVRec(1).PBFlag = "P" Then
        fpBusPers.ListIndex = 0
      ElseIf DCVRec(1).PBFlag = "B" Then
        fpBusPers.ListIndex = 1
      Else
        fpBusPers.ListIndex = -1
      End If
      fpTAmtOwed = DCVRec(1).Fee
      NewVeh = False
    Else  'New Payment for new vehicle
      Label8.Caption = "New Vehicle Payment"
      fpDecalCat.ListIndex = -1
      fpFee = 0
      fpMakeModl = ""
      fpStateLic = ""
      'fpExpireDate = ""
      'fpstickernum = ""
      NewVeh = True
      fpVinDesc = ""
      fpNotes = ""
      fpBusPers.ListIndex = -1
      
    End If
    fpExpireDate = DefExpDate
'  Else 'existing payment edit for existing vehicle
'    'get record from payment file
'    NewVeh = False
'    Oper$ = QPTrim$(lblOperator.Caption)
'    DCLog "Oper:" + Oper$ + " Editing payment for Account:" + Str$(CustAcct&)
'    PayFileName$ = UBPath$ + "DCPAY" + Oper$ + ".DAT"
'    DCPayRecLen = Len(DCPaymentRec(1))
'    ListFile = FreeFile
'    Open PayFileName$ For Random Shared As ListFile Len = DCPayRecLen
'    Get ListFile, PayListRec&, DCPaymentRec(1)
'    Close ListFile
'
'    Label8.Caption = "Edit Existing Payment"
'    fpDecalCat.SearchText = QPTrim$(DCPaymentRec(1).DecalCat)
'    fpDecalCat.ColumnSearch = 1
'    fpDecalCat.Action = ActionSearch
'    If fpDecalCat.SearchIndex <> -1 Then
'      fpDecalCat.ListIndex = fpDecalCat.SearchIndex
'    End If
'    txtPaymentDate = Num2Date$(DCPaymentRec(1).TranDate)
'    fpFee = DCPaymentRec(1).Amount
'    fpstickernum = DCPaymentRec(1).Sticker
'    fpVinDesc = DCPaymentRec(1).VinDesc
'    fpExpireDate = Num2Date$(DCPaymentRec(1).ExpDate)
'    fpMakeModl = DCPaymentRec(1).makemodel
'    fpStateLic = DCPaymentRec(1).StateTag
'    fpResident = DCPaymentRec(1).resident
'    fpOwner = DCPaymentRec(1).Owner
'    If DCPaymentRec(1).PersBuss = "P" Then
'      fpBusPers.ListIndex = 0
'    ElseIf DCPaymentRec(1).PersBuss = "B" Then
'      fpBusPers.ListIndex = 1
'    Else
'      fpBusPers.ListIndex = -1
'    End If
'    fptxtDesc = DCPaymentRec(1).PayDesc
'    fpNotes = DCPaymentRec(1).Notes
'    Select Case DCPaymentRec(1).TransTender
'      Case 1:
'        fpcboTenderType.ListIndex = 0
'      Case 2:
'        fpcboTenderType.ListIndex = 1
'      Case 3:
'        fpcboTenderType.ListIndex = 2
'      Case 4:
'        fpcboTenderType.ListIndex = 3
'      Case Else:
'        fpcboTenderType.ListIndex = -1
'    End Select
'    fpCashAmt = DCPaymentRec(1).CashAmt
'    fpChkAmt = DCPaymentRec(1).CheckAmt
'    fpTotReceived = Round(DCPaymentRec(1).CashAmt + DCPaymentRec(1).CheckAmt)
'    fpChangeDue = DCPaymentRec(1).Change
'
'    fpTAmtOwed = DCPaymentRec(1).Amount
'End If
End Sub

Private Sub ClearScn()
  Dim cnt As Integer
  BeenDone = False
  fpAcct.Enabled = True
  fpAcct = ""
  'LabelDel.Visible = False
  'fpCmdTranHist.Enabled = False
  txtPaymentDate = DefPayDate
  NewVeh = False
  fptxtName.Caption = ""
  fptxtAddress.Caption = ""
  fptxtCity.Caption = ""
  fptxtDesc = ""
  fpCustRecNo = 0
  fpVehRecNo = 0
  fpPaymentRec = 0
  fpcboTenderType.ListIndex = -1
  fpCashAmt = 0
  fpChkAmt = 0
  fpChangeDue = 0
  'fpTotOwed = 0
  fpFee = 0
  fpTAmtOwed = 0
  fpTotReceived = 0
  fpstickernum = ""
  fpVinDesc = ""
  fpMakeModl = ""
  fpStateLic = ""
  fpNotes = ""
  fpBusPers.ListIndex = -1
  fpDecalCat.ListIndex = -1
  fpFee = 0
  fpExpireDate = DefExpDate
  fpResident = ""
  fpOwner = ""
  Label4.Caption = ""
  fpAcct.SetFocus
End Sub

Private Sub CalcBALFlds()
  Dim cnt As Integer, TotP As Double
  If fpTAmtOwed.Value = fpFee.Value Then
    TotP# = Round#(Val(fpCashAmt) + Val(fpChkAmt))
    If fpTotReceived.Value > 0 Then
      If fpcboTenderType.ListIndex <> 3 Then
        fpChangeDue.Value = Round#(fpTotReceived.Value - fpFee.Value)
      End If
  End If
  Else
    MsgBox "Verify Fee Amount before continuing.", vbOKOnly, "Fee/Amount Owed"
  End If
End Sub

Private Sub CheckInfo()
  Dim TestDate As Integer, TestAmt As Double
  TestAmt = 0
  ChkOKFlag = True
  TestDate = Date2Num(txtPaymentDate)
  If TestDate < 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Date.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If NewVeh = False And fpVehRecNo < 0 Then
    MsgBox "Invalid Vehicle Selectio. Please Select from List.", vbOKOnly, "Invalid Entry"
    ChkOKFlag = False
    GoTo BadDate
  End If
  If fpDecalCat.ListIndex = -1 Then
    MsgBox "You MUST select a category.", vbOKOnly, "Invalid Code"
    ChkOKFlag = False
    GoTo BadDate
  End If
  If Len(fpMakeModl.Text) <= 0 Then
    ChkOKFlag = False
    MsgBox "Please enter Vehicle Make/Model.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = -1 Then
    MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
    ChkOKFlag = False
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = 0 And fpChkAmt.DoubleValue > 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Tender Type. The Check/Charge Amount Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If (fpcboTenderType.ListIndex = 1 Or fpcboTenderType.ListIndex = 3) And Val(fpCashAmt) > 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Tender Type. The Cash Amount Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = 2 And (fpChkAmt.DoubleValue <= 0 Or fpCashAmt.DoubleValue <= 0) Then
    ChkOKFlag = False
    MsgBox "Invalid Amounts. The Check and Cash Amount Should Be Greater than ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
 
  If fpTotReceived.Value < 0 Or fpTAmtOwed.Value < 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Amount. The Total Received Should NOT Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpChangeDue.Value >= 0 Then
    TestAmt = Round#(fpTotReceived.Value - fpChangeDue.Value)
    If TestAmt <> fpFee.Value Then '.DoubleValue Then
      ChkOKFlag = False
      MsgBox "The Amount Paid does NOT equal the Fee Amount.", vbOKOnly, "Request Canceled."
      GoTo BadDate
    End If
  Else
    ChkOKFlag = False
    MsgBox "The Fee Amount May Not Be More Than Amount Received.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  Exit Sub
BadDate:
  Exit Sub
End Sub

Private Sub PrintReceipt()
  Dim ListFile As Integer, PayFileName As String, DCPayRecLen As Integer
  Dim RecptNum As Long, RHandle As Integer, PayRecpName As String
  Dim CutPaper As String, PostDate As String
  Dim RHandle2 As Integer, PayRecpName2 As String, RptHandle2 As Integer
  ReDim DCPaymentRec(1) As DCEditPaymentRecType
'  ReDim Preserve RevText$(1 To MaxRevsCnt)
  CutPaper$ = Chr$(29) + Chr$(86) + Chr$(66) + Chr$(64)
  DCPayRecLen = Len(DCPaymentRec(1))
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  PayRecpName$ = "c:\CPWork\CMRCP" + Oper$ + ".RPT"
  PayRecpName2$ = "C:\CPWork\CMVLD" + "DCVLD" + Oper$ + ".Rpt"
  PostDate$ = txtPaymentDate
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = DCPayRecLen
  'RecptNum& = LOF(ListFile) / DCPayRecLen
  Get #ListFile, 1, DCPaymentRec(1)
  Close
  If PrnRecp = False And PrnVali = True Then GoTo Validationthing
  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  If CntrlDef = 1 Then
    Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    Print #RHandle, Chr$(7)
  End If
  Print #RHandle, TownName$
  Print #RHandle, "CM DECAL PAYMENT"
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle, "Time: "; Time
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, DCPaymentRec(1).CustName
  Print #RHandle, DCPaymentRec(1).CustAddr
  Print #RHandle, DCPaymentRec(1).PayDesc
  Print #RHandle, "Acct. No. "; DCPaymentRec(1).CustNumber
  Print #RHandle,
   Select Case DCPaymentRec(1).TransTender
    Case 1:
      Print #RHandle, "Cash"
    Case 2:
      Print #RHandle, "Check"
    Case 3:
      Print #RHandle, "Cash & Check"
    Case 4:
      Print #RHandle, "Charge"
    Case Else:
  End Select
  Print #RHandle,
  Print #RHandle, "       Cash: "; Using("$##,###,###.##", DCPaymentRec(1).CashAmt)
  If DCPaymentRec(1).TransTender <> 4 Then
    Print #RHandle, "      Check: "; Using("$##,###,###.##", DCPaymentRec(1).CheckAmt)
    Print #RHandle, "     Charge: "; Using("$##,###,###.##", 0)
  Else
    Print #RHandle, "      Check: "; Using("$##,###,###.##", 0)
    Print #RHandle, "     Charge: "; Using("$##,###,###.##", DCPaymentRec(1).CheckAmt)
  End If
  Print #RHandle, " Total Owed: "; Using("$##,###,###.##", DCPaymentRec(1).Amount)
  Print #RHandle, " Total Paid: "; Using("$##,###,###.##", Round(DCPaymentRec(1).CashAmt + DCPaymentRec(1).CheckAmt))
  Print #RHandle, " Change Due: "; Using("$##,###,###.##", DCPaymentRec(1).Change)
  Print #RHandle, "Amt Applied: "; Using("$##,###,###.##", Round((DCPaymentRec(1).CashAmt + DCPaymentRec(1).CheckAmt) - DCPaymentRec(1).Change))
  Print #RHandle,
  Print #RHandle,
  Print #RHandle, " Vin/Desc: "; QPTrim$(DCPaymentRec(1).VinDesc)
  Print #RHandle, "     Make: "; QPTrim$(DCPaymentRec(1).makemodel)
  Print #RHandle, "State Tag: "; QPTrim$(DCPaymentRec(1).StateTag)
  Print #RHandle, "Decal Num: "; QPTrim$(DCPaymentRec(1).Sticker)
  Print #RHandle,
  Print #RHandle, "Operator: "; OperNum
  Print #RHandle, "Receipt#: "; Using("######", CmNum&)
  Print #RHandle,
  Print #RHandle, "       T H A N K   Y O U !"
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  If CntrlDef = 1 Then
    Print #RHandle, CutPaper$
  Else
    Print #RHandle,
    Print #RHandle,
    Print #RHandle,
  End If
  Close RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
  If CntrlDef = 1 Then
    fpcmdDrawer_Click
  End If
  'PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
10:
  DefPrinter = RecpPort '"LPT" + QPTrim$(Str$(RecpPort)) + ":"
20:
 ' MsgBox "Printer -" + DefPrinter, vbOKOnly

  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
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
Validationthing:
  If DCPaymentRec(1).TransTender = 2 Or DCPaymentRec(1).TransTender = 3 Then
   If RctValidate And PrnVali = True Then
     RHandle2 = FreeFile
     Open PayRecpName2$ For Output As RHandle2
     Print #RHandle2, Chr$(27); Chr$(&H63); Chr$(&H30); Chr$(&H4)
     Print #RHandle2, Chr$(13); Chr$(10)
     Print #RHandle2, Tab(12); TownName$
     Print #RHandle2, Tab(12); "Bank- "; BnkAcctNum$
     Print #RHandle2, Tab(12); "FOR DEPOSIT ONLY"
     Print #RHandle2, Tab(12); "Acct. No. "; DCPaymentRec(1).CustNumber
     Print #RHandle2, Tab(12); "Date: "; PostDate$
     Print #RHandle2, Tab(12); "Time: "; Time
     Print #RHandle2,
     Print #RHandle2, Chr$(12)
     Close RHandle2

     LPTHandle = FreeFile
     Open DefPrinter For Output As LPTHandle
     RptHandle2 = FreeFile
     Open PayRecpName2$ For Input As RptHandle2
     Do
       If frmPrint.cmdCancel = False Then
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
    DCLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(DCPaymentRec(1).CustNumber)
  End If
 End If

70:
If PrnRecp = True Then
 DCLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(DCPaymentRec(1).CustNumber)
 KillFile PayRecpName$
 KillFile PayFileName$
End If
80:
  Exit Sub
Cancel:
  If Err > 0 Then
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub
Private Sub GetVehList()
  Dim NumOfDCRecs As Long, DCFile As Integer, Num1 As Long, Num2 As Long
  Dim cnt As Long, dcnt As Long, Cust As String
  Dim Build As String * 80

  ReDim DCCustREc(1) As DCCustRecType
  If fpCustRecNo > 0 Then
    OpenDCCustFile NumOfDCRecs, DCFile
    Get DCFile, fpCustRecNo, DCCustREc(1)
    Close DCFile
    If DCCustREc(1).FirstCar <= 0 Then
      Num1 = 0
      Cust$ = QPTrim$(DCCustREc(1).BILLNAME)
    Else
      Num1 = DCCustREc(1).FirstCar
      Cust$ = QPTrim$(DCCustREc(1).BILLNAME)
    End If
  
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  ReDim DCVRec(1) As DCVehType
  If Num1 > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    cnt = Num1
    frmVehDisplayList1.fpList1.AddItem "  Add New Vehicle" + Chr9$ + "0"
    Do Until cnt = 0
    'For cnt = Num1 To Num2
    Get DCvFile, cnt, DCVRec(1)
    If DCVRec(1).Active = "Y" Then
      LSet Build$ = QPTrim$(DCVRec(1).makemodel)
      Mid$(Build$, 30) = QPTrim$(DCVRec(1).StateTag)
      Mid$(Build$, 55) = QPTrim$(DCVRec(1).Desc)
      Mid$(Build$, 75) = Chr9$ + Str$(cnt)
      frmVehDisplayList1.fpList1.AddItem Build$
      dcnt = dcnt + 1
    End If
      cnt = DCVRec(1).NextRec
    Loop 'Next
    Close DCvFile
    frmVehDisplayList1.fpList1.ListIndex = 0
    frmVehDisplayList1.Caption = "Vehicle List - " & Cust$
    frmVehDisplayList1.Show 1
  Else
    frmVehDisplayList1.fpList1.AddItem "  Add New Vehicle" + Chr9$ + "0"
  End If
    If dcnt <= 0 Then
      frmVehDisplayList1.fpList1.ListIndex = 0
      frmVehDisplayList1.Caption = "Vehicle List - " & Cust$
      frmVehDisplayList1.Show 1
    End If
  End If
End Sub
Private Sub SaveTransaction()
  Dim ListFile As Integer, DCvFile As Integer, VehRecord As Long
  Dim PayFileName As String, DCPayRecLen As Integer, NumOfVRecs As Long
  Dim NumOfRecs As Long, DCVehReclen As Integer, NumOfDCRecs As Long
  Dim cnt As Integer, DCFile As Integer, PRevRec As Long, TAmtPaid As Double
  Dim CMTrRecLen As Integer, CHandle As Integer, Trnum2save As Long
  ReDim DCPaymentRec(1) As DCEditPaymentRecType
  ReDim DCVRec(1) As DCVehType
  DCPayRecLen = Len(DCPaymentRec(1))
  Oper$ = QPTrim$(lblOperator.Caption)
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  GoSub SaveVehicle
  If NewVeh = True Then
    'Must Update Account With This Vehicle
    GoSub AddNewVehicle
  End If
  DCPaymentRec(1).VehRecord = fpVehRecNo
  DCPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  DCPaymentRec(1).TranDate = Date2Num(txtPaymentDate)
  DCPaymentRec(1).CustNumber = fpCustRecNo
  DCPaymentRec(1).CustName = QPTrim(fptxtName.Caption)
  DCPaymentRec(1).CustAddr = QPTrim(fptxtAddress.Caption)
  DCPaymentRec(1).Amount = fpTAmtOwed
  TAmtPaid# = fpTAmtOwed
  Select Case fpcboTenderType.ListIndex
    Case 0:
      DCPaymentRec(1).TransTender = 1
    Case 1:
      DCPaymentRec(1).TransTender = 2
    Case 2:
      DCPaymentRec(1).TransTender = 3
    Case 3:
      DCPaymentRec(1).TransTender = 4
  End Select
  DCPaymentRec(1).CashAmt = fpCashAmt
  DCPaymentRec(1).CheckAmt = fpChkAmt
  DCPaymentRec(1).Amount = fpFee
  DCPaymentRec(1).Change = fpChangeDue
  DCPaymentRec(1).PayDesc = QPTrim(fptxtDesc)
  fpDecalCat.col = 1
  DCPaymentRec(1).DecalCat = QPTrim$(fpDecalCat.ColText)
  DCPaymentRec(1).VinDesc = QPTrim$(fpVinDesc)
  DCPaymentRec(1).makemodel = QPTrim$(fpMakeModl)
  DCPaymentRec(1).StateTag = QPTrim$(fpStateLic)
  DCPaymentRec(1).ExpDate = Date2Num%(fpExpireDate)
  DCPaymentRec(1).Sticker = QPTrim$(fpstickernum)
  DCPaymentRec(1).resident = QPTrim$(fpResident)
  DCPaymentRec(1).Owner = QPTrim$(fpOwner)
  DCPaymentRec(1).PersBuss = QPTrim$(fpBusPers.Text)
  DCPaymentRec(1).Notes = QPTrim$(fpNotes)
  DCPaymentRec(1).VoidFlag = 0
  VehRecord = fpVehRecNo
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = DCPayRecLen
  Put #ListFile, 1, DCPaymentRec(1)
  Close ListFile
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Shared As DCvFile Len = DCVehReclen
  Get DCvFile, VehRecord, DCVRec(1)
  DCVRec(1).PBFlag = QPTrim$(fpBusPers.Text)
  Put DCvFile, VehRecord, DCVRec(1)
  Close DCvFile

  DCLog "Oper:" + Oper$ + " Updated Paylist for Account:" + Str$(DCPaymentRec(1).CustNumber)
  'ClearScn
  Close
  GoSub Dothesave
Exit Sub

Dothesave:
'*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*
  
  Dim DCEditRecLen As Integer, DCEdFile As Integer, DCTransRecLen As Integer
  Dim DCTransFile As Integer, NumOfTransRecs As Long, NextTransRec As Long
  Dim Prev As Long
   ReDim EditPaymentRec(1) As DCEditPaymentRecType
 
  ReDim DCCustREc(1) As DCCustRecType
  OpenDCCustFile NumOfDCRecs, DCFile

  DCEditRecLen = Len(EditPaymentRec(1))
  DCEdFile = FreeFile
  Open PayFileName$ For Random Access Read Write Shared As DCEdFile Len = DCEditRecLen
  NumOfDCRecs = LOF(DCEdFile) \ DCEditRecLen
  
  ReDim DCVRec(1) As DCVehType
  
  ReDim DCTransRec(1) As DCTransRecType
  DCTransRecLen = Len(DCTransRec(1))
  DCTransFile = FreeFile
  Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
  NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
  NextTransRec = NumOfTransRecs + 1


    Get DCEdFile, 1, EditPaymentRec(1)

    If EditPaymentRec(1).Amount >= 0 And Val(EditPaymentRec(1).CustNumber) > 0 Then

      'GoSub OldVehPost
      GoSub UpdateVehRecord

      Get DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
      ' Post Charge First to Offset Payment of Decal
      DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
      DCTransRec(1).TransType = 1               ' Type 1 = Charge
      DCTransRec(1).TRVinDesc = EditPaymentRec(1).VinDesc
      DCTransRec(1).TransTender = EditPaymentRec(1).TransTender
      DCTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EditPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCCustREc(1).AcctBal + EditPaymentRec(1).Amount
      DCTransRec(1).makemodel = EditPaymentRec(1).makemodel
      DCTransRec(1).StateTag = EditPaymentRec(1).StateTag
      DCTransRec(1).Sticker = EditPaymentRec(1).Sticker
      DCTransRec(1).ExpireDate = EditPaymentRec(1).ExpDate
      DCTransRec(1).OperNum = EditPaymentRec(1).OperNum
      If Len(QPTrim$(EditPaymentRec(1).PayDesc)) > 0 Then
        DCTransRec(1).ExtraDesc = "DC-" + EditPaymentRec(1).PayDesc
      Else
        DCTransRec(1).ExtraDesc = "DC-Purchase"
      End If
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      DCTransRec(1).GLInterfaced = "Y"
      DCTransRec(1).DecalCat = EditPaymentRec(1).DecalCat
      DCTransRec(1).ChkByte = Chr$(1)
      DCTransRec(1).VoidFlag = "N"
      DCTransRec(1).VehRecord = EditPaymentRec(1).VehRecord
      Put DCTransFile, NextTransRec, DCTransRec(1)

      Get DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
      DCCustREc(1).AcctBal = DCCustREc(1).AcctBal + EditPaymentRec(1).Amount
      Put DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
      If DCCustREc(1).FirstTrans = 0 Then
        DCCustREc(1).FirstTrans = NextTransRec
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
      Else
        Prev = DCCustREc(1).LastTrans
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      NextTransRec = NextTransRec + 1

      ' Post Transaction Record First
      DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
      DCTransRec(1).TransType = 2               ' Type 2 = Payment
      DCTransRec(1).TRVinDesc = EditPaymentRec(1).VinDesc
      DCTransRec(1).TransTender = EditPaymentRec(1).TransTender
      DCTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EditPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCCustREc(1).AcctBal - EditPaymentRec(1).Amount
      DCTransRec(1).makemodel = EditPaymentRec(1).makemodel
      DCTransRec(1).StateTag = EditPaymentRec(1).StateTag
      DCTransRec(1).Sticker = EditPaymentRec(1).Sticker
      DCTransRec(1).ExpireDate = EditPaymentRec(1).ExpDate
      DCTransRec(1).OperNum = EditPaymentRec(1).OperNum
      If Len(QPTrim$(EditPaymentRec(1).PayDesc)) > 0 Then
        DCTransRec(1).ExtraDesc = "CM-DC" + EditPaymentRec(1).PayDesc
      Else
        DCTransRec(1).ExtraDesc = "CM-DCPayment"
      End If
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      DCTransRec(1).GLInterfaced = "N"
      DCTransRec(1).DecalCat = EditPaymentRec(1).DecalCat
      DCTransRec(1).ChkByte = Chr$(1)
      DCTransRec(1).VoidFlag = "N"
      DCTransRec(1).VehRecord = EditPaymentRec(1).VehRecord
      Put DCTransFile, NextTransRec, DCTransRec(1)
      Trnum2save = NextTransRec
      Get DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
      DCCustREc(1).AcctBal = DCCustREc(1).AcctBal - EditPaymentRec(1).Amount
      DCCustREc(1).LICENSE = EditPaymentRec(1).Sticker
      Put DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)

      If DCCustREc(1).FirstTrans = 0 Then
        DCCustREc(1).FirstTrans = NextTransRec
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
      Else
        Prev = DCCustREc(1).LastTrans
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(EditPaymentRec(1).CustNumber), DCCustREc(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      NextTransRec = NextTransRec + 1
    End If
  
  
  'Close
  
  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CMTrRec(1).TransDate = DCTransRec(1).TransDate
  CMTrRec(1).TransAmount = DCTransRec(1).TransAmount 'UBTransRec(1).CashAmount + UBTransRec(1).CheckAmount
  CMTrRec(1).TransCash = DCTransRec(1).CashAmount
  CMTrRec(1).TransAmtOwed = TAmtPaid#
  CMTrRec(1).TransCheck = DCTransRec(1).ChkAmount
  CMTrRec(1).TransDesc = "St#" + QPTrim$(fpstickernum) + QPTrim$(fptxtDesc)
  CMTrRec(1).TransSource = 151
  CMTrRec(1).TransName = QPTrim(fptxtName.Caption)
  CMTrRec(1).TransAcctNum = CustAcct&
  CMTrRec(1).TransDetNum = Trnum2save
  CMTrRec(1).TransOperNum = OperNum
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransTender = fpcboTenderType.ListIndex + 1
  CMTrRec(1).ChkByte = Chr$(1)
  CMTrRec(1).TransRevAmt(1) = CDbl(EditPaymentRec(1).VehRecord)
  fpDecalCat.col = 0  'this is the cat rec num
  CMTrRec(1).TransRevAmt(2) = CDbl(fpDecalCat.ColText)
  
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Shared As CHandle Len = CMTrRecLen
  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  CmNum = (LOF(CHandle) / CMTrRecLen) ' + 1
  Close CHandle
  CMLog "SAVED PAYMENT in CM ACCT:" + Str$(CustAcct&) + " Oper -" + Oper$
  DCLog "SAVED PAYMENT in CM ACCT:" + Str$(CustAcct&) + " Oper -" + Oper$

  
  ' Show All Posted
 
  'MsgBox "Save Complete", vbOKOnly, "Complete"
  Close
  Exit Sub

UpdateVehRecord:
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  VehRecord = EditPaymentRec(1).VehRecord
  If VehRecord <= 0 Or VehRecord > NumOfVRecs Then Close DCvFile: Return
  Get DCvFile, VehRecord, DCVRec(1)
  DCVRec(1).ExpireDate = EditPaymentRec(1).ExpDate
  DCVRec(1).Sticker = LTrim$(EditPaymentRec(1).Sticker)
  DCVRec(1).valid = "Y"
  DCVRec(1).Fee = EditPaymentRec(1).Amount
  Put DCvFile, VehRecord, DCVRec(1)
  Close DCvFile
Return
SaveVehicle:
  GoSub OpenVehFile
  GoSub SaveVRecord
Return
'************************* NEW VEHICLE ENTRY
AddNewVehicle:
  GoSub UpdateVendorPointer
Return


OpenVehFile:
  'Open Vehicle File
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  If fpVehRecNo > 0 Then
    VehRecord = fpVehRecNo
    Get DCvFile, VehRecord, DCVRec(1)
  Else
    VehRecord = NumOfVRecs + 1
    fpVehRecNo = VehRecord
    DCVRec(1).NextRec = 0
  End If
Return

SaveVRecord:
  fpDecalCat.col = 1
  DCVRec(1).DecalCat = QPTrim$(fpDecalCat.ColText)
  DCVRec(1).Fee = fpFee.Value
  DCVRec(1).makemodel = QPTrim$(fpMakeModl)
  DCVRec(1).StateTag = QPTrim$(fpStateLic)
  DCVRec(1).ExpireDate = Date2Num%(fpExpireDate)
  DCVRec(1).Sticker = QPTrim$(fpstickernum)
  DCVRec(1).valid = "Y"
  DCVRec(1).Active = "Y"
  DCVRec(1).Desc = QPTrim$(fpVinDesc)
  DCVRec(1).Notes = QPTrim$(fpNotes)
  DCVRec(1).MoreRoom = ""
  DCVRec(1).PBFlag = QPTrim$(fpBusPers.Text)
  DCVRec(1).MasterRecord = fpCustRecNo
  Put DCvFile, VehRecord, DCVRec(1)
Return
UpdateVendorPointer:
  ReDim DCCustREc(1) As DCCustRecType
  If fpCustRecNo > 0 Then
    OpenDCCustFile NumOfDCRecs, DCFile
    Get DCFile, fpCustRecNo, DCCustREc(1)
    If DCCustREc(1).FirstCar = 0 Then
      DCCustREc(1).FirstCar = VehRecord
      DCCustREc(1).LastCar = VehRecord
      Put DCFile, fpCustRecNo, DCCustREc(1)
    Else
      PRevRec = DCCustREc(1).LastCar
      DCCustREc(1).LastCar = VehRecord
      Put DCFile, fpCustRecNo, DCCustREc(1)
      Get DCvFile, PRevRec, DCVRec(1)
      DCVRec(1).NextRec = VehRecord
      Put DCvFile, PRevRec, DCVRec(1)
    End If
    Close DCvFile
  End If
Return
End Sub
