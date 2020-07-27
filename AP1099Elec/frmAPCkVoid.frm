VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAPCkVoid 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Void A/P Check"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmAPCkVoid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboVendor 
      Height          =   384
      Left            =   4272
      TabIndex        =   0
      Top             =   2304
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmAPCkVoid.frx":08CA
   End
   Begin LpLib.fpList fplstCks 
      Height          =   2208
      Left            =   4380
      TabIndex        =   1
      Top             =   3192
      Width           =   4932
      _Version        =   196608
      _ExtentX        =   8700
      _ExtentY        =   3895
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
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
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
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
      ColumnHeaderShow=   0   'False
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
      ColDesigner     =   "frmAPCkVoid.frx":0CB5
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   4320
      TabIndex        =   3
      Top             =   7536
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      ColDesigner     =   "frmAPCkVoid.frx":0FB1
   End
   Begin LpLib.fpCombo fpcboVoidDate 
      Height          =   384
      Left            =   4320
      TabIndex        =   2
      Top             =   6744
      Width           =   2892
      _Version        =   196608
      _ExtentX        =   5101
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
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmAPCkVoid.frx":134F
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "2:13 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/29/2007"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGo 
      Height          =   468
      Left            =   7968
      TabIndex        =   4
      Top             =   7440
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmAPCkVoid.frx":16E2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   468
      Left            =   9672
      TabIndex        =   5
      Top             =   7440
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmAPCkVoid.frx":18C0
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Void Date:"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   6816
      Width           =   1428
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
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
      Left            =   1752
      TabIndex        =   15
      Top             =   7584
      Width           =   2388
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   8112
      TabIndex        =   14
      Top             =   2976
      Width           =   876
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Check to Void, Press SpaceBar Or Arrow Key To Highlight."
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
      Height          =   588
      Left            =   2040
      TabIndex        =   13
      Top             =   5712
      Width           =   8148
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8208
      TabIndex        =   12
      Top             =   3432
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6240
      TabIndex        =   11
      Top             =   2976
      Width           =   876
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4272
      TabIndex        =   10
      Top             =   2976
      Width           =   1548
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check:"
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
      Left            =   2016
      TabIndex        =   9
      Top             =   3000
      Width           =   2028
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2676
      Left            =   4224
      Top             =   2928
      Width           =   5220
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4356
      Left            =   1920
      Top             =   2016
      Width           =   8388
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2532
      Top             =   984
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Void A/P Check"
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
      Left            =   3636
      TabIndex        =   7
      Top             =   1224
      Width           =   4836
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   2976
      TabIndex        =   6
      Top             =   2352
      Width           =   1092
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2544
      Top             =   864
      Width           =   7020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmAPCkVoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim LPDate As Integer, HPDate As Integer
Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmAPChkProcessMenu.Show
  Unload frmAPCkVoid
End Sub

Private Sub cmdGo_Click()
  Dim Pcnt As Integer, ChkVendor As String, cnt As Integer
  Dim VoidTransRecNum As Long, UseChkDate As Boolean
  Dim VoidChkNum As Long, VoidVendorRecNum As Long, TmpDate As Integer
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  
  cnt = 0
  If fplstCks.ListCount <> 0 Then
  For Pcnt = 0 To fplstCks.ListCount - 1
    If fplstCks.Selected(Pcnt) Then
      cnt = 1
      fpcboVendor.col = 2
      VoidVendorRecNum& = QPTrim$(Val(fpcboVendor.ColText))
      If fpcboVoidDate.ListIndex = 0 Then
        UseChkDate = True
        TmpDate = DateDiff("d", "12/31/1979", QPTrim(Mid$(fplstCks.Text, 13, 10)))
        'TmpDate = DateDiff("d", "12/31/1979", QPTrim$(Val(fplstchecks.ColText)))
      Else
        UseChkDate = False
        TmpDate = DateDiff("d", "12/31/1979", Date$)
      End If
      If (TmpDate < LPDate) Or (TmpDate > HPDate) Then
        If MsgBox("This Date Is Not Within Allowable Posting Range.Ok to Continue or Cancel?", vbOKCancel, "Invalid Date") = vbCancel Then
          Call MainLog("APInvVoid Cancel - DateOutofRange, Canceled")
          Close
          Exit Sub
        End If
      End If
      fplstCks.Row = Pcnt
      VoidTransRecNum& = QPTrim(Val(Mid$(fplstCks.Text, 55, 10)))
      VoidChkNum& = QPTrim(Val(Mid$(fplstCks.Text, 1, 9)))
      If MsgBox("Are You Sure You Wish To Void This Check?", vbYesNo, "Void Check") = vbNo Then
        fplstCks.SetFocus
        fplstCks.Selected = False
        Exit Sub
      End If
      If rptopt = 1 Then
        VoidCkTrans VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&, UseChkDate
      ElseIf rptopt = 2 Then
        VoidCkTrans2 VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&, UseChkDate
      End If
'Instead of staying on screen exit whole deal
      cmdExit_Click
    End If
    If cnt = 1 Then Exit For
   Next
    If cnt = 0 Then
      MsgBox "You Must Select A Check First.", vbOKOnly, "Select Check"
      fplstCks.SetFocus
    End If

  Else
    MsgBox "No Checks to Select, Try Another Vendor.", vbOKOnly, "No Checks"
    fpcboVendor.SetFocus
 
   End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%G"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile ("APCHK.opn")
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpVoidPostedAPChk
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetPostDates LPDate, HPDate
  StatusBar1.Panels.Item(1).Text = GLUserName
  VendCodeName fpcboVendor
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  fpcboVoidDate.InsertRow = "Check Date"
  fpcboVoidDate.InsertRow = "Current Date"
  fpcboVoidDate.ListIndex = 0
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub fpcboVoidDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVoidDate.ListDown = True
  End If
  If fpcboVoidDate.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fplstCks.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fplstCks_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
   If fplstCks.Selected = True Then
    fplstCks.Selected = False
   Else
    fplstCks.Selected = True
   End If
  End If
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdGo.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboVoidDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboVendor_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendor.ListDown = True
  End If
  If fpcboVendor.ListDown <> True Then
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

Private Sub fpcboVendor_LostFocus()
  fpcboVendor.Action = ActionClearSearchBuffer
End Sub


Private Sub fplstCks_GotFocus()
  If fplstCks.ListCount = 0 Then
    MsgBox "No Checks For This Vendor, Please Select Another.", vbOKOnly, "No PO's"
    fpcboVendor.SetFocus
  End If
End Sub

Private Sub fpcboVendor_Click()
  fplstCks.Clear
  If fpcboVendor.ListIndex <> -1 Then
    LoadUp
  End If
End Sub
Private Sub LoadUp()
  Dim VendorFile As Integer, NumVRecs As Integer, VRecNum As Integer
  Dim Last As Integer, cnt As Integer, Dcnt As Integer, TmpAcct As Integer
  fpcboVendor.col = 2
  VRecNum = fpcboVendor.ColText
  If VRecNum > 0 Then
    OpenVendorFile VendorFile, NumVRecs
    Get VendorFile, VRecNum, Vendor
    Findchk VRecNum
    End If
 Close
End Sub
Private Sub Findchk(vrec As Integer)
  Dim CKCnt As Integer, NextTrans As Long, fmt As String
  Dim VendorFile As Integer, NumVRecs As Integer, tempstr As String
  Dim APLedgerFile As Integer, NumTrans As Long, LdRecLen As Integer
  ReDim APLedgerRec(1) As APLedger81RecType
  LdRecLen = Len(APLedgerRec(1))

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  fmt = "$ ###,###,###.##"
  Get VendorFile, vrec, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 3 Then
      CKCnt = CKCnt + 1
    End If
    NextTrans& = APLedgerRec(1).NextTrans

  Loop

  If CKCnt <> 0 Then

  NextTrans& = Vendor.FrstTran
  fplstCks.Clear
  fplstCks.Row = 0
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 3 Then
      tempstr = Space$(60)
      Mid$(tempstr, 1) = QPTrim$(APLedgerRec(1).DOCNum)
      Mid$(tempstr, 13) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(tempstr, 23) = Using(fmt, Str$(APLedgerRec(1).amt))
      Mid$(tempstr, 55) = NextTrans&
      fplstCks.InsertRow = tempstr
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  End If
  Close
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

'  If InStr(User$, "WRIGHTSVILLE") > 0 Then
'    WBFlag = True
'  End If

'      If Not WBFlag Then
'        QPrintRC Space$(25), 13, 18, 0
'        Fld(3).Protected = True
'      End If

Private Sub VoidCkTrans(VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&, UseChkDate)
  Dim Interfacefile As String, PRNfileName As String, PrnTitle As String
  Dim fmt As String, PRNFile As Integer, GLIFRecLen As Integer, LdRecLen As Integer
  Dim DistRecLEn As Integer, APLedgerFile As Integer, NumTran As Long
  Dim VendorFile As Integer, NumVRecs As Integer, NextTrans As Long
  Dim VoidDate As String, GLIFFile As Integer, NextDist As Long
  Dim APDistFile As Integer, NumDistRecs As Long, DistCnt As Integer
  Dim APAcct As String, CashAcct As String, CDCash As String, CDDue As String
  Dim CDActive As String, GTDebit As Double, GTCredit As Double
  Dim BadAcct As Integer, GLLogFileName As Integer, ReportFile As String
  Dim NoInv As Boolean, Invcnt As Integer, TempCash As String, LG As String
  Dim ToPrint As String, tempfund As String, PadChars As Integer
  Dim DetPad As String
  Interfacefile$ = "APVOIDIF.DAT"
  If Exist("APVOIDIF.DAT") Then
    MsgBox "A Void Check Interface File Already Exists, Please Contact Software Support! Do NOT Try To Void Again!", vbOKOnly, "Error"
    Call MainLog("Void APChk Interface Error.")
    Exit Sub
  End If
  GetAPAcct APAcct
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  
  Invcnt = 0
  PRNfileName$ = "VOIDCHK.PRN"
  PrnTitle$ = "Voided Check Report"
  fmt$ = "###,###.##"
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim DistRec(1) As APDistRecType
  ReDim GLifRec(1) As GLTransRecType

  GLIFRecLen = Len(GLifRec(1))
  LdRecLen = Len(APLedgerRec(1))
  DistRecLEn = Len(DistRec(1))

  PRNFile = FreeFile
  Open PRNfileName$ For Output As PRNFile

  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  OpenVendorFile VendorFile, NumVRecs
  Get VendorFile, VoidVendorRecNum&, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 3 And APLedgerRec(1).PDCheckNum = VoidChkNum& Then
      If UseChkDate Then
        VoidDate$ = Format(DateAdd("d", (APLedgerRec(1).PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
      Else
        VoidDate$ = Date$
      End If
      NoInv = False
      Invcnt = Invcnt + 1
      APLedgerRec(1).PAYCODE = 1
      APLedgerRec(1).PDCheckNum = 0
      APLedgerRec(1).PDCheckDate = 0
      APLedgerRec(1).Bankcode = 0
    'If no errors first go thru will write record further down in code......
      'Put APLedgerFile, NextTrans&, APLedgerRec(1)
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  Get APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  'APLedgerRec(1).TrCode = -3
  'Put APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  'Close APLedgerFile
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  'If used bank code need correct acct for void
  '--If we are not using a central depository, then use the
  '--cash account$ assigned with the record (glbank.dat)
  If CDActive$ <> "Y" Then
    If APLedgerRec(1).Bankcode <> 0 Then
      If Exist("GLBANK.DAT") Then
        TempCash$ = GetBankGLAcct(APLedgerRec(1).Bankcode)
      End If
    End If
  End If
  If TempCash$ <> "" Then
    CashAcct$ = TempCash$
  End If
  If VoidDate$ = "" Then
    VoidDate$ = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
    NoInv = True
  End If
  If CDActive$ = "Y" Then
    PadChars = GLDetLen - GLFundLen
    If PadChars > 0 Then
      DetPad$ = String$(PadChars, "0")
    End If
  End If

  GLIFFile = FreeFile
  Open Interfacefile$ For Random As GLIFFile Len = GLIFRecLen

  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn

  NextDist& = APLedgerRec(1).FrstDist
  DistCnt = 0
  Do Until NextDist& = 0
    Get APDistFile, NextDist&, DistRec(1)

    '--Set common values
    GLifRec(1).Src = "VC" + Format(Date$, "mmddyy")
    GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", VoidDate$)
    GLifRec(1).Desc = "TO VOID CHK" + Str$(VoidChkNum&)
    GLifRec(1).LDesc = APLedgerRec(1).VendorCode
    GLifRec(1).Ref = Str$(VoidChkNum&)
    LG$ = "VOID Check " + Str$(VoidChkNum&) + VoidDate$
 '--make a debit
    tempfund$ = Left$(DistRec(1).DistAcctNum, GLFundLen)
    GLifRec(1).AcctNum = Left$(DistRec(1).DistAcctNum, GLFundLen) + CashAcct$
    GLifRec(1).CrAmt = 0
    GLifRec(1).DrAmt = DistRec(1).DistAmt
    Put GLIFFile, , GLifRec(1)
    GTDebit# = GTDebit# + GLifRec(1).DrAmt
    GoSub PrintGLIF
 
 '--Make a credit
    'GLifRec(1).AcctNum = Left$(DistRec(1).DistAcctNum, GLFundLen) + APAcct$
'Changed this for correct acct ps-7/30-02
    GLifRec(1).AcctNum = QPTrim(DistRec(1).DistAcctNum)
    GLifRec(1).CrAmt = DistRec(1).DistAmt
    GLifRec(1).DrAmt = 0
    Put GLIFFile, , GLifRec(1)
    GTCredit# = GTCredit# + GLifRec(1).CrAmt
    GoSub PrintGLIF
    
    If CDActive$ = "Y" Then
        '--Credit Central Depository Due From
        GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", VoidDate$)
        GLifRec(1).Desc = "TO VOID CHK" + Str$(VoidChkNum&)
        GLifRec(1).LDesc = APLedgerRec(1).VendorCode
        GLifRec(1).Src = "VC" + Format(Date$, "mmddyy")
        GLifRec(1).Ref = Str$(VoidChkNum&)
        GLifRec(1).AcctNum = CDDue$ + tempfund$ + DetPad$
        GLifRec(1).DrAmt = 0
        GLifRec(1).CrAmt = DistRec(1).DistAmt
        Put GLIFFile, , GLifRec(1)
        GoSub PrintGLIF
    End If
    NextDist& = DistRec(1).NextDist

  Loop
'''################
      '--IF Cent Dep
      If CDActive$ = "Y" Then
      
        '--Debit Central Depository Cash
          GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", VoidDate$)
          GLifRec(1).AcctNum = CDCash$
          GLifRec(1).Desc = "TO VOID CHK" + Str$(VoidChkNum&)
          GLifRec(1).LDesc = APLedgerRec(1).VendorCode
          GLifRec(1).DrAmt = GTDebit#
          GLifRec(1).CrAmt = 0
          GLifRec(1).Src = "VC" + Format(Date$, "mmddyy")
          GLifRec(1).Ref = Str$(VoidChkNum&)
          Put GLIFFile, , GLifRec(1)
          GoSub PrintGLIF
      End If
'''''#################

'  Print #PRNFile,
'  Print #PRNFile, "Grand Total";
'  Print #PRNFile, Tab(58); Using(fmt$, GTDebit#);
'  Print #PRNFile, Tab(70); Using(fmt$, GTCredit#)
  If NoInv = True Then
    Call MainLog("AP Chk Void Error - No Invoice.")
    MsgBox "Error:Invalid Invoice Transaction- An Open Invoice was not created for this Check.", vbOKOnly, "Error"
  End If
'  Print #PRNFile, Chr$(12)
  Close

  Post2GL Interfacefile$, BadAcct, frmAPCkVoid, False
  If BadAcct Then
    Call MainLog("AP Chk Void Error No Accts Updated.")
    MsgBox "Void Error.  Review GL Utility Log for details.", vbOKOnly, "Void aborted"
    ARptErrorLog.GetName PRNfileName$
    ARptErrorLog.startrpt

    'ViewPrint PRNfileName$, PrnTitle$
    MsgBox "Accounts were not updated.", vbOKOnly, "Void Canceled"
    Exit Sub
  End If
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  OpenVendorFile VendorFile, NumVRecs
  Get VendorFile, VoidVendorRecNum&, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
  'Update invoice ledger record
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 3 And APLedgerRec(1).PDCheckNum = VoidChkNum& Then
      APLedgerRec(1).PAYCODE = 1
      APLedgerRec(1).PDCheckNum = 0
      APLedgerRec(1).PDCheckDate = 0
      APLedgerRec(1).Bankcode = 0
      Put APLedgerFile, NextTrans&, APLedgerRec(1)
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  'Update Check ledger record
  Get APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  APLedgerRec(1).TRCode = -3
  Put APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  Close APLedgerFile

  Post2GL Interfacefile$, BadAcct, frmAPCkVoid, True
  If BadAcct <> 0 Then
    Call MainLog("AP chk Void Err -Not Complete.")
    MsgBox "Errors Were Found. Review GL Posting Log.", vbOKOnly, "GL Account Error"
    GLLogFileName = "GLlog.dat"
    ReportFile$ = "GLlog.dat"
    ARptErrorLog.GetName ReportFile$
    ARptErrorLog.startrpt

    'ViewPrint ReportFile$, "Posting Log"
  End If
  ARptVoidEntries.totCred = Using(fmt$, GTCredit#)
  ARptVoidEntries.totDeb = Using(fmt$, GTDebit#)
  ARptVoidEntries.Label1.Caption = PrnTitle$
  ARptVoidEntries.txtDate.Caption = Date$
  ARptVoidEntries.txtTown.Caption = GLUserName$
  ARptVoidEntries.GetName PRNfileName$
  ARptVoidEntries.startrpt
 ' ViewPrint PRNfileName$, PrnTitle$
  KillFile Interfacefile$
  Call MainLog(LG$ + " Completed.")
  MsgBox "Void Procedure Completed.", vbOKOnly, "Void Complete"
  If Invcnt = 0 Then
    frmVoidCkMsg.Label1.Visible = False
    frmVoidCkMsg.Label2.Visible = False
    frmVoidCkMsg.Label3.Visible = False
    frmVoidCkMsg.Label4 = "An invalid invoice transaction was discovered for this check. This was noted on the previous report, also there will not be an open invoice for this check. If you have questions, contact software support."
    frmVoidCkMsg.Label4.Visible = True
  End If
  frmVoidCkMsg.Show 1

Exit Sub
PrintGLIF:
  ToPrint$ = GLifRec(1).AcctNum + "~" + Format(DateAdd("d", (GLifRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = ToPrint$ + "~" + Left$(GLifRec(1).Desc, 17) + "~" + GLifRec(1).Ref
  ToPrint$ = ToPrint$ + "~" + Using(fmt$, GLifRec(1).DrAmt) + "~" + Using(fmt$, GLifRec(1).CrAmt)
  Print #PRNFile, ToPrint$
Return
End Sub

Private Sub VoidCkTrans2(VoidVendorRecNum&, VoidTransRecNum&, VoidChkNum&, UseChkDate)
  Dim Interfacefile As String, PRNfileName As String, PrnTitle As String
  Dim fmt As String, PRNFile As Integer, GLIFRecLen As Integer, LdRecLen As Integer
  Dim DistRecLEn As Integer, APLedgerFile As Integer, NumTran As Long
  Dim VendorFile As Integer, NumVRecs As Integer, NextTrans As Long
  Dim VoidDate As String, GLIFFile As Integer, NextDist As Long
  Dim APDistFile As Integer, NumDistRecs As Long, DistCnt As Integer
  Dim APAcct As String, CashAcct As String, CDCash As String, CDDue As String
  Dim CDActive As String, GTDebit As Double, GTCredit As Double
  Dim BadAcct As Integer, GLLogFileName As Integer, ReportFile As String
  Dim NoInv As Boolean, Invcnt As Integer, TempCash As String, LG As String
  Dim tempfund As String, PadChars As Integer, DetPad As String
  Interfacefile$ = "APVOIDIF.DAT"
  If Exist("APVOIDIF.DAT") Then
    MsgBox "A Void Check Interface File Already Exists, Please Contact Software Support! Do NOT Try To Void Again!", vbOKOnly, "Error"
    Call MainLog("Void APChk Interface Error.")
    Exit Sub
  End If
  GetAPAcct APAcct
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  
  Invcnt = 0
  PRNfileName$ = "VOIDCHK.PRN"
  PrnTitle$ = "Voided Check Report"
  fmt$ = "###,###.##"
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim DistRec(1) As APDistRecType
  ReDim GLifRec(1) As GLTransRecType

  GLIFRecLen = Len(GLifRec(1))
  LdRecLen = Len(APLedgerRec(1))
  DistRecLEn = Len(DistRec(1))

  PRNFile = FreeFile
  Open PRNfileName$ For Output As PRNFile
  GoSub PrintReportHeader

  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  OpenVendorFile VendorFile, NumVRecs
  Get VendorFile, VoidVendorRecNum&, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 3 And APLedgerRec(1).PDCheckNum = VoidChkNum& Then
      If UseChkDate Then
        VoidDate$ = Format(DateAdd("d", (APLedgerRec(1).PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
      Else
        VoidDate$ = Date$
      End If
      NoInv = False
      Invcnt = Invcnt + 1
      APLedgerRec(1).PAYCODE = 1
      APLedgerRec(1).PDCheckNum = 0
      APLedgerRec(1).PDCheckDate = 0
      APLedgerRec(1).Bankcode = 0
    'If no errors first go thru will write record further down in code......
      'Put APLedgerFile, NextTrans&, APLedgerRec(1)
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  Get APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  'APLedgerRec(1).TrCode = -3
  'Put APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  'Close APLedgerFile
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  'If used bank code need correct acct for void
  '--If we are not using a central depository, then use the
  '--cash account$ assigned with the record (glbank.dat)
  If CDActive$ <> "Y" Then
    If APLedgerRec(1).Bankcode <> 0 Then
      If Exist("GLBANK.DAT") Then
        TempCash$ = GetBankGLAcct(APLedgerRec(1).Bankcode)
      End If
    End If
  End If
  If TempCash$ <> "" Then
    CashAcct$ = TempCash$
  End If

  If VoidDate$ = "" Then
    VoidDate$ = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
    NoInv = True
  End If
  If CDActive$ = "Y" Then
    PadChars = GLDetLen - GLFundLen
    If PadChars > 0 Then
      DetPad$ = String$(PadChars, "0")
    End If
  End If

  GLIFFile = FreeFile
  Open Interfacefile$ For Random As GLIFFile Len = GLIFRecLen

  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn

  NextDist& = APLedgerRec(1).FrstDist
  DistCnt = 0
  Do Until NextDist& = 0
    Get APDistFile, NextDist&, DistRec(1)

    '--Set common values
    GLifRec(1).Src = "VC" + Format(Date$, "mmddyy")
    GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", VoidDate$)
    GLifRec(1).Desc = "TO VOID CHK" + Str$(VoidChkNum&)
    GLifRec(1).Ref = Str$(VoidChkNum&)
    LG$ = "VOID Check " + Str$(VoidChkNum&) + VoidDate$
 '--make a debit
    tempfund$ = Left$(DistRec(1).DistAcctNum, GLFundLen)
    GLifRec(1).AcctNum = Left$(DistRec(1).DistAcctNum, GLFundLen) + CashAcct$
    GLifRec(1).CrAmt = 0
    GLifRec(1).DrAmt = DistRec(1).DistAmt
    Put GLIFFile, , GLifRec(1)
    GTDebit# = GTDebit# + GLifRec(1).DrAmt
    GoSub PrintGLIF
 
 '--Make a credit
    'GLifRec(1).AcctNum = Left$(DistRec(1).DistAcctNum, GLFundLen) + APAcct$
'Changed this for correct acct ps-7/30-02
    GLifRec(1).AcctNum = QPTrim(DistRec(1).DistAcctNum)
    GLifRec(1).CrAmt = DistRec(1).DistAmt
    GLifRec(1).DrAmt = 0
    Put GLIFFile, , GLifRec(1)
    GTCredit# = GTCredit# + GLifRec(1).CrAmt
    GoSub PrintGLIF
    
    If CDActive$ = "Y" Then
        '--Credit Central Depository Due From
        GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", VoidDate$)
        GLifRec(1).Desc = "TO VOID CHK" + Str$(VoidChkNum&)
        GLifRec(1).LDesc = APLedgerRec(1).VendorCode
        GLifRec(1).Src = "VC" + Format(Date$, "mmddyy")
        GLifRec(1).Ref = Str$(VoidChkNum&)
        GLifRec(1).AcctNum = CDDue$ + tempfund$ + DetPad$
        GLifRec(1).DrAmt = 0
        GLifRec(1).CrAmt = DistRec(1).DistAmt
        Put GLIFFile, , GLifRec(1)
        GoSub PrintGLIF
    End If
    NextDist& = DistRec(1).NextDist

  Loop
'''################
      '--IF Cent Dep
      If CDActive$ = "Y" Then
      
        '--Debit Central Depository Cash
          GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", VoidDate$)
          GLifRec(1).AcctNum = CDCash$
          GLifRec(1).Desc = "TO VOID CHK" + Str$(VoidChkNum&)
          GLifRec(1).LDesc = APLedgerRec(1).VendorCode
          GLifRec(1).DrAmt = GTDebit#
          GLifRec(1).CrAmt = 0
          GLifRec(1).Src = "VC" + Format(Date$, "mmddyy")
          GLifRec(1).Ref = Str$(VoidChkNum&)
          Put GLIFFile, , GLifRec(1)
          GoSub PrintGLIF
      End If
'''''#################

  Print #PRNFile,
  Print #PRNFile, "Grand Total";
  Print #PRNFile, Tab(58); Using(fmt$, GTDebit#);
  Print #PRNFile, Tab(70); Using(fmt$, GTCredit#)
  If NoInv = True Then
    Call MainLog("AP Chk Void Error - No Invoice.")
    Print #PRNFile, "Error:Invalid Invoice Transaction- An Open Invoice was not created for this Check."
  End If
  Print #PRNFile, Chr$(12)
  Close

  Post2GL Interfacefile$, BadAcct, frmAPCkVoid, False
  If BadAcct Then
    Call MainLog("AP Chk Void Error No Accts Updated.")
    MsgBox "Void Error.  Review GL Utility Log for details.", vbOKOnly, "Void aborted"
    ViewPrint PRNfileName$, PrnTitle$
    MsgBox "Accounts were not updated.", vbOKOnly, "Void Canceled"
    Exit Sub
  End If
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  OpenVendorFile VendorFile, NumVRecs
  Get VendorFile, VoidVendorRecNum&, Vendor
  NextTrans& = Vendor.FrstTran
  Do Until NextTrans& = 0
  'Update invoice ledger record
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 3 And APLedgerRec(1).PDCheckNum = VoidChkNum& Then
      APLedgerRec(1).PAYCODE = 1
      APLedgerRec(1).PDCheckNum = 0
      APLedgerRec(1).PDCheckDate = 0
      APLedgerRec(1).Bankcode = 0
      Put APLedgerFile, NextTrans&, APLedgerRec(1)
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  'Update Check ledger record
  Get APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  APLedgerRec(1).TRCode = -3
  Put APLedgerFile, VoidTransRecNum&, APLedgerRec(1)
  Close APLedgerFile

  Post2GL Interfacefile$, BadAcct, frmAPCkVoid, True
  If BadAcct <> 0 Then
    Call MainLog("AP chk Void Err -Not Complete.")
    MsgBox "Errors Were Found. Review GL Posting Log.", vbOKOnly, "GL Account Error"
    GLLogFileName = "GLlog.dat"
    ReportFile$ = "GLlog.dat"
    ViewPrint ReportFile$, "Posting Log"
  End If
  

  ViewPrint PRNfileName$, PrnTitle$
  KillFile Interfacefile$
  Call MainLog(LG$ + " Completed.")
  MsgBox "Void Procedure Completed.", vbOKOnly, "Void Complete"
  If Invcnt = 0 Then
    frmVoidCkMsg.Label1.Visible = False
    frmVoidCkMsg.Label2.Visible = False
    frmVoidCkMsg.Label3.Visible = False
    frmVoidCkMsg.Label4 = "An invalid invoice transaction was discovered for this check. This was noted on the previous report, also there will not be an open invoice for this check. If you have questions, contact software support."
    frmVoidCkMsg.Label4.Visible = True
  End If
  frmVoidCkMsg.Show 1

Exit Sub
PrintReportHeader:
  Print #PRNFile, PrnTitle$
  Print #PRNFile, "Run Date: " + Date$
  Print #PRNFile,
  Print #PRNFile, "Acct No         Date        Description        Ref            Debit      Credit"
  Print #PRNFile, "==============================================================================="
  Return
PrintGLIF:
  Print #PRNFile, GLifRec(1).AcctNum;
  Print #PRNFile, Tab(17); Format(DateAdd("d", (GLifRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy");
  Print #PRNFile, Tab(29); Left$(GLifRec(1).Desc, 17);
  Print #PRNFile, Tab(48); GLifRec(1).Ref;
  Print #PRNFile, Tab(58); Using(fmt$, GLifRec(1).DrAmt);
  Print #PRNFile, Tab(70); Using(fmt$, GLifRec(1).CrAmt)
Return
End Sub


