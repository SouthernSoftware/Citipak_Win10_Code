VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmIFInitialize 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface Initialize"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   12192
   Icon            =   "frmIFInitialize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboCashMgmt 
      Height          =   384
      Left            =   6924
      TabIndex        =   0
      Top             =   4152
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmIFInitialize.frx":08CA
   End
   Begin LpLib.fpCombo fpcboUtilBill 
      Height          =   384
      Left            =   6924
      TabIndex        =   1
      Top             =   4644
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmIFInitialize.frx":0BC1
   End
   Begin LpLib.fpCombo fpcboPropTax 
      Height          =   384
      Left            =   6924
      TabIndex        =   2
      Top             =   5124
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmIFInitialize.frx":0EB8
   End
   Begin LpLib.fpCombo fpcboBusLic 
      Height          =   384
      Left            =   6924
      TabIndex        =   3
      Top             =   5616
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmIFInitialize.frx":11AF
   End
   Begin LpLib.fpCombo fpcboVehDec 
      Height          =   384
      Left            =   6924
      TabIndex        =   4
      Top             =   6096
      Width           =   1164
      _Version        =   196608
      _ExtentX        =   2053
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmIFInitialize.frx":14A6
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   9648
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7704
      Width           =   1356
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7968
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7704
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   7
      Top             =   8640
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4:32 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "3/3/2006"
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
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      X1              =   3648
      X2              =   9672
      Y1              =   3792
      Y2              =   3792
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Va.Vehicle Decals (Y/N):"
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
      Height          =   348
      Left            =   4164
      TabIndex        =   16
      Top             =   6096
      Width           =   2652
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Licenses (Y/N):"
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
      Height          =   348
      Left            =   3924
      TabIndex        =   15
      Top             =   5628
      Width           =   2892
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Property Taxes (Y/N):"
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
      Height          =   348
      Left            =   4260
      TabIndex        =   14
      Top             =   5148
      Width           =   2556
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Utility Billing (Y/N):"
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
      Height          =   348
      Left            =   4692
      TabIndex        =   13
      Top             =   4680
      Width           =   2124
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Management (Y/N):"
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
      Height          =   348
      Left            =   3948
      TabIndex        =   12
      Top             =   4200
      Width           =   2868
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00404040&
      Height          =   4788
      Left            =   2400
      Top             =   2112
      Width           =   7452
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   852
      Left            =   3288
      Top             =   912
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initialize Interface Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3720
      TabIndex        =   11
      Top             =   1152
      Width           =   4764
   End
   Begin VB.Label lblPosting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please make sure you wish to do this!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   372
      Left            =   3672
      TabIndex        =   10
      Top             =   3096
      Width           =   4908
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Function will set all Transactions as Posted to the General Ledger from each Module as Indicated below."
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
      Height          =   732
      Index           =   0
      Left            =   2616
      TabIndex        =   9
      Top             =   2424
      Width           =   7116
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modules "
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
      Index           =   1
      Left            =   2568
      TabIndex        =   8
      Top             =   3624
      Width           =   1188
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3288
      Top             =   792
      Width           =   5772
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
Attribute VB_Name = "frmIFInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GJEdit As TrEditRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

'Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmGetDistMenu.Show
  Unload frmIFInitialize
End Sub

Private Sub cmdOk_Click()
  DeActivateControls frmIFInitialize, True
  InitIF
  ActivateControls frmIFInitialize, True
  cmdExit_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"  'Arrow Down
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"   'arrow up
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"     'Esc key
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"     'alt O or f10
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpInitialize
  GetYN fpcboCashMgmt
  GetYN fpcboUtilBill
  GetYN fpcboPropTax
  GetYN fpcboBusLic
  GetYN fpcboVehDec
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub GetYN(x As fpCombo)
  x.AddItem "No"
  x.AddItem "Yes"
  x.ListIndex = 0
End Sub
Private Sub InitIF()
  Dim CMTransRecLen As Integer, CMTran As Integer, NumOfTRecs As Long
  Dim TCnt As Long, UBTransRecLen As Integer, UBTran As Integer
  Dim TaxTranRecLen As Integer, TranFile As Integer, NumOfRec As Long
  Dim cnt As Long, ARTransRecLen As Integer, ARTransFile As Integer
  Dim NumOfTransRecs As Long, DCTransRecLen As Integer
  Dim DCTransFile As Integer, Okcnt As Integer
  Okcnt = 0
  If fpcboCashMgmt.ListIndex = 1 Then
  If Exist("cmtrans.dat") Then
    FrmShowPctComp.Label1 = "INITIALIZING CASH TRANSACTIONS."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show
    ReDim CMTransRec(1) As CMTransRecType
    CMTransRecLen = Len(CMTransRec(1))
    CMTran = FreeFile
    Open "CMTRANS.DAT" For Random Shared As CMTran Len = CMTransRecLen
    NumOfTRecs& = LOF(CMTran) \ CMTransRecLen
    For TCnt& = 1 To NumOfTRecs& 'To 1 Step -1
      FrmShowPctComp.ShowPctComp TCnt&, NumOfTRecs&
      Get #CMTran, TCnt&, CMTransRec(1)
      CMTransRec(1).Trans2GL = "Y"
      Put #CMTran, TCnt&, CMTransRec(1)
    Next TCnt&
    Close CMTran
    Call MainLog("Init CM for IF.")
    Okcnt = Okcnt + 1
  End If
  End If
  If fpcboUtilBill.ListIndex = 1 Then
  If Exist("ubtrans.dat") Then
    FrmShowPctComp.Label1 = "INITIALIZING UTILITY TRANSACTIONS."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show
    ReDim UBTransRec(1) As UBTransRecType
    UBTransRecLen = Len(UBTransRec(1))
    UBTran = FreeFile
    Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
    NumOfTRecs& = LOF(UBTran) \ UBTransRecLen
    For TCnt& = 1 To NumOfTRecs& 'To 1 Step -1
      FrmShowPctComp.ShowPctComp TCnt&, NumOfTRecs&
      Get #UBTran, TCnt&, UBTransRec(1)
      UBTransRec(1).Posted2GL = "Y"
      Put #UBTran, TCnt&, UBTransRec(1)
    Next TCnt&
    Close UBTran
    Call MainLog("Init UB for IF.")
    Okcnt = Okcnt + 1
  End If
  End If
  If fpcboPropTax.ListIndex = 1 Then
    If Exist("CitiTaxes.EXE") Then
      If Exist("taxtrans.dat") Then
        FrmShowPctComp.Label1 = "INITIALIZING TAX TRANSACTIONS."
        FrmShowPctComp.cmdCancel.Enabled = False
        FrmShowPctComp.Show
        ReDim TaxTrans(1) As TaxTransactionType
        TaxTranRecLen = Len(TaxTrans(1))
        TranFile = FreeFile
        Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
        NumOfRec& = LOF(TranFile) \ TaxTranRecLen
        For cnt& = 1 To NumOfRec&
          FrmShowPctComp.ShowPctComp cnt&, NumOfRec&
          Get #TranFile, cnt&, TaxTrans(1)
          TaxTrans(1).Posted2GL = "Y"
          Put #TranFile, cnt&, TaxTrans(1)
        Next cnt&
        Close TranFile
        Call MainLog("Init TX for IF.")
        Okcnt = Okcnt + 1
      End If
    ElseIf Exist("VACitiTax.EXE") Then
      If Exist("taxtrans.dat") Then
        FrmShowPctComp.Label1 = "INITIALIZING TAX TRANSACTIONS."
        FrmShowPctComp.cmdCancel.Enabled = False
        FrmShowPctComp.Show
        ReDim TaxTransV(1) As TaxVATransactionType
        TaxTranRecLen = Len(TaxTransV(1))
        TranFile = FreeFile
        Open "TAXTRANS.DAT" For Random Shared As TranFile Len = TaxTranRecLen
        NumOfRec& = LOF(TranFile) \ TaxTranRecLen
        For cnt& = 1 To NumOfRec&
          FrmShowPctComp.ShowPctComp cnt&, NumOfRec&
          Get #TranFile, cnt&, TaxTransV(1)
          TaxTransV(1).Posted2GL = "Y"
          Put #TranFile, cnt&, TaxTransV(1)
        Next cnt&
        Close TranFile
        Call MainLog("Init VTX for IF.")
        Okcnt = Okcnt + 1
      End If
    End If
  End If
  If fpcboBusLic.ListIndex = 1 Then
  If Exist("artrans.dat") Then
    FrmShowPctComp.Label1 = "INITIALIZING BUS. LICENSE TRANSACTIONS."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show
    ReDim ARTransRec(1) As ARTransRecType
    ARTransRecLen = Len(ARTransRec(1))
    ARTransFile = FreeFile
    Open "ARTRANS.DAT" For Random Access Read Write Shared As ARTransFile Len = ARTransRecLen
    NumOfTransRecs& = LOF(ARTransFile) \ ARTransRecLen
    For TCnt& = 1 To NumOfTransRecs&
      FrmShowPctComp.ShowPctComp TCnt&, NumOfTransRecs&
      Get ARTransFile, TCnt&, ARTransRec(1)
      ARTransRec(1).Posted2GL = "Y"
      Put ARTransFile, TCnt&, ARTransRec(1)
    Next TCnt&
    Close ARTransFile
    Call MainLog("Init BL for IF.")
    Okcnt = Okcnt + 1
  End If
  End If
  If fpcboVehDec.ListIndex = 1 Then
  If Exist("dctrans.dat") Then
    FrmShowPctComp.Label1 = "INITIALIZING DECAL TRANSACTIONS."
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show
    ReDim DCTransRec(1) As DCTransRecType
    DCTransRecLen = Len(DCTransRec(1))
    DCTransFile = FreeFile
    Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
    NumOfTRecs& = LOF(DCTransFile) \ DCTransRecLen
    For TCnt& = 1 To NumOfTRecs&
      FrmShowPctComp.ShowPctComp TCnt&, NumOfTRecs&
      Get DCTransFile, TCnt&, DCTransRec(1)
      DCTransRec(1).GLInterfaced = "Y"
      Put DCTransFile, TCnt&, DCTransRec(1)
    Next
    Close DCTransFile
    Call MainLog("Init DC for IF.")
    Okcnt = Okcnt + 1
  End If
  End If
  If Okcnt > 0 Then
    MsgBox "Initialization Complete.", vbOKOnly, "Complete"
    Call MainLog("Init Complete.")
  Else
    Call MainLog("Init - No Trans.")
    MsgBox "No Transactions to Initialize.", vbOKOnly, "Complete"
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
