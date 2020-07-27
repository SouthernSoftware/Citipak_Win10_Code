VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptStreetList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Street Listing"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptStreetList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5532
      TabIndex        =   7
      Top             =   5904
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptStreetList.frx":08CA
   End
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5532
      TabIndex        =   3
      Top             =   3804
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      EditMarginLeft  =   2
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptStreetList.frx":0BF8
   End
   Begin LpLib.fpCombo fpcboInactive 
      Height          =   348
      Left            =   5532
      TabIndex        =   6
      Top             =   5376
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptStreetList.frx":0F1B
   End
   Begin LpLib.fpCombo fpcboPageBreak 
      Height          =   348
      Left            =   5532
      TabIndex        =   2
      Top             =   3276
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   614
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptStreetList.frx":1249
   End
   Begin EditLib.fpText fptxtSearch 
      Height          =   348
      Left            =   5520
      TabIndex        =   5
      Top             =   4848
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
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
      ThreeDOutsideStyle=   2
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
   Begin VB.CommandButton cmdExit 
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
      Left            =   10080
      TabIndex        =   9
      Top             =   7560
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F10 &Print"
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
      Left            =   8400
      TabIndex        =   8
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
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
            TextSave        =   "12:34 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/23/2005"
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5532
      TabIndex        =   1
      Top             =   2772
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      ThreeDOutsideStyle=   2
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
      CharValidationText=   "0123456789"
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtRoute1 
      Height          =   348
      Left            =   5532
      TabIndex        =   0
      Top             =   2256
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      ThreeDOutsideStyle=   2
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
      CharValidationText=   "0123456789"
      MaxLength       =   2
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
   Begin EditLib.fpText fptxtCustType 
      Height          =   348
      Left            =   5532
      TabIndex        =   4
      Top             =   4332
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
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
      ThreeDOutsideStyle=   2
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break On Book:"
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
      Height          =   396
      Left            =   2856
      TabIndex        =   19
      Top             =   3312
      Width           =   2484
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Inactive: "
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
      Left            =   3000
      TabIndex        =   18
      Top             =   5340
      Width           =   2364
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Search For:"
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
      Left            =   3312
      TabIndex        =   17
      Top             =   4848
      Width           =   2004
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Type:"
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
      Height          =   324
      Index           =   0
      Left            =   3264
      TabIndex        =   16
      Top             =   4368
      Width           =   2076
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
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
      Height          =   324
      Index           =   7
      Left            =   3624
      TabIndex        =   15
      Top             =   3852
      Width           =   1716
   End
   Begin VB.Label Label2 
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
      Left            =   2976
      TabIndex        =   14
      Top             =   5880
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4620
      Left            =   2472
      Top             =   1920
      Width           =   7284
   End
   Begin VB.Label LabelB1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Book:"
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
      Height          =   324
      Left            =   3864
      TabIndex        =   13
      Top             =   2316
      Width           =   1476
   End
   Begin VB.Label LabelB2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Thru Book:"
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
      Height          =   324
      Left            =   3960
      TabIndex        =   12
      Top             =   2808
      Width           =   1380
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   840
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Street Listing"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   1080
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   720
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
Attribute VB_Name = "frmRptStreetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String
Dim UseCycle As Boolean
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptStreetList
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptStreetList by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fptxtRoute1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtRoute1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute2.SetFocus
  End If
End Sub
Private Sub fptxtRoute2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboPageBreak.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboInactive.SetFocus
  End If
End Sub

Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fptxtCustType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboPageBreak.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboPageBreak_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPageBreak.ListDown = True
  End If
  If fpcboPageBreak.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtCustType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtSearch.SetFocus
  End If
End Sub
Private Sub fpcboInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboInactive.ListDown = True
  End If
  If fpcboInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtSearch.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboInactive.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Selection, The Beginning Value Should Be Less or Equal to Ending Value.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Fields May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub cmdPrint_Click()
  If ValidRoutes Then
    If fpcboPageBreak.ListIndex = 1 And fpcboPrintOrder.ListIndex <> 2 Then
       MsgBox "If you select page break on book then you must select print order by location.", vbOKOnly, "Invalid Selections"
       fpcboPrintOrder.SetFocus
       Exit Sub
    End If
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      CustStreetListing2
    ElseIf fpcboRptType.ListIndex = 1 Then
      CustStreetListing
      ActivateControls Me, True
    Else
      ActivateControls Me, True
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBSetupreclen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  Me.HelpContextID = hlpCustomerStreet
  If UBSetUp(1).BILLCYCL = "Y" Then
    UseCycle = True
  End If

  If UseCycle Then
    LabelB1.Caption = "From Cycle"
    LabelB2.Caption = "Thru Cycle"
  Else
    LabelB1.Caption = "From Book"
    LabelB2.Caption = "Thru Book"
  End If
  fptxtRoute1 = "00"
  fptxtRoute2 = "99"
  fptxtSearch = ""
  fptxtCustType = ""
  fpcboPageBreak.AddItem "No"
  fpcboPageBreak.AddItem "Yes"
  fpcboPageBreak.ListIndex = 0
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  fpcboPrintOrder.AddItem "Location Number Order"
  fpcboPrintOrder.AddItem "Service Address Order"
  fpcboPrintOrder.ListIndex = 0
  fpcboInactive.AddItem "Yes"
  fpcboInactive.AddItem "No"
  fpcboInactive.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub CustStreetListing()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, Ctype As String, UsingName As Boolean
  Dim PageNo As Integer, UseStatus As Boolean, AcctNo As Long
  Dim Dash80 As String, IndexName As String, ChkCType As Boolean
  Dim IdxRecLen As Integer, IdxFileSize As Long, Book As String
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer
  Dim CStatus As String, CUSTTYPE As String, inactiveflag As Boolean
  Dim SEQNUMB As String, CustBook As Integer, ChkBalance As Double
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim StrSrch As String, UsingAcct As Boolean, ThisBook As Integer
  Dim First As Integer, Last As Integer, FirstBook As Boolean
  Dim Order As String, GroverFlag As Boolean, CustCheck As Integer
  Dim UsingAddr As Boolean, ReportFile As String, Tmptyp As String
  UsingAcct = False
  UseStatus = False
  UsingName = False
  UsingBook = False
  UsingAddr = False
  PageNo = 0
  If fpcboPageBreak.ListIndex = 1 Then
    GroverFlag = True
  End If
  FrmShowPctComp.Label1 = "Creating Customer Street Listing"
  FrmShowPctComp.Show , Me

  First = BegRoute
  Last = EndRoute
  Order$ = Left$(QPTrim$(fpcboPrintOrder.Text), 1)
  Ctype$ = QPTrim$(fptxtCustType)
  If Len(Ctype$) > 0 Then
    ChkCType = True
  End If
  StrSrch$ = QPTrim$(fptxtSearch)
  If fpcboInactive.ListIndex = 0 Then
    inactiveflag = True
  Else
    inactiveflag = False
  End If

    Select Case Order$
    Case "C"
      IndexName$ = NameIndexFile
      UsingName = True
      
    Case "A"
      IndexName$ = ""
      UsingAcct = True
      
    Case "L"
      IndexName$ = BookIndexFile
      UsingBook = True
    Case "S"
      IndexName$ = TempIndexName
      UsingAddr = True
  
    Case Else
      'GoSub ShowErrScrn
    End Select


  '***************
  MaxLines = 55

  PageNo = 0
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  'AgeDate = Date2Num%(DATE$)

  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptMastCust
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBSTLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt


 ' ShowProcessingScrn

  GoSub DoCutOffRptHeader

  ThisBook = 0
  FirstBook = True

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitCutOffListing
    End If

    If UsingName Or UsingBook Or UsingAddr Then
      AcctNo& = IdxBuff(cnt).RecNum
    Else
      AcctNo& = cnt
    End If
    Get UBCust, AcctNo&, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      If UseCycle Then
        CustCheck = UBCustRec(1).BILLCYCL
      Else
        CustCheck = Val(UBCustRec(1).Book)
      End If

      If CustCheck >= First And CustCheck <= Last Then
        If Not inactiveflag Then
          If UBCustRec(1).Status <> "A" Then
            GoTo SkipEMStreet
          End If
        End If
        If ChkCType Then
          CUSTTYPE$ = QPTrim$(UBCustRec(1).CUSTTYPE)
          If CUSTTYPE$ <> Ctype$ Then
            GoTo SkipEMStreet
          End If
        End If
        If InStr(UBCustRec(1).ServAddr, StrSrch$) > 0 Then
          If FirstBook Then
            FirstBook = False
            ThisBook = Val(UBCustRec(1).Book)
          End If
          GoSub PrintLine
        End If
      End If
SkipEMStreet:
    End If
    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub DoCutOffRptHeader
    End If
   ' ShowPctComp cnt, NumOfRecs
  Next

  GoSub DoCutOffRptFooter:

  Close
  Erase IdxBuff, UBCustRec
  '  PrintRptFile , "UBSTLIST.RPT", 1, RetCode, Entry
  ViewPrint ReportFile$, "Customer Street Listing."

  KillFile "UBSTLIST.RPT"

ExitCutOffListing:

  Exit Sub

DoCutOffRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(26); "Customer Street Listing Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$
  Print #UBRpt, "Location  Acct#/Type   Customer Name"; Tab(56); "Service Address"
  Print #UBRpt, Dash80$
  LineCnt = 6
Return
DoCutOffRptFooter:
  Print #UBRpt, ""
  Print #UBRpt, "Total Customers: "; Using("#####,#", CustomerCnt)
  Print #UBRpt,
  Print #UBRpt, "         Report Parameters"
  Print #UBRpt, "                      From:"; First; "  Thru:"; Last
  Print #UBRpt, "               Search Text: ";
  If Len(StrSrch$) = 0 Then
    Print #UBRpt, "NONE."
  Else
    Print #UBRpt, StrSrch$
  End If
  Print #UBRpt, "             Customer Type: ";
  If ChkCType Then
    Print #UBRpt, Ctype$
  Else
    Print #UBRpt, "ALL"
  End If
  Print #UBRpt, "          Include Inactive: ";
  If inactiveflag Then
    Print #UBRpt, "Y"
  Else
    Print #UBRpt, "N"
  End If
Return


PrintLine:

  If GroverFlag Then
    If UsingBook Then
      CustBook = Val(UBCustRec(1).Book)
      If ThisBook <> CustBook Then
        ThisBook = CustBook
        Print #UBRpt, FF$
        GoSub DoCutOffRptHeader
      End If
    End If
  End If
  '*************************************
  '   Main body of Printing goes here
  If Len(QPTrim$(UBCustRec(1).CUSTTYPE)) > 0 Then
    Tmptyp$ = " /" + QPTrim$(UBCustRec(1).CUSTTYPE)
  Else
    Tmptyp$ = ""
  End If

  Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Using("######", AcctNo&); Tmptyp$;
  Print #UBRpt, Tab(25); Left$(UBCustRec(1).CustName, 30); Tab(56); UBCustRec(1).ServAddr
  LineCnt = LineCnt + 1
  Print #UBRpt, String$(79, "-")
  LineCnt = LineCnt + 1

Skip1:
    Print #UBRpt,
    LineCnt = LineCnt + 2
  CustomerCnt = CustomerCnt + 1
  '*************************************
Return

End Sub
Private Sub CustStreetListing2()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, Ctype As String, UsingName As Boolean
  Dim UseStatus As Boolean, AcctNo As Long
  Dim IndexName As String, ChkCType As Boolean
  Dim IdxRecLen As Integer, IdxFileSize As Long, Book As String
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim cnt As Long, UBCust As Integer, UBRpt As Integer
  Dim CStatus As String, CUSTTYPE As String, inactiveflag As Boolean
  Dim SEQNUMB As String, CustBook As Integer, ChkBalance As Double
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim StrSrch As String, UsingAcct As Boolean, ThisBook As Integer
  Dim First As Integer, Last As Integer, FirstBook As Boolean
  Dim Order As String, GroverFlag As Boolean, CustCheck As Integer
  Dim UsingAddr As Boolean, ToPrint As String, ToPrnH1 As String
  Dim ToPrnH2 As String, ReportFile As String, Tmptyp As String
  UsingAcct = False
  UseStatus = False
  UsingName = False
  UsingBook = False
  UsingAddr = False
  If fpcboPageBreak.ListIndex = 1 Then
    GroverFlag = True
  End If
  FrmShowPctComp.Label1 = "Creating Customer Street Listing"
  FrmShowPctComp.Show , Me

  First = BegRoute
  Last = EndRoute
  Order$ = Left$(QPTrim$(fpcboPrintOrder.Text), 1)
  Ctype$ = QPTrim$(fptxtCustType)
  If Len(Ctype$) > 0 Then
    ChkCType = True
  End If
  StrSrch$ = QPTrim$(fptxtSearch)
  If fpcboInactive.ListIndex = 0 Then
    inactiveflag = True
  Else
    inactiveflag = False
  End If

    Select Case Order$
    Case "C"
      IndexName$ = NameIndexFile
      UsingName = True
    Case "A"
      IndexName$ = ""
      UsingAcct = True
    Case "L"
      IndexName$ = BookIndexFile
      UsingBook = True
    Case "S"
      IndexName$ = TempIndexName
      UsingAddr = True
    Case Else
      'GoSub ShowErrScrn
    End Select


  '***************

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  'AgeDate = Date2Num%(DATE$)

  If UsingName Or UsingBook Then
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  ElseIf UsingAddr Then
'unrem
    SortServiceAddrs frmRptMastCust
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle

  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  ReportFile$ = UBPath$ + "UBSTLIST.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt


 ' ShowProcessingScrn "Customer Street Listing."

  ThisBook = 0
  FirstBook = True

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me, True
      GoTo ExitCutOffListing
    End If

    If UsingName Or UsingBook Or UsingAddr Then
      AcctNo& = IdxBuff(cnt).RecNum
    Else
      AcctNo& = cnt
    End If
    Get UBCust, AcctNo&, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      If UseCycle Then
        CustCheck = UBCustRec(1).BILLCYCL
      Else
        CustCheck = Val(UBCustRec(1).Book)
      End If

      If CustCheck >= First And CustCheck <= Last Then
        If Not inactiveflag Then
          If UBCustRec(1).Status <> "A" Then
            GoTo SkipEMStreet
          End If
        End If
        If ChkCType Then
          CUSTTYPE$ = QPTrim$(UBCustRec(1).CUSTTYPE)
          If CUSTTYPE$ <> Ctype$ Then
            GoTo SkipEMStreet
          End If
        End If
        If InStr(UBCustRec(1).ServAddr, StrSrch$) > 0 Then
          If FirstBook Then
            FirstBook = False
            ThisBook = Val(UBCustRec(1).Book)
          End If
          GoSub PrintLine
        End If
      End If
SkipEMStreet:
    End If
   ' ShowPctComp cnt, NumOfRecs
  Next


  Close
  Erase IdxBuff, UBCustRec
  '  PrintRptFile , "UBSTLIST.RPT", 1, RetCode, Entry
'  ViewPrint "UBSTLIST.RPT", "Customer Street Listing."
'for ar report send fpcbopagebreak.listindex with file name
'1 is page break 0 is not
'  KillFile "UBSTLIST.RPT"
  If Not ChkCType Then
    Ctype$ = "ALL"
  End If

  If CustomerCnt > 0 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptStreetList
    ARptStreetList.txtDate = Now
    ARptStreetList.txtTown = TOWNNAME$
    ARptStreetList.Title = "Customer Street Listing."
    ARptStreetList.totCust = CustomerCnt
    ToPrnH1$ = LabelB1.Caption + ": " + Str(First) + " " + LabelB2.Caption + ": " + Str(Last)
    ToPrnH1$ = ToPrnH1$ + "     Printing Order: " + fpcboPrintOrder.Text + " "
    ToPrnH2$ = "Customer Type: " + Ctype$ + "     Include Inactive: " + fpcboInactive.Text
    ToPrnH2$ = ToPrnH2$ + "     Search Text: " + StrSrch$
    ARptStreetList.txtRptParm1 = ToPrnH1$
    ARptStreetList.txtRptParm2 = ToPrnH2$
    ARptStreetList.GetName ReportFile$, fpcboPageBreak.ListIndex
    ARptStreetList.startrpt

  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
    ActivateControls Me, True
  End If

ExitCutOffListing:

  Exit Sub

'  Print #UBRpt, "         Report Parameters"
'  Print #UBRpt, "                      From:"; First; "  Thru:"; Last
'  Print #UBRpt, "               Search Text: ";
'  If Len(StrSrch$) = 0 Then
'    Print #UBRpt, "NONE."
'  Else
'    Print #UBRpt, StrSrch$
'  End If
'  Print #UBRpt, "             Customer Type: ";
'  If ChkCType Then
'    Print #UBRpt, Ctype$
'  Else
'    Print #UBRpt, "ALL"
'  End If
'  Print #UBRpt, "          Include Inactive: ";
'  If inactiveflag Then
'    Print #UBRpt, "Y"
'  Else
'    Print #UBRpt, "N"
'  End If
'Return


PrintLine:

'  If GroverFlag Then
'    If UsingBook Then
'      CustBook = Val(UBCustRec(1).Book)
'      If ThisBook <> CustBook Then
'        ThisBook = CustBook
'        Print #UBRpt, FF$
'        GoSub DoCutOffRptHeader
'      End If
'    End If
'  End If
  '*************************************
  '   Main body of Printing goes here
  If Len(QPTrim$(UBCustRec(1).CUSTTYPE)) > 0 Then
    Tmptyp$ = " /" + QPTrim$(UBCustRec(1).CUSTTYPE)
  Else
    Tmptyp$ = ""
  End If
  ToPrint$ = UBCustRec(1).Book + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~" + Str$(AcctNo&) + Tmptyp$ + "~"
  ToPrint$ = ToPrint$ + Left$(UBCustRec(1).CustName, 30) + "~" + UBCustRec(1).ServAddr
  Print #UBRpt, ToPrint$
  CustomerCnt = CustomerCnt + 1
  Return
Skip1:
  CustomerCnt = CustomerCnt + 1
  '*************************************
Return

End Sub

