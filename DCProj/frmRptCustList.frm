VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCustList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Customer Listing"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12192
   ClipControls    =   0   'False
   Icon            =   "frmRptCustList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5298
      TabIndex        =   4
      Top             =   5040
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
      ColDesigner     =   "frmRptCustList.frx":08CA
   End
   Begin LpLib.fpCombo fpCombo1 
      Height          =   348
      Left            =   5298
      TabIndex        =   3
      Top             =   4524
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
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
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   0
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
      AutoMenu        =   0   'False
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptCustList.frx":0CA4
   End
   Begin LpLib.fpCombo fpcboCustStatus 
      Height          =   348
      Left            =   5298
      TabIndex        =   2
      Top             =   4008
      Width           =   2004
      _Version        =   196608
      _ExtentX        =   3535
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
      ColDesigner     =   "frmRptCustList.frx":1047
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
      Left            =   7848
      TabIndex        =   5
      Top             =   7392
      Width           =   1332
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
      Left            =   9648
      TabIndex        =   6
      Top             =   7392
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   9
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   593
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
            TextSave        =   "10:56 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5/19/2005"
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
      Left            =   5304
      TabIndex        =   1
      Top             =   3480
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
      Left            =   5298
      TabIndex        =   0
      Top             =   2976
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Book:"
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
      Index           =   2
      Left            =   3714
      TabIndex        =   13
      Top             =   3525
      Width           =   1380
   End
   Begin VB.Label Label3 
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
      Index           =   8
      Left            =   3618
      TabIndex        =   12
      Top             =   3036
      Width           =   1476
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Status:"
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
      Index           =   1
      Left            =   3018
      TabIndex        =   11
      Top             =   4014
      Width           =   2076
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type:"
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
      Index           =   0
      Left            =   2706
      TabIndex        =   10
      Top             =   5040
      Width           =   2388
   End
   Begin VB.Label Label2 
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
      Height          =   372
      Left            =   3312
      TabIndex        =   8
      Top             =   4524
      Width           =   1788
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Customer Listing Report"
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
      Left            =   3888
      TabIndex        =   7
      Top             =   1608
      Width           =   4452
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2988
      Left            =   2682
      Top             =   2664
      Width           =   6828
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   1248
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
Attribute VB_Name = "frmRptCustList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim BegRoute As String, EndRoute As String

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCombo1.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpCombo1.ListDown = True
  End If
  If fpCombo1.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboCustStatus.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptCustList by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  Load frmUBCustMenu
  DoEvents
  frmUBCustMenu.Show
  Unload frmRptCustList
  'LoadDisplayForm frmUBCustMenu, Me
End Sub

Private Sub cmdPrint_Click()
  If ValidRoutes Then
    DeActivateControls Me
    If fpcboRptType.ListIndex = 0 Then
      UBQuickCustList2 Me
    ElseIf fpcboRptType.ListIndex = 1 Then
      UBQuickCustList Me
      ActivateControls Me
    Else
      ActivateControls Me
    End If
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
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  fptxtRoute1 = "00"
  fptxtRoute2 = "99"
  fpcboCustStatus.AddItem "ALL"
  fpcboCustStatus.AddItem "Active"
  fpcboCustStatus.AddItem "Inactive"
  fpcboCustStatus.AddItem "Balance"
  fpcboCustStatus.AddItem "Pending"
  fpcboCustStatus.AddItem "Delinquent"
  fpcboCustStatus.AddItem "Final"
  fpcboCustStatus.ListIndex = 0

  fpCombo1.AddItem "Customer Name Order"
  fpCombo1.AddItem "Location Number Order"
  fpCombo1.AddItem "Read Sequence Number Order"
  fpCombo1.AddItem "Account Number Order"
  fpCombo1.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Public Sub UBQuickCustList(Parent As Form)
  Dim Dash80 As String, IdxName As String, RStatus As String
  Dim Title As String, ReportFile As String, UseStatus As Boolean
  Dim UBCustRecLen As Integer, IdxRecLen As Integer, Stat As String
  Dim IdxFileSize As Long, cnt As Long, AcctNumber As Long
  Dim IdxNumOfRecs As Long, LineCnt As Integer, CStatus As String
  Dim Handle As Integer, UBCust As Integer, NumOfRecs As Long
  Dim UBRpt As Integer, CustCnt As Long, Pending As Long
  Dim Active As Long, Final As Long, InActive As Long, Book As String
  Dim Balance As Long, UnKnown As Long, DeletedCnt As Long
  Dim RetCode As Integer, EntryPoint As Integer, Delinquent As Long
  Dim AbortFlag As Boolean, UsingName As Boolean, UsingAcct As Boolean
  Dim UsingBook As Boolean, UsingRead As Boolean, SEQNUMB As String
  Dim bk As Integer
  MaxLines = 59
  PageNo = 0
  ReportFile = UBPath$ + "UBLOLIST.RPT"
  Dash80$ = String$(80, "-")
  BegRoute = fptxtRoute1
  EndRoute = fptxtRoute2
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 2)
  Select Case RStatus$
  Case "Ac"
    UseStatus = True
    Stat$ = " ACTIVE"
  Case "In"
      UseStatus = True
    Stat$ = " INACTIVE"
  Case "Ba"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  Case "Pe"
    Stat$ = " PENDING"
    UseStatus = True
  Case "De"
    Stat$ = " DELINQUENT"
    UseStatus = True
  Case "Fi"
    Stat$ = " FINAL"
    UseStatus = True
  Case Else
    Stat$ = " ALL"
    UseStatus = False
  End Select
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 1)


  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBCustBlank(1) As NewUBCustRecType
  
  UBCustRecLen = Len(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer

  Select Case fpCombo1.ListIndex
  Case 0
    UsingName = True
    IdxName$ = UBPath$ + "UBCUSTNM.IDX"
    Title$ = "Quick Customer Listing by Name."
  Case 1
    UsingBook = True
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
    Title$ = "Quick Customer Listing by Location."
  Case 2
    UsingRead = True
    MakeSequenceIndex "Sequence Number", Parent
    IdxName$ = TempIndexName
    Title$ = "Quick Customer Listing by Sequence No."
  Case 3
    UsingAcct = True
    IdxName$ = ""
    Title$ = "Quick Customer Listing by Account."
  End Select
  If Not UsingAcct Then
    IdxFileSize& = FileSize(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    NumOfRecs = IdxNumOfRecs
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  UBRpt = FreeFile
  Open ReportFile For Output As UBRpt
  
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show

  GoSub DoLocaRptHeader

  For cnt = 1 To NumOfRecs
    If Not UsingAcct Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
      AcctNumber = IdxBuff(cnt).RecNum
    Else
      Get UBCust, cnt, UBCustRec(1)
      AcctNumber = cnt
    End If
    If LineCnt > MaxLines Then
      Print #UBRpt, Chr$(12)
      GoSub DoLocaRptHeader
    End If
    If UBCustRec(1).DelFlag = 0 Then
      If UseStatus Then           'if they care about the cust status, or want all.
        CStatus$ = Left$(QPTrim$(UBCustRec(1).Status), 1)
        If CStatus$ <> RStatus$ Then
          GoTo SkipEm
        End If
      End If
      Book$ = QPTrim$(UBCustRec(1).Book)
      SEQNUMB$ = QPTrim$(UBCustRec(1).SEQNUMB)
      If Len(Book$) = 0 Then
        Book$ = "  "
      End If
      bk = Val(Book$)
      If bk < Val(BegRoute) Or bk > Val(EndRoute) Then
        GoTo SkipEm
      End If

      
      Print #UBRpt, QPTrim$(UBCustRec(1).Book); Tab(3); "-"; QPTrim$(UBCustRec(1).SEQNUMB);
      Print #UBRpt, Tab(11); Using$("#####", AcctNumber);
      Print #UBRpt, Tab(18); QPTrim$(Left$(UBCustRec(1).CustName, 25));
      If Len(QPTrim$(UBCustRec(1).Status)) = 0 Then
        UBCustRec(1).Status = "?"
      End If
      
      If UsingRead Then
        Print #UBRpt, Tab(45); QPTrim$(Left$(UBCustRec(1).ServAddr, 24));
        If UBCustRec(1).Seq < 0 Then
          UBCustRec(1).Seq = 0
        End If
        Print #UBRpt, Tab(70); UBCustRec(1).Status; Using$("#########", UBCustRec(1).Seq)
      Else
        Print #UBRpt, Tab(45); QPTrim$(Left$(UBCustRec(1).ServAddr, 25));
        Print #UBRpt, Tab(70); UBCustRec(1).Status; "     "; QPTrim$(UBCustRec(1).ZONE)
      End If
      CustCnt = CustCnt + 1
      Select Case UBCustRec(1).Status
      Case "A"
        Active = Active + 1
      Case "F"
        Final = Final + 1
      Case "I"
        InActive = InActive + 1
      Case "B"
        Balance = Balance + 1
      Case "P"
        Pending = Pending + 1
      Case "D"
        Delinquent = Delinquent + 1
      Case Else
        UnKnown = UnKnown + 1
      End Select
      LineCnt = LineCnt + 1
    Else
      DeletedCnt = DeletedCnt + 1
    End If
SkipEm:
    
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
  Next
  
  GoSub DoLocaRptTotals
  
  Close UBCust, UBRpt
  
  Erase IdxBuff, UBCustRec
  If Not AbortFlag Then
    ViewPrint ReportFile$, "Quick Customer Listing"
  End If
  
  'KillFile "UBLOLIST.RPT"
  'Parent.Enabled = True
Exit Sub
  
DoLocaRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, "Customer Listing Report      "; "Date: "; Date$; Tab(70); "Page: "; PageNo
  Print #UBRpt, "Report Options - "; "Status("; Stat$; "), Book Range("; BegRoute; " - "; EndRoute; ")"
  If UsingRead Then
    Print #UBRpt, "           Acct                                                    Acct    SEQ."
    Print #UBRpt, "Location    No.  Customer Name              Service Address        Status  Numb"
  Else
    Print #UBRpt, "           Acct                                                    Acct    Post"
    Print #UBRpt, "Location    No.  Customer Name              Service Address        Status  Route"
  End If
  Print #UBRpt, Dash80$
  LineCnt = 4
Return
  
DoLocaRptTotals:
  PageNo = PageNo + 1
  Print #UBRpt,
  Print #UBRpt, Dash80$
  Print #UBRpt, "Customer Summary"
  Print #UBRpt,
  Print #UBRpt, "  Active: "; Using$("#####", Active)
  Print #UBRpt, "   Final: "; Using$("#####", Final)
  Print #UBRpt, "Inactive: "; Using$("#####", InActive)
  Print #UBRpt, " Balance: "; Using$("#####", Balance)
  Print #UBRpt, " Pending: "; Using$("#####", Pending)
  Print #UBRpt, "Delinqnt: "; Using$("#####", Delinquent)
  Print #UBRpt, " Unknown: "; Using$("#####", UnKnown)
  Print #UBRpt, " Deleted: "; Using$("#####", DeletedCnt)
  Print #UBRpt,
  Print #UBRpt, "   TOTAL: "; Using$("#####", CustCnt)
  Print #UBRpt, Chr$(12)
Return
  
  
End Sub
Public Sub UBQuickCustList2(Parent As Form)
  Dim IdxName As String, ToPrint As String, NumOfRecs As Long
  Dim Title As String, ReportFile As String, RStatus As String
  Dim UBCustRecLen As Integer, IdxRecLen As Integer, UseStatus As Boolean
  Dim IdxFileSize As Long, cnt As Long, UsingAcct As Boolean
  Dim IdxNumOfRecs As Long, LineCnt As Integer, AcctNumber As Long
  Dim Handle As Integer, UBCust As Integer, UsingName As Boolean
  Dim UsingBook As Boolean, UsingRead As Boolean, Stat As String
  Dim UBRpt As Integer, CustCnt As Long, Pending As Long
  Dim Active As Long, Final As Long, InActive As Long, CStatus As String
  Dim Balance As Long, UnKnown As Long, DeletedCnt As Long
  Dim RetCode As Integer, EntryPoint As Integer, Delinquent As Long
  Dim Book As String, SEQNUMB As String, bk As Integer
  ReportFile = UBPath$ + "UBLOLIST.RPT"
  BegRoute = fptxtRoute1
  EndRoute = fptxtRoute2
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 2)
  Select Case RStatus$
  Case "Ac"
    UseStatus = True
    Stat$ = " ACTIVE"
  Case "In"
      UseStatus = True
    Stat$ = " INACTIVE"
  Case "Ba"
    UseStatus = True
    Stat$ = " BALANCE DUE"
  Case "Pe"
    Stat$ = " PENDING"
    UseStatus = True
  Case "De"
    Stat$ = " DELINQUENT"
    UseStatus = True
  Case "Fi"
    Stat$ = " FINAL"
    UseStatus = True
  Case Else
    Stat$ = " ALL"
    UseStatus = False
  End Select
  RStatus$ = Mid$(fpcboCustStatus.Text, 1, 1)
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBCustBlank(1) As NewUBCustRecType
  
  UBCustRecLen = Len(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer

  Select Case fpCombo1.ListIndex
  Case 0
    UsingName = True
    IdxName$ = UBPath$ + "UBCUSTNM.IDX"
    Title$ = "Quick Customer Listing by Name."
  Case 1
    UsingBook = True
    IdxName$ = UBPath$ + "UBCUSTBK.IDX"
    Title$ = "Quick Customer Listing by Location."
  Case 2
    UsingRead = True
    MakeSequenceIndex "Sequence Number", Parent
    IdxName$ = TempIndexName
    Title$ = "Quick Customer Listing by Sequence No."
  Case 3
    UsingAcct = True
    IdxName$ = ""
    Title$ = "Quick Customer Listing by Account."
  End Select
  If Not UsingAcct Then
    IdxFileSize& = FileSize(IdxName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    NumOfRecs = IdxNumOfRecs
    ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IdxName$ For Random Shared As Handle Len = IdxRecLen
    For cnt& = 1 To IdxNumOfRecs
      Get #Handle, cnt&, IdxBuff(cnt&)
    Next
    Close Handle
  Else
    NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  End If

  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  UBRpt = FreeFile
  Open ReportFile For Output As UBRpt
  
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show

  
  For cnt = 1 To NumOfRecs
    If Not UsingAcct Then
      Get UBCust, IdxBuff(cnt).RecNum, UBCustRec(1)
      AcctNumber = IdxBuff(cnt).RecNum
    Else
      Get UBCust, cnt, UBCustRec(1)
      AcctNumber = cnt
    End If
    If UBCustRec(1).DelFlag = 0 Then
      If UseStatus Then           'if they care about the cust status, or want all.
        CStatus$ = Left$(QPTrim$(UBCustRec(1).Status), 1)
        If CStatus$ <> RStatus$ Then
          GoTo SkipEm
        End If
      End If
      Book$ = QPTrim$(UBCustRec(1).Book)
      SEQNUMB$ = QPTrim$(UBCustRec(1).SEQNUMB)
      If Len(Book$) = 0 Then
        Book$ = "  "
      End If
      bk = Val(Book$)
      If bk < Val(BegRoute) Or bk > Val(EndRoute) Then
        GoTo SkipEm
      End If

      ToPrint$ = QPTrim$(UBCustRec(1).Book) + "-" + QPTrim$(UBCustRec(1).SEQNUMB)
      ToPrint$ = ToPrint$ + "~" + Str$(AcctNumber)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(Left$(UBCustRec(1).CustName, 25))
      If Len(QPTrim$(UBCustRec(1).Status)) = 0 Then
        UBCustRec(1).Status = "?"
      End If
      
      If UsingRead Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(Left$(UBCustRec(1).ServAddr, 24))
        If UBCustRec(1).Seq < 0 Then
          UBCustRec(1).Seq = 0
        End If
        ToPrint$ = ToPrint$ + "~" + UBCustRec(1).Status + "~" + Using$("#########", UBCustRec(1).Seq)
      Else
        ToPrint$ = ToPrint$ + "~" + QPTrim$(Left$(UBCustRec(1).ServAddr, 25))
        ToPrint$ = ToPrint$ + "~" + UBCustRec(1).Status + "~" + QPTrim$(UBCustRec(1).ZONE)
      End If
      CustCnt = CustCnt + 1
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      Select Case UBCustRec(1).Status
      Case "A"
        Active = Active + 1
      Case "F"
        Final = Final + 1
      Case "I"
        InActive = InActive + 1
      Case "B"
        Balance = Balance + 1
      Case "P"
        Pending = Pending + 1
      Case "D"
        Delinquent = Delinquent + 1
      Case Else
        UnKnown = UnKnown + 1
      End Select
      LineCnt = LineCnt + 1
    Else
      DeletedCnt = DeletedCnt + 1
    End If
SkipEm:
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls Me
      GoTo ExitRpt
    End If
  Next
  
  GoSub DoLocaRptTotals
  
  Close UBCust, UBRpt
  
  Erase IdxBuff, UBCustRec
  GoSub DoLocaRptHeader
  Load frmLoadingRpt
  frmLoadingRpt.setwherefrom frmRptCustList
  ARptQCustList.txtDate = Now
  ARptQCustList.txtTown = TOWNNAME$
  ARptQCustList.Title = Title$
  ARptQCustList.lblRptOpt.Caption = "Report Options - Status(" & Stat$ & "), Book Range(" & BegRoute & " - " & EndRoute & ")"
  ARptQCustList.GetName ReportFile$
  ARptQCustList.startrpt

'  If Not AbortFlag Then
'    ViewPrint ReportFile$, "Quick Customer Listing"
'  End If
  
  'KillFile "UBLOLIST.RPT"
  'Parent.Enabled = True
Exit Sub
  
DoLocaRptHeader:
  If UsingRead Then
    ARptQCustList.Label6 = "SEQ Numb"
  Else
   ARptQCustList.Label6 = "Post Route"
  End If
Return
  
DoLocaRptTotals:
  ARptQCustList.totActive = Using$("#####", Active)
  ARptQCustList.totFinal = Using$("#####", Final)
  ARptQCustList.totInactive = Using$("#####", InActive)
  ARptQCustList.totBalance = Using$("#####", Balance)
  ARptQCustList.totPending = Using$("#####", Pending)
  ARptQCustList.totDelinquent = Using$("#####", Delinquent)
  ARptQCustList.totUnknown = Using$("#####", UnKnown)
  ARptQCustList.totDeleted = Using$("#####", DeletedCnt)
  ARptQCustList.totTotal = Using$("#####", CustCnt)
Return
ExitRpt:
  Exit Sub
  
End Sub
Private Function ValidRoutes()
  If fptxtRoute1 <> "" And fptxtRoute2 <> "" Then
    If fptxtRoute1 > fptxtRoute2 Then
      MsgBox "Invalid Book Selection, The Beginning Book Should Be Less or Equal to Ending Book.", vbOKOnly, "Invalid Selection"
      ValidRoutes = False
    Else
      ValidRoutes = True
      BegRoute = QPTrim(fptxtRoute1)
      EndRoute = QPTrim(fptxtRoute2)
    End If
  Else
    MsgBox "Book Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

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
    fpcboCustStatus.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpcboCustStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboCustStatus.ListDown = True
  End If
  If fpcboCustStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpCombo1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
