VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBillInfoLetter 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Letter Setup Information"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmBillInfoLetter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboLogo 
      Height          =   330
      Left            =   9690
      TabIndex        =   21
      Top             =   2130
      Width           =   840
      _Version        =   196608
      _ExtentX        =   1482
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ColDesigner     =   "frmBillInfoLetter.frx":08CA
   End
   Begin LpLib.fpCombo fpcboMtrNum 
      Height          =   330
      Left            =   9690
      TabIndex        =   0
      Top             =   2490
      Width           =   1725
      _Version        =   196608
      _ExtentX        =   3043
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ColDesigner     =   "frmBillInfoLetter.frx":0C30
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "F5 &Test Print"
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
      Left            =   990
      TabIndex        =   24
      Top             =   7740
      Width           =   1644
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10 &Save"
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
      Left            =   8424
      TabIndex        =   7
      Top             =   7752
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
      Left            =   10080
      TabIndex        =   8
      Top             =   7752
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
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
            TextSave        =   "11:45 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "12/8/2008"
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
   Begin EditLib.fpText fpTxtHead1 
      Height          =   300
      Left            =   1032
      TabIndex        =   1
      Top             =   2016
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      Text            =   "TOWN OF ANYWHERE"
      CharValidationText=   ""
      MaxLength       =   30
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
   Begin EditLib.fpText fptxtHead3 
      Height          =   300
      Left            =   1032
      TabIndex        =   3
      Top             =   2616
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
   Begin EditLib.fpText fpTxtHead2 
      Height          =   300
      Left            =   1032
      TabIndex        =   2
      Top             =   2316
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
   Begin EditLib.fpText fpTxtOpt5 
      Height          =   300
      Left            =   1008
      TabIndex        =   6
      Top             =   7032
      Width           =   4476
      _Version        =   196608
      _ExtentX        =   7895
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   30
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
   Begin EditLib.fpText fpTxtOpt3 
      Height          =   300
      Left            =   1008
      TabIndex        =   4
      Top             =   6048
      Width           =   5844
      _Version        =   196608
      _ExtentX        =   10308
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   50
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
   Begin EditLib.fpText fpTxtOpt4 
      Height          =   300
      Left            =   1008
      TabIndex        =   5
      Top             =   6720
      Width           =   4476
      _Version        =   196608
      _ExtentX        =   7895
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   7.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   30
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
   Begin EditLib.fpText fptxtOpt1 
      Height          =   300
      Left            =   1032
      TabIndex        =   22
      Top             =   3480
      Width           =   4500
      _Version        =   196608
      _ExtentX        =   7937
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
         Size            =   7.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
   Begin EditLib.fpText fpTxtOpt2 
      Height          =   300
      Left            =   1032
      TabIndex        =   23
      Top             =   3792
      Width           =   4500
      _Version        =   196608
      _ExtentX        =   7937
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
         Size            =   7.5
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   252
      Index           =   0
      Left            =   1008
      TabIndex        =   25
      Top             =   4488
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   252
      Index           =   1
      Left            =   1008
      TabIndex        =   26
      Top             =   4728
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   252
      Index           =   2
      Left            =   1008
      TabIndex        =   27
      Top             =   4968
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   252
      Index           =   3
      Left            =   1008
      TabIndex        =   28
      Top             =   5208
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   125
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
   Begin EditLib.fpText fptxtPgph 
      Height          =   252
      Index           =   4
      Left            =   1008
      TabIndex        =   29
      Top             =   5448
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tw Cen MT"
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
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
      MaxLength       =   125
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
   Begin VB.Line Line5 
      X1              =   10080
      X2              =   10080
      Y1              =   3960
      Y2              =   4296
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill #19 lines end here \/"
      Height          =   228
      Left            =   8424
      TabIndex        =   30
      Top             =   4200
      Width           =   2340
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Meter Number to Print On Bill:"
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
      Left            =   5784
      TabIndex        =   20
      Top             =   2544
      Width           =   3852
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Town Logo on Bill:"
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
      Height          =   324
      Left            =   7200
      TabIndex        =   19
      Top             =   2208
      Width           =   2436
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(5 Lines Available)  "
      Height          =   276
      Left            =   5016
      TabIndex        =   18
      Top             =   4272
      Width           =   1404
   End
   Begin VB.Line Line4 
      X1              =   11832
      X2              =   11832
      Y1              =   1584
      Y2              =   7464
   End
   Begin VB.Line Line3 
      X1              =   576
      X2              =   11856
      Y1              =   7632
      Y2              =   7632
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   600
      X2              =   600
      Y1              =   1776
      Y2              =   7632
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Letter Laser Bill Format #16 Default Information  Bill #19 Paragraph lines only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   504
      TabIndex        =   17
      Top             =   1128
      Width           =   5364
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Letter Setup Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   16
      Top             =   408
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   636
      Left            =   3192
      Top             =   264
      Width           =   5772
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Paragraph prints below service charge information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Index           =   0
      Left            =   936
      TabIndex        =   15
      Top             =   4248
      Width           =   4452
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Message Lines:"
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
      Left            =   528
      TabIndex        =   14
      Top             =   2928
      Width           =   2748
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Town Name and Address:"
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
      Height          =   324
      Index           =   5
      Left            =   672
      TabIndex        =   13
      Top             =   1728
      Width           =   2820
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prints on the Top of Return Stub."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Index           =   6
      Left            =   936
      TabIndex        =   12
      Top             =   5808
      Width           =   2484
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prints below Totals Center of Return Stub."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   7
      Left            =   960
      TabIndex        =   11
      Top             =   6456
      Width           =   4524
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Prints on Top Right Section Below Logo and Town Address."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   960
      TabIndex        =   10
      Top             =   3216
      Width           =   4284
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5808
      X2              =   11832
      Y1              =   1392
      Y2              =   1392
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   756
      Left            =   3192
      Top             =   144
      Width           =   5772
   End
End
Attribute VB_Name = "frmBillInfoLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CBeach As Boolean

Private Sub cmdExit_Click()
  DoEvents
  Unload Me
  DoEvents
End Sub

Private Sub cmdSave_Click()
  If ChkStuff Then
    SavebillLTr
    cmdExit_Click
  End If
End Sub

Private Sub cmdTest_Click()
  Dim UBBillSetuplen As Integer
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen

  If UBBillSetup(1).Bill = 16 Then
    SavebillLTr
    TestLtrBillPrint16
  ElseIf UBBillSetup(1).Bill = 19 Then
    SavebillLTr
    TestLtrBillPrint19
  Else
    MsgBox "Test Print only for Letter Bill #16 or #19.", vbOKOnly, "Invalid Bill Type"
  End If
End Sub

'Private Sub fpcboPostalBar_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboPostalBar.ListDown = True
'  End If
'  If fpcboPostalBar.ListDown <> True Then
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
'      fpcboAcctBar.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        fpcboLateNotices.SetFocus
'        KeyCode = 0
'      End If
'    End If
'  End If
'End Sub
'Private Sub fpcboAcctBar_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboAcctBar.ListDown = True
'  End If
'  If fpcboAcctBar.ListDown <> True Then
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
'      fpcboBalType.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        fpcboPostalBar.SetFocus
'        KeyCode = 0
'      End If
'    End If
'  End If
'End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via Billltrinfo setup by " + PWUser$
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
      Call cmdExit_Click
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      Call cmdTest_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdSave_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBBillLetterlen As Integer
  Dim UBBillSetuplen As Integer, cnt As Integer
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
  If InStr(TOWNNAME$, "CAROLINA BEACH") Then
    CBeach = True
  Else
    CBeach = False
  End If

  If UBBillSetup(1).Bill = 19 Then
    For cnt = 0 To 4
      fptxtPgph(cnt).FontName = "Times New Roman"
      fptxtPgph(cnt).FontSize = 11
      fptxtPgph(cnt).Maxlength = 100
    Next
  End If
  fpcboLogo.AddItem "No"
  fpcboLogo.AddItem "Yes"
  fpcboMtrNum.AddItem "1-Meter Serial Number"
  fpcboMtrNum.AddItem "2-Meter ID Number"
  
  If Exist(UBPath$ + "UBBilLtr.DAT") Then
    ReDim UBBillLetter(1) As UBBillLetterType
    UBBillLetterlen = Len(UBBillLetter(1))
    LoadUBBillLetterFile UBBillLetter(), UBBillLetterlen
    fpcboLogo.ListIndex = UBBillLetter(1).IncLogoFlag
    fpcboMtrNum.ListIndex = UBBillLetter(1).MtrNumFlag - 1
    fpTxtHead1 = QPTrim(UBBillLetter(1).BL1Head1)
    fpTxtHead2 = QPTrim(UBBillLetter(1).BL1Head2)
    fptxtHead3 = QPTrim(UBBillLetter(1).BL1Head3)
    fptxtOpt1 = QPTrim(UBBillLetter(1).MsgOpt1)
    fptxtPgph(0) = RTrim(UBBillLetter(1).MsgPgph1)
    fptxtPgph(1) = RTrim(UBBillLetter(1).MsgPgph2)
    fptxtPgph(2) = RTrim(UBBillLetter(1).MsgPgph3)
    fptxtPgph(3) = RTrim(UBBillLetter(1).MsgPgph4)
    fptxtPgph(4) = RTrim(UBBillLetter(1).MsgPgph5)
    fpTxtOpt2 = QPTrim(UBBillLetter(1).MsgOpt2)
    fpTxtOpt3 = QPTrim(UBBillLetter(1).MsgOpt3)
    fpTxtOpt4 = QPTrim(UBBillLetter(1).MsgOpt4)
    fpTxtOpt5 = QPTrim(UBBillLetter(1).MsgOpt5)
  End If
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub SavebillLTr()
  Dim UBBillLetter As UBBillLetterType
  Dim Handle As Integer, UBBillLetterlen As Integer
  UBBillLetterlen = Len(UBBillLetter)
  Handle = FreeFile
  Open UBPath$ + "UBBilLtr.DAT" For Random Shared As Handle Len = UBBillLetterlen
    UBBillLetter.IncLogoFlag = fpcboLogo.ListIndex
    UBBillLetter.MtrNumFlag = fpcboMtrNum.ListIndex + 1
    UBBillLetter.BL1Head1 = QPTrim(fpTxtHead1)
    UBBillLetter.BL1Head2 = QPTrim(fpTxtHead2)
    UBBillLetter.BL1Head3 = QPTrim(fptxtHead3)
    UBBillLetter.MsgOpt1 = QPTrim(fptxtOpt1)
    UBBillLetter.MsgPgph1 = QPTrim(fptxtPgph(0))
    UBBillLetter.MsgPgph2 = QPTrim(fptxtPgph(1))
    UBBillLetter.MsgPgph3 = QPTrim(fptxtPgph(2))
    UBBillLetter.MsgPgph4 = QPTrim(fptxtPgph(3))
    UBBillLetter.MsgPgph5 = QPTrim(fptxtPgph(4))
    UBBillLetter.MsgOpt2 = QPTrim(fpTxtOpt2)
    UBBillLetter.MsgOpt3 = QPTrim(fpTxtOpt3)
    UBBillLetter.MsgOpt4 = QPTrim(fpTxtOpt4)
    UBBillLetter.MsgOpt5 = QPTrim(fpTxtOpt5)
  Put #Handle, 1, UBBillLetter
  Close Handle
  MsgBox "Save Complete", vbOKOnly, "Saved"
End Sub
Private Function ChkStuff()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, OKFlag As Boolean
  MsgText(2) = ""

  If fpcboLogo.ListIndex = -1 Then
    fpcboLogo.ListIndex = 0
  If fpcboMtrNum.ListIndex = -1 Then
    fpcboMtrNum.ListIndex = 0
  End If
  ElseIf Not Len(fpTxtHead1.Text) > 0 Then
    MsgText(3) = "Invalid Heading"
    MsgText(4) = "Please Enter the Heading."
  Else
    OKFlag = True
  End If
  If Not OKFlag Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If
  ChkStuff = OKFlag
End Function

Private Sub fpTxtHead1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtHead2.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      cmdSave.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtHead3.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtHead1.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead3_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpcboLogo.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtHead2.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpcboLogo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboLogo.ListDown = True
  End If
  If fpcboLogo.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboMtrNum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtHead3.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub


Private Sub fpTxtOpt1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtPgph(0).SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpcboMtrNum.SetFocus
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fptxtPgph_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    If Index = 4 Then
      fpTxtOpt2.SetFocus
    Else
      fptxtPgph(Index + 1).SetFocus
    End If
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      If Index = 0 Then
        fptxtOpt1.SetFocus
      Else
        fptxtPgph(Index - 1).SetFocus
      End If
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpTxtOpt2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt3.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtPgph(4).SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt3_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt4.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt2.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt4_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt5.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt3.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt5_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    cmdSave.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt4.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub TestLtrBillPrint16()
  Dim ToPrint As String, UBRptT As Integer, ReportFile As String
  Dim UBBillLetterlen As Integer
  Dim UBBillSetuplen As Integer
  ReportFile$ = UBPath$ + "UBTstbil.PRN"
  UBRptT = FreeFile
  Open ReportFile$ For Output As UBRptT
    ToPrint$ = Using("########", 1)
    ToPrint$ = ToPrint$ + "~" + "10/01/2222" + "~" + "11/01/2222"
    ToPrint$ = ToPrint$ + "~" + "  31"
    ToPrint$ = ToPrint$ + "~123789~1444~1457~13"
    ToPrint$ = ToPrint$ + "~542239~500~1000~500"
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~" + "WATER~12.00"
    ToPrint$ = ToPrint$ + "~" + "SEWER~110.23"
    ToPrint$ = ToPrint$ + "~" + "GARBAGE~3.00"
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ "
    ToPrint$ = ToPrint$ + "~Tax:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 1.5)
    ToPrint$ = ToPrint$ + "~Previous Balance:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 4.5)
    ToPrint$ = ToPrint$ + "~Current:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 153.73)
    ToPrint$ = ToPrint$ + "~ ~ "
    ToPrint$ = ToPrint$ + "~Total Due:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 158.23)
    ToPrint$ = ToPrint$ + "~" + "11/02/2222" + "~" + "11/12/2222"
    ToPrint$ = ToPrint$ + "~This is 1 test message."
    ToPrint$ = ToPrint$ + "~This is 2 test message."
    ToPrint$ = ToPrint$ + "~This is 3 test message."
    ToPrint$ = ToPrint$ + "~This is 4 test message."
    ToPrint$ = ToPrint$ + "~Customer Bill Message Line."
    ToPrint$ = ToPrint$ + "~Draft Message Displayed Here."
    ToPrint$ = ToPrint$ + "~" + "4322" 'Using("##########", )
    ToPrint$ = ToPrint$ + "~" + "2003 NORTH FIRST AVENUE"
    ToPrint$ = ToPrint$ + "~" + "JOHN WILBER ZUCKERMAN"
    ToPrint$ = ToPrint$ + "~" + "2003 NORTH FIRST AVENUE"
    ToPrint$ = ToPrint$ + "~" + "PO BOX 12311"
    ToPrint$ = ToPrint$ + "~" + "ANYWHERE      " + " " + "NC" + " 22233"
    ToPrint$ = ToPrint$ + "~" + Using("#######.##", 173.6)
    ToPrint$ = ToPrint$ + "~" + "22233" + "~01-032234" + "~2223303"
'    If CBeach = True Then
'      ToPrint$ = ToPrint$ + "~" + "110222111222"
'      ToPrint$ = ToPrint$ + "00004322"
'      ToPrint$ = ToPrint$ + "000000000450"
'      ToPrint$ = ToPrint$ + "000000015823"
'      ToPrint$ = ToPrint$ + "0402"
'    Else
      ToPrint$ = ToPrint$ + "~ "
'    End If

    Print #UBRptT, ToPrint$

    Close
    ReDim UBBillSetup(1) As UBBillSetupType
    UBBillSetuplen = Len(UBBillSetup(1))
    LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
    ReDim UBBillLetter(1) As UBBillLetterType
    UBBillLetterlen = Len(UBBillLetter(1))
    LoadUBBillLetterFile UBBillLetter(), UBBillLetterlen
    If UBBillLetter(1).IncLogoFlag = 1 Then
      If Exist(UBPath$ + "UBTNlogo.bmp") Then
        ARptBillLaserLetterForm.Image1.Picture = LoadPicture(UBPath$ + "UBTNlogo.bmp")
        ARptBillLaserLetterForm.Image1.Visible = True
      End If
    End If
    If Not UBBillSetup(1).PostBar = "Y" Then ARptBillLaserLetterForm.Barcode1.Visible = False
    If Not UBBillSetup(1).AcctBar = "Y" Then ARptBillLaserLetterForm.Barcode2.Visible = False
    ARptBillLaserLetterForm.Head1 = QPTrim(UBBillLetter(1).BL1Head1)
    ARptBillLaserLetterForm.Head2 = QPTrim(UBBillLetter(1).BL1Head2)
    ARptBillLaserLetterForm.Head3 = QPTrim(UBBillLetter(1).BL1Head3)
    ARptBillLaserLetterForm.LblHead4 = QPTrim(UBBillLetter(1).BL1Head1)
    ARptBillLaserLetterForm.LblHead5 = QPTrim(UBBillLetter(1).BL1Head2)
    ARptBillLaserLetterForm.LblHead6 = QPTrim(UBBillLetter(1).BL1Head3)
    ARptBillLaserLetterForm.LblOpt1 = QPTrim(UBBillLetter(1).MsgOpt1)
    ARptBillLaserLetterForm.LblOpt2 = QPTrim(UBBillLetter(1).MsgOpt2)
    ARptBillLaserLetterForm.LblOpt3 = QPTrim(UBBillLetter(1).MsgOpt3)
    ARptBillLaserLetterForm.LblOpt4 = QPTrim(UBBillLetter(1).MsgOpt4)
    ARptBillLaserLetterForm.LblOpt5 = QPTrim(UBBillLetter(1).MsgOpt5)
    ARptBillLaserLetterForm.LblPgph1 = QPTrim(UBBillLetter(1).MsgPgph1)
    ARptBillLaserLetterForm.LblPgph2 = QPTrim(UBBillLetter(1).MsgPgph2)
    ARptBillLaserLetterForm.LblPgph3 = QPTrim(UBBillLetter(1).MsgPgph3)
    ARptBillLaserLetterForm.LblPgph4 = QPTrim(UBBillLetter(1).MsgPgph4)
    ARptBillLaserLetterForm.LblPgph5 = QPTrim(UBBillLetter(1).MsgPgph5)
    ARptBillLaserLetterForm.GetName ReportFile$
    ARptBillLaserLetterForm.startrpt

End Sub
Private Sub TestLtrBillPrint19()
  Dim ToPrint As String, UBRptT As Integer, ReportFile As String
  Dim UBBillLetterlen As Integer
  Dim UBBillSetuplen As Integer
  ReportFile$ = UBPath$ + "UBTstbil.PRN"
  UBRptT = FreeFile
  Open ReportFile$ For Output As UBRptT
    ToPrint$ = Using("########", 1)
    ToPrint$ = ToPrint$ + "~" + "10/01/2222" + "~" + "11/01/2222"
    ToPrint$ = ToPrint$ + "~" + "  31"
    ToPrint$ = ToPrint$ + "~123789~1444~1457~13"
    ToPrint$ = ToPrint$ + "~542239~500~1000~500"
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    ToPrint$ = ToPrint$ + "~" + "WATER~WR1~12.00"
    ToPrint$ = ToPrint$ + "~" + "SEWER~SWR~110.23"
    ToPrint$ = ToPrint$ + "~" + "GARBAGE~ ~3.00"
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~" + " ~ ~ "
    ToPrint$ = ToPrint$ + "~"
    ToPrint$ = ToPrint$ + "~ "
    ToPrint$ = ToPrint$ + "~Previous Balance:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 4.5)
    ToPrint$ = ToPrint$ + "~Current:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 125.23)
    ToPrint$ = ToPrint$ + "~ ~ "
    ToPrint$ = ToPrint$ + "~Total Due:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", 132.73)
    ToPrint$ = ToPrint$ + "~" + "11/02/2222" + "~" + "11/12/2222"
    ToPrint$ = ToPrint$ + "~This is 1 test message.50 SPACES EACH LINE"
    ToPrint$ = ToPrint$ + "~This is 2 test message."
    ToPrint$ = ToPrint$ + "~This is 3 test message."
    ToPrint$ = ToPrint$ + "~This is 4 test message."
    ToPrint$ = ToPrint$ + "~Customer Bill Message Line."
    ToPrint$ = ToPrint$ + "~Draft Message Displayed Here."
    ToPrint$ = ToPrint$ + "~" + "4322" 'Using("##########", )
    ToPrint$ = ToPrint$ + "~" + "12180 UNIVERSITY CITY BLVD"
    ToPrint$ = ToPrint$ + "~" + "JOHN WILBER ZUCKERMAN"
    ToPrint$ = ToPrint$ + "~" + "12180 UNIVERSITY CITY BLVD"
    ToPrint$ = ToPrint$ + "~" + ""
    ToPrint$ = ToPrint$ + "~" + "HARRISBURG" + " " + "NC" + " 28075-7406"
    ToPrint$ = ToPrint$ + "~" + Using("#######.##", 137.75)
    ToPrint$ = ToPrint$ + "~" + "28075" + "~01-032234" + "~2223303~11/02/2222"

    Print #UBRptT, ToPrint$

    Close
    ReDim UBBillSetup(1) As UBBillSetupType
    UBBillSetuplen = Len(UBBillSetup(1))
    LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
    ReDim UBBillLetter(1) As UBBillLetterType
    UBBillLetterlen = Len(UBBillLetter(1))
    LoadUBBillLetterFile UBBillLetter(), UBBillLetterlen
    If Not UBBillSetup(1).PostBar = "Y" Then ARptBillLaserLetterPrePrinted.Barcode1.Visible = False
    If Not UBBillSetup(1).AcctBar = "Y" Then ARptBillLaserLetterPrePrinted.Barcode2.Visible = False
    ARptBillLaserLetterPrePrinted.LblPgph1 = QPTrim(UBBillLetter(1).MsgPgph1)
    ARptBillLaserLetterPrePrinted.LblPgph2 = QPTrim(UBBillLetter(1).MsgPgph2)
    ARptBillLaserLetterPrePrinted.LblPgph3 = QPTrim(UBBillLetter(1).MsgPgph3)
    ARptBillLaserLetterPrePrinted.LblPgph4 = QPTrim(UBBillLetter(1).MsgPgph4)
    ARptBillLaserLetterPrePrinted.LblPgph5 = QPTrim(UBBillLetter(1).MsgPgph5)
    ARptBillLaserLetterPrePrinted.GetName ReportFile$
    ARptBillLaserLetterPrePrinted.startrpt

End Sub

