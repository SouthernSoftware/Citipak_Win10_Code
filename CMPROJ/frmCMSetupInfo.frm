VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmCMSetupInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CM Setup Information"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmCMSetupInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAdj 
      Height          =   312
      Left            =   6096
      TabIndex        =   5
      Top             =   5928
      Width           =   1980
      _Version        =   196608
      _ExtentX        =   3492
      _ExtentY        =   550
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
      ColDesigner     =   "frmCMSetupInfo.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVoids 
      Height          =   312
      Left            =   6096
      TabIndex        =   2
      Top             =   4080
      Width           =   1980
      _Version        =   196608
      _ExtentX        =   3492
      _ExtentY        =   550
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
      ColDesigner     =   "frmCMSetupInfo.frx":0C6C
   End
   Begin LpLib.fpCombo fpcboGLInterface 
      Height          =   312
      Left            =   6096
      TabIndex        =   1
      Top             =   3048
      Width           =   828
      _Version        =   196608
      _ExtentX        =   1460
      _ExtentY        =   550
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
      ColDesigner     =   "frmCMSetupInfo.frx":100E
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
      TabIndex        =   8
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
      TabIndex        =   9
      Top             =   7752
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
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
            TextSave        =   "4:52 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "5/31/2005"
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
   Begin EditLib.fpText fptxtCMTownName 
      Height          =   300
      Left            =   4464
      TabIndex        =   0
      Top             =   2520
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   1
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
   Begin EditLib.fpText fptxtAdjConfirm 
      Height          =   324
      Left            =   6096
      TabIndex        =   7
      ToolTipText     =   "No Spaces or Special Characters Allowed."
      Top             =   6816
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      AutoBeep        =   -1  'True
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
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
      Text            =   ""
      CharValidationText=   "~"" `!@#$%^&*()_+-={}|[]\:"";'<>?,./"""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
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
   Begin EditLib.fpText fptxtAdjWord 
      Height          =   324
      Left            =   6096
      TabIndex        =   6
      ToolTipText     =   "No Spaces or Special Characters Allowed."
      Top             =   6384
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      AutoBeep        =   -1  'True
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
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
      Text            =   ""
      CharValidationText=   "~"" `!@#$%^&*()-_=+{}|[]\:"";'<>?,./"""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
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
   Begin EditLib.fpText fptxtVoidConfirm 
      Height          =   324
      Left            =   6096
      TabIndex        =   4
      ToolTipText     =   "No Spaces or Special Characters Allowed."
      Top             =   4968
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      AutoBeep        =   -1  'True
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
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
      Text            =   ""
      CharValidationText=   "~"" `!@#$%^&*()_+-={}|[]\:"";'<>?,./"""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
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
   Begin EditLib.fpText fptxtVoidWord 
      Height          =   324
      Left            =   6096
      TabIndex        =   3
      ToolTipText     =   "No Spaces or Special Characters Allowed."
      Top             =   4524
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      AutoBeep        =   -1  'True
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
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
      Text            =   ""
      CharValidationText=   "~"" `!@#$%^&*()-_=+{}|[]\:"";'<>?,./"""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password Requirement:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   3216
      TabIndex        =   19
      Top             =   5952
      Width           =   2604
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password Requirement:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   3168
      TabIndex        =   18
      Top             =   4104
      Width           =   2652
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   4752
      TabIndex        =   17
      Top             =   4536
      Width           =   1068
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   3600
      TabIndex        =   16
      Top             =   4992
      Width           =   2220
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   4536
      TabIndex        =   15
      Top             =   6396
      Width           =   1284
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   3576
      TabIndex        =   14
      Top             =   6864
      Width           =   2244
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   5316
      Left            =   2508
      Top             =   2016
      Width           =   7212
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interface with General Ledger:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   2808
      TabIndex        =   13
      Top             =   3072
      Width           =   3156
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Management Setup Information"
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
      Left            =   3222
      TabIndex        =   12
      Top             =   1200
      Width           =   5772
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   636
      Left            =   3222
      Top             =   1056
      Width           =   5772
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Town Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   5
      Left            =   2856
      TabIndex        =   11
      Top             =   2544
      Width           =   1476
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   756
      Left            =   3222
      Top             =   936
      Width           =   5772
   End
   Begin VB.Line Line1 
      X1              =   2544
      X2              =   9696
      Y1              =   3624
      Y2              =   3624
   End
   Begin VB.Line Line2 
      X1              =   2544
      X2              =   9720
      Y1              =   5472
      Y2              =   5472
   End
   Begin VB.Label Label3 
      Caption         =   "  ADJUSTMENTS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Top             =   5496
      Width           =   7188
   End
   Begin VB.Label Label3 
      Caption         =   "  VOIDS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2520
      TabIndex        =   21
      Top             =   3648
      Width           =   7188
   End
End
Attribute VB_Name = "frmCMSetupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Pz As String, Z As String

Private Sub cmdExit_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  FntSize = frmMsgDialog.Label(1).FontSize
  If ChkifChange Then
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:Changes Have Been Made"
    MsgText(1) = ""
    MsgText(2) = "Do You Want to Abandon changes?"
    MsgText(3) = "Ok to Abandon,"
    MsgText(4) = "Cancel to Remain on screen."
    MsgText(5) = ""
    If GetOKorNot(MsgText()) Then
     CMLog "USER WANTS TO Exit"
     Load frmCMSetupMenu
     DoEvents
     frmCMSetupMenu.Show
     Unload Me
     DoEvents
    Else
     CMLog "USER Canceled"
    End If
  Else
    Load frmCMSetupMenu
    DoEvents
    frmCMSetupMenu.Show
    Unload Me
    DoEvents
  End If
End Sub

Private Sub cmdSave_Click()
  If ChkStuff Then
    Savesetup
    Load frmCMSetupMenu
    DoEvents
    frmCMSetupMenu.Show
    Unload Me
    DoEvents
  End If
End Sub

Private Sub fpcboGLInterface_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboGLInterface.ListDown = True
  End If
  If fpcboGLInterface.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboVoids.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtCMTownName.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboVoids_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVoids.ListDown = True
  End If
  If fpcboVoids.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fptxtVoidWord.Enabled = True Then
        fptxtVoidWord.SetFocus
      Else
        fpcboAdj.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboGLInterface.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboAdj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAdj.ListDown = True
  End If
  If fpcboAdj.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fptxtAdjWord.Enabled = True Then
        fptxtAdjWord.SetFocus
      Else
        cmdSave.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If fptxtVoidConfirm.Enabled = True Then
          fptxtVoidConfirm.SetFocus
        Else
          fpcboVoids.SetFocus
        End If
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcboVoids_SelChange(ItemIndex As Long)
  If fpcboVoids.ListIndex = 1 Then
    fptxtVoidWord.Enabled = True
    fptxtVoidConfirm.Enabled = True
  Else
    fptxtVoidWord.Enabled = False
    fptxtVoidConfirm.Enabled = False
  End If
End Sub
Private Sub fpcboAdj_SelChange(ItemIndex As Long)
  If fpcboAdj.ListIndex = 1 Then
    fptxtAdjWord.Enabled = True
    fptxtAdjConfirm.Enabled = True
  Else
    fptxtAdjWord.Enabled = False
    fptxtAdjConfirm.Enabled = False
  End If
End Sub

Private Sub fptxtCMTownName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     fpcboGLInterface.SetFocus
  End If
End Sub
Private Sub fptxtVoidWord_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtVoidConfirm.SetFocus
  End If
End Sub
Private Sub fptxtAdjWord_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtAdjConfirm.SetFocus
  End If
End Sub

Private Sub fptxtVoidConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fptxtAdjConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
    KeyCode = 0
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
        CMLog "Closed via CMsetup by " + PWUser$
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
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdSave_Click
    Case Else:
  End Select
End Sub
Private Function ChkifChange()
  Dim CMSetuplen As Integer, chg As Integer, cnt As Integer
  chg = 0
  If Exist(UBPath$ + "CMSetTown.DAT") Then
    ReDim CMSetup(1) As CMSetupType
    CMSetuplen = Len(CMSetup(1))
    LoadCMSetUpFile CMSetup(), CMSetuplen
    If QPTrim(fptxtCMTownName) <> QPTrim(CMSetup(1).CMTOWNNAME) Then chg = chg + 1
    If QPTrim(CMSetup(1).GLInterface) = "Y" Then
      If fpcboGLInterface.ListIndex <> 0 Then chg = chg + 1
    Else
      If fpcboGLInterface.ListIndex <> 1 Then chg = chg + 1
    End If
    
    If QPTrim(CMSetup(1).Pass4Voids) = "Y" Then
      If fpcboVoids.ListIndex <> 1 Then chg = chg + 1
    ElseIf QPTrim(CMSetup(1).Pass4Voids) = "F" Then
       If fpcboVoids.ListIndex <> 2 Then chg = chg + 1
    Else
      If fpcboVoids.ListIndex <> 0 Then chg = chg + 1
    End If
    If QPTrim(CMSetup(1).Pass4Adj) = "Y" Then
      If fpcboAdj.ListIndex <> 1 Then chg = chg + 1
    ElseIf QPTrim(CMSetup(1).Pass4Adj) = "F" Then
      If fpcboAdj.ListIndex <> 2 Then chg = chg + 1
    Else
      If fpcboAdj.ListIndex <> 0 Then chg = chg + 1
    End If
    If Len(CMSetup(1).VoidPW) > 0 Then
      Pz$ = ""
      Z$ = QPTrim(CMSetup(1).VoidPW)
      For cnt = 1 To Len(Z$)
        Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
      Next
      If QPTrim$(fptxtVoidWord) <> Pz$ Then chg = chg + 1
    Else
      If QPTrim$(fptxtVoidWord) <> QPTrim(CMSetup(1).VoidPW) Then chg = chg + 1
    End If
    If Len(CMSetup(1).AdjPW) > 0 Then
      Pz$ = ""
      Z$ = QPTrim(CMSetup(1).AdjPW)
      For cnt = 1 To Len(Z$)
        Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
      Next
      If QPTrim$(fptxtAdjWord) <> Pz$ Then chg = chg + 1
    Else
      If QPTrim$(fptxtAdjWord) <> QPTrim(CMSetup(1).AdjPW) Then chg = chg + 1
    End If
  Else
    If Len(QPTrim(fptxtCMTownName)) > 0 Then chg = chg + 1
    If fpcboGLInterface.ListIndex <> 1 Then chg = chg + 1
    If fpcboVoids.ListIndex <> 0 Then chg = chg + 1
    If fpcboAdj.ListIndex <> 0 Then chg = chg + 1
    If Len(QPTrim$(fptxtVoidWord)) > 0 Then chg = chg + 1
    If Len(QPTrim$(fptxtAdjWord)) > 0 Then chg = chg + 1
  End If
  If chg > 0 Then
    ChkifChange = True
  End If
End Function
Private Sub Form_Load()
  Dim CMSetuplen As Integer, cnt As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  fpcboGLInterface.AddItem "Yes"
  fpcboGLInterface.AddItem "No"
  fpcboVoids.AddItem "None"
  fpcboVoids.AddItem "Yes"
  fpcboVoids.AddItem "Full Access Only"
  fpcboAdj.AddItem "None"
  fpcboAdj.AddItem "Yes"
  fpcboAdj.AddItem "Full Access Only"
  If Exist(UBPath$ + "CMSetTown.DAT") Then
    ReDim CMSetup(1) As CMSetupType
    CMSetuplen = Len(CMSetup(1))
    LoadCMSetUpFile CMSetup(), CMSetuplen
    If QPTrim(CMSetup(1).GLInterface) = "Y" Then
      fpcboGLInterface.ListIndex = 0
    Else
      fpcboGLInterface.ListIndex = 1
    End If
    fptxtCMTownName = QPTrim(CMSetup(1).CMTOWNNAME)
    If QPTrim(CMSetup(1).Pass4Voids) = "Y" Then
      fpcboVoids.ListIndex = 1
    ElseIf QPTrim(CMSetup(1).Pass4Voids) = "F" Then
      fpcboVoids.ListIndex = 2
    Else
      fpcboVoids.ListIndex = 0
    End If
    If QPTrim(CMSetup(1).Pass4Adj) = "Y" Then
      fpcboAdj.ListIndex = 1
    ElseIf QPTrim(CMSetup(1).Pass4Adj) = "F" Then
      fpcboAdj.ListIndex = 2
    Else
      fpcboAdj.ListIndex = 0
    End If
    If Len(CMSetup(1).VoidPW) > 0 Then
      Pz$ = ""
      Z$ = QPTrim(CMSetup(1).VoidPW)
      For cnt = 1 To Len(Z$)
        Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
      Next
      fptxtVoidWord = Pz$
      fptxtVoidConfirm = Pz$
    End If
    If Len(CMSetup(1).AdjPW) > 0 Then
      Pz$ = ""
      Z$ = QPTrim(CMSetup(1).AdjPW)
      For cnt = 1 To Len(Z$)
        Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
      Next
      fptxtAdjWord = Pz$
      fptxtAdjConfirm = Pz$
    End If
  If fpcboVoids.ListIndex = 1 Then
    fptxtVoidWord.Enabled = True
    fptxtVoidConfirm.Enabled = True
  Else
    fptxtVoidWord.Enabled = False
    fptxtVoidConfirm.Enabled = False
  End If
  If fpcboAdj.ListIndex = 1 Then
    fptxtAdjWord.Enabled = True
    fptxtAdjConfirm.Enabled = True
  Else
    fptxtAdjWord.Enabled = False
    fptxtAdjConfirm.Enabled = False
  End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub Savesetup()
  Dim CMSetupFile As CMSetupType
  Dim Handle As Integer, CMSetuplen As Integer, cnt As Integer
  CMSetuplen = Len(CMSetupFile)
  Handle = FreeFile
  Open UBPath$ + "CMSetTown.dat" For Random Shared As Handle Len = CMSetuplen
    If fpcboGLInterface.ListIndex = 0 Then
      CMSetupFile.GLInterface = "Y"
    Else
      CMSetupFile.GLInterface = "N"
    End If
    CMSetupFile.CMTOWNNAME = QPTrim(fptxtCMTownName.Text)
    If fpcboVoids.ListIndex = 1 Then
      CMSetupFile.Pass4Voids = "Y"
    ElseIf fpcboVoids.ListIndex = 2 Then
      CMSetupFile.Pass4Voids = "F"
    Else
      CMSetupFile.Pass4Voids = "N"
    End If
    If fpcboAdj.ListIndex = 1 Then
      CMSetupFile.Pass4Adj = "Y"
    ElseIf fpcboAdj.ListIndex = 2 Then
      CMSetupFile.Pass4Adj = "F"
    Else
      CMSetupFile.Pass4Adj = "N"
    End If
    If fpcboVoids.ListIndex = 1 Then
    If Len(QPTrim$(fptxtVoidWord)) > 0 Then
      Pz$ = ""
      Z$ = QPTrim(fptxtVoidWord)
      For cnt = 1 To Len(Z$)
        Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
      Next
      CMSetupFile.VoidPW = Pz$
    End If
    Else
      CMSetupFile.VoidPW = ""
    End If
    If fpcboAdj.ListIndex = 1 Then
    If Len(QPTrim(fptxtAdjWord)) > 0 Then
      Pz$ = ""
      Z$ = QPTrim(fptxtAdjWord)
      For cnt = 1 To Len(Z$)
        Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
      Next
      CMSetupFile.AdjPW = Pz$
    End If
    Else
      CMSetupFile.AdjPW = ""
    End If

    
  Put #Handle, 1, CMSetupFile
  Close Handle
  MsgBox "Save Complete", vbOKOnly, "Saved"
End Sub
Private Function ChkStuff()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, OKFlag As Boolean
  MsgText(2) = ""
  MsgText(3) = ""
  If Len(QPTrim(fptxtCMTownName)) = 0 Then
    MsgText(2) = "You Must Enter A Town Name."
    OKFlag = False
  ElseIf fpcboGLInterface.ListIndex = -1 Then
    MsgText(3) = "You Must Select An Interface Option."
    OKFlag = False
  Else
    OKFlag = True
  End If
  If fpcboVoids.ListIndex = 1 Then
    If Len(QPTrim$(fptxtVoidWord)) > 0 Then
      If QPTrim(fptxtVoidWord) <> QPTrim(fptxtVoidConfirm) Then
        OKFlag = False
        MsgText(2) = "Confirm Password does not match Password"
        MsgText(3) = "Please Try Again."
      Else
        OKFlag = True
      End If
    Else
      OKFlag = False
      MsgText(2) = "Password may not be blank."
      MsgText(3) = "Please Try Again."
    End If
  End If
  If fpcboAdj.ListIndex = 1 Then
    If Len(QPTrim$(fptxtAdjWord)) > 0 Then
      If QPTrim(fptxtAdjWord) <> QPTrim(fptxtAdjConfirm) Then
        OKFlag = False
        MsgText(2) = "Confirm Password does not match Password"
        MsgText(3) = "Please Try Again."
      Else
        OKFlag = True
      End If
    Else
      OKFlag = False
      MsgText(2) = "Password may not be blank."
      MsgText(3) = "Please Try Again."
    End If
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

Private Sub fptxtCMTownName_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

