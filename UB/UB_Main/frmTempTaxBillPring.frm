VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmTempTaxBillPrint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Letter Printing"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmTempTaxBillPring.frx":0000
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
      Left            =   3075
      TabIndex        =   2
      Top             =   2250
      Width           =   825
      _Version        =   196608
      _ExtentX        =   1455
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
      ColDesigner     =   "frmTempTaxBillPring.frx":08CA
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
      Left            =   1008
      TabIndex        =   21
      Top             =   7752
      Width           =   1644
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
      Left            =   8424
      TabIndex        =   22
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
      TabIndex        =   23
      Top             =   7752
      Width           =   1332
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
            TextSave        =   "4:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/1/2007"
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
      Left            =   1008
      TabIndex        =   0
      Top             =   1560
      Width           =   5916
      _Version        =   196608
      _ExtentX        =   10435
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
      Text            =   "Town of Spruce Pine, North Carolina"
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
   Begin EditLib.fpText fptxtHead4 
      Height          =   252
      Left            =   3048
      TabIndex        =   17
      Top             =   6264
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   444
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
      Text            =   "Town of Spruce Pine"
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
      Left            =   1008
      TabIndex        =   1
      Top             =   1872
      Width           =   5916
      _Version        =   196608
      _ExtentX        =   10435
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
      Text            =   "2004 Property Tax Notice"
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
   Begin EditLib.fpText fpTxtOpt7 
      Height          =   252
      Left            =   1008
      TabIndex        =   20
      Top             =   7200
      Width           =   6468
      _Version        =   196608
      _ExtentX        =   11409
      _ExtentY        =   444
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
      Text            =   "If a receipt is desired, send a self-addressed stamped envelope."
      CharValidationText=   ""
      MaxLength       =   75
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
      Height          =   252
      Left            =   1032
      TabIndex        =   15
      Top             =   5376
      Width           =   6468
      _Version        =   196608
      _ExtentX        =   11409
      _ExtentY        =   444
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
      Text            =   "Please detach and return this portion with your payment."
      CharValidationText=   ""
      MaxLength       =   75
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
   Begin EditLib.fpText fpTxtOpt6 
      Height          =   252
      Left            =   1008
      TabIndex        =   16
      Top             =   5928
      Width           =   4476
      _Version        =   196608
      _ExtentX        =   7895
      _ExtentY        =   444
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
      Text            =   "Make check payable to:"
      CharValidationText=   ""
      MaxLength       =   45
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
      Height          =   252
      Left            =   7200
      TabIndex        =   3
      Top             =   1392
      Width           =   4500
      _Version        =   196608
      _ExtentX        =   7937
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Text            =   "Questions? Contact us by-"
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
      Height          =   252
      Left            =   7200
      TabIndex        =   4
      Top             =   1680
      Width           =   4500
      _Version        =   196608
      _ExtentX        =   7937
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Text            =   "Phone:   828-765-3000"
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
      Height          =   228
      Index           =   0
      Left            =   1008
      TabIndex        =   7
      Top             =   3192
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Text            =   "Taxes are due and payable upon receipt of notice.  Interest begins to accrue on unpaid taxes at the rate of 2% on January"
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
      Height          =   228
      Index           =   1
      Left            =   1008
      TabIndex        =   8
      Top             =   3432
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Text            =   "6th.  Beginning on the first day of February and continuing thereafter,  interest accrues at the rate of 0.75 percent per"
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
      Height          =   228
      Index           =   2
      Left            =   1008
      TabIndex        =   9
      Top             =   3672
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Text            =   "month, until all taxes, interest and penalties are paid.  "
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
      Height          =   228
      Index           =   3
      Left            =   1008
      TabIndex        =   10
      Top             =   3912
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Height          =   228
      Index           =   4
      Left            =   1008
      TabIndex        =   11
      Top             =   4152
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Text            =   "Of the current tax rate of $0.43 per hundred dollars valuation, $0.06 is earmarked for the Spruce Pine Volunteer Fire"
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
   Begin EditLib.fpText fpTxtOpt3 
      Height          =   252
      Left            =   7200
      TabIndex        =   5
      Top             =   1968
      Width           =   4500
      _Version        =   196608
      _ExtentX        =   7937
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Text            =   "Fax:     828-765-3014"
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
   Begin EditLib.fpText fpTxtOpt4 
      Height          =   252
      Left            =   7200
      TabIndex        =   6
      Top             =   2256
      Width           =   4500
      _Version        =   196608
      _ExtentX        =   7937
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Text            =   "e-mail:  spmgr@bellsouth.net"
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
      Height          =   228
      Index           =   5
      Left            =   1008
      TabIndex        =   12
      Top             =   4392
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Text            =   "Department for contractual fire protection within the town.  "
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
      Height          =   228
      Index           =   6
      Left            =   1008
      TabIndex        =   13
      Top             =   4632
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
      Height          =   228
      Index           =   7
      Left            =   1008
      TabIndex        =   14
      Top             =   4872
      Width           =   10380
      _Version        =   196608
      _ExtentX        =   18309
      _ExtentY        =   402
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
   Begin EditLib.fpText fptxtHead5 
      Height          =   252
      Left            =   3048
      TabIndex        =   18
      Top             =   6528
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   444
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
      Text            =   "P O Box 189"
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
   Begin EditLib.fpText fptxtHead6 
      Height          =   252
      Left            =   3048
      TabIndex        =   19
      Top             =   6792
      Width           =   4716
      _Version        =   196608
      _ExtentX        =   8318
      _ExtentY        =   444
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
      Text            =   "Spruce Pine, NC 28777"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   924
      Left            =   2688
      TabIndex        =   36
      Top             =   6216
      Width           =   348
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return Address:"
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
      Index           =   0
      Left            =   816
      TabIndex        =   35
      Top             =   6504
      Width           =   1740
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
      Left            =   528
      TabIndex        =   34
      Top             =   2280
      Width           =   2436
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(8 Lines Available)"
      Height          =   276
      Left            =   5016
      TabIndex        =   33
      Top             =   2976
      Width           =   1404
   End
   Begin VB.Line Line4 
      X1              =   11832
      X2              =   11832
      Y1              =   1152
      Y2              =   7032
   End
   Begin VB.Line Line3 
      X1              =   576
      X2              =   11856
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   672
      X2              =   672
      Y1              =   1104
      Y2              =   6960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Letter Laser Bill Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   32
      Top             =   936
      Width           =   4860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Letter Printing"
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
      Left            =   3624
      TabIndex        =   31
      Top             =   336
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   516
      Left            =   3192
      Top             =   240
      Width           =   5772
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Paragraph prints below Property Charge Information:"
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
      TabIndex        =   30
      Top             =   2952
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
      Left            =   552
      TabIndex        =   29
      Top             =   2616
      Width           =   2748
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Heading Lines 1 and 2:"
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
      Left            =   576
      TabIndex        =   28
      Top             =   1224
      Width           =   2388
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
      TabIndex        =   27
      Top             =   5160
      Width           =   2484
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prints below Totals Left  Side of Return Stub."
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
      TabIndex        =   26
      Top             =   5664
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
      Left            =   6888
      TabIndex        =   25
      Top             =   1152
      Width           =   4284
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4488
      X2              =   11664
      Y1              =   1104
      Y2              =   1104
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   636
      Left            =   3192
      Top             =   120
      Width           =   5772
   End
End
Attribute VB_Name = "frmTempTaxBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Dumfriesflag As Boolean
Private Sub cmdExit_Click()
  Load frmUBSetupMenu
  DoEvents
  frmUBSetupMenu.Show
  Unload Me
  DoEvents
End Sub


Private Sub cmdPrint_Click()
  
  SaveTxDef
  
  PrintBills
End Sub
Private Sub PrintBills()
  Dim ToPrint As String, UBRptT As Integer, ReportFile As String
  'Dim UBBillLetterlen As Integer
  'Dim UBBillSetuplen As Integer
  ReportFile$ = UBPath$ + "TaxBil.PRN"
'    UBRptT = FreeFile
'    Open ReportFile$ For Output As UBRptT
'      ToPrint$ = "2004-1505"
'      ToPrint$ = ToPrint$ + "~" + "JOHN SMITH"
'      ToPrint$ = ToPrint$ + "~" + "P O BOX 1520"
'      ToPrint$ = ToPrint$ + "~337 HWY NORTH"
'      ToPrint$ = ToPrint$ + "~HENDERSONVILLE NC 28793"
'      ToPrint$ = ToPrint$ + "~001079847"
'      ToPrint$ = ToPrint$ + "~0798-00-49-7585"
'      ToPrint$ = ToPrint$ + "~2004 DISCOVERY"
'      ToPrint$ = ToPrint$ + "~105,200"
'      ToPrint$ = ToPrint$ + "~0"
'      ToPrint$ = ToPrint$ + "~0"
'      ToPrint$ = ToPrint$ + "~105,200"
'      ToPrint$ = ToPrint$ + "~$0.43"
'      ToPrint$ = ToPrint$ + "~$452.36"
'      Print #UBRptT, ToPrint$
'
'      Close
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmTempTaxBillPrint
    If Not Dumfriesflag Then
      If fpcboLogo.ListIndex = 1 Then
        If Exist(UBPath$ + "txTNlogo.bmp") Then
          ARptTempTaxBill.Image1.Picture = LoadPicture(UBPath$ + "txTNlogo.bmp")
          ARptTempTaxBill.Image1.Visible = True
        End If
      End If
      ARptTempTaxBill.Head1 = QPTrim(fpTxtHead1)
      ARptTempTaxBill.Head2 = QPTrim(fpTxtHead2)
      ARptTempTaxBill.LblOpt1 = QPTrim(fptxtOpt1)
      ARptTempTaxBill.LblOpt2 = QPTrim(fpTxtOpt2)
      ARptTempTaxBill.LblOpt3 = QPTrim(fpTxtOpt3)
      ARptTempTaxBill.LblOpt4 = QPTrim(fpTxtOpt4)
      ARptTempTaxBill.LblPgph1 = QPTrim(fptxtPgph(0))
      ARptTempTaxBill.LblPgph2 = QPTrim(fptxtPgph(1))
      ARptTempTaxBill.LblPgph3 = QPTrim(fptxtPgph(2))
      ARptTempTaxBill.LblPgph4 = QPTrim(fptxtPgph(3))
      ARptTempTaxBill.LblPgph5 = QPTrim(fptxtPgph(4))
      ARptTempTaxBill.LblPgph6 = QPTrim(fptxtPgph(5))
      ARptTempTaxBill.LblPgph7 = QPTrim(fptxtPgph(6))
      ARptTempTaxBill.LblPgph8 = QPTrim(fptxtPgph(7))
      ARptTempTaxBill.LblOpt5 = QPTrim(fpTxtOpt5)
      ARptTempTaxBill.LblHead4 = QPTrim(fptxtHead4)
      ARptTempTaxBill.LblHead5 = QPTrim(fptxtHead5)
      ARptTempTaxBill.LblHead6 = QPTrim(fptxtHead6)
      ARptTempTaxBill.LblOpt6 = QPTrim(fpTxtOpt6)
      ARptTempTaxBill.LblOpt7 = QPTrim(fpTxtOpt7)
      ARptTempTaxBill.GetName ReportFile$
      ARptTempTaxBill.startrpt
   Else
      If fpcboLogo.ListIndex = 1 Then
        If Exist(UBPath$ + "txTNlogo.bmp") Then
          ARptTempTaxBillDum.Image1.Picture = LoadPicture(UBPath$ + "txTNlogo.bmp")
          ARptTempTaxBillDum.Image1.Visible = True
        End If
      End If
      ARptTempTaxBillDum.Head1 = QPTrim(fpTxtHead1)
      ARptTempTaxBillDum.Head2 = QPTrim(fpTxtHead2)
      ARptTempTaxBillDum.LblOpt1 = QPTrim(fptxtOpt1)
      ARptTempTaxBillDum.LblOpt2 = QPTrim(fpTxtOpt2)
      ARptTempTaxBillDum.LblOpt3 = QPTrim(fpTxtOpt3)
      ARptTempTaxBillDum.LblOpt4 = QPTrim(fpTxtOpt4)
      ARptTempTaxBillDum.LblPgph1 = QPTrim(fptxtPgph(0))
      ARptTempTaxBillDum.LblPgph2 = QPTrim(fptxtPgph(1))
      ARptTempTaxBillDum.LblPgph3 = QPTrim(fptxtPgph(2))
      ARptTempTaxBillDum.LblPgph4 = QPTrim(fptxtPgph(3))
      ARptTempTaxBillDum.LblPgph5 = QPTrim(fptxtPgph(4))
      ARptTempTaxBillDum.LblPgph6 = QPTrim(fptxtPgph(5))
      ARptTempTaxBillDum.LblPgph7 = QPTrim(fptxtPgph(6))
      ARptTempTaxBillDum.LblPgph8 = QPTrim(fptxtPgph(7))
      ARptTempTaxBillDum.LblOpt5 = QPTrim(fpTxtOpt5)
      ARptTempTaxBillDum.LblHead4 = QPTrim(fptxtHead4)
      ARptTempTaxBillDum.LblHead5 = QPTrim(fptxtHead5)
      ARptTempTaxBillDum.LblHead6 = QPTrim(fptxtHead6)
      ARptTempTaxBillDum.LblOpt6 = QPTrim(fpTxtOpt6)
      ARptTempTaxBillDum.LblOpt7 = QPTrim(fpTxtOpt7)
      ARptTempTaxBillDum.GetName ReportFile$
      ARptTempTaxBillDum.startrpt

   End If
  
End Sub

Private Sub cmdTest_Click()
'  If fpcboBills.ListIndex = 0 Then
    'SavebillLTr
    SaveTxDef
    TestLtrBillPrint
  
'  Else
'    MsgBox "Test Print only for Laser Bill #1.", vbOKOnly, "Invalid Bill Type"
'  End If
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
        UBLog "Closed via TempTaxBill print by " + PWUser$
        CitiTerminate
      End If
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
      KeyCode = 0
      DoEvents
      Call cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdPrint_Click
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      Call cmdTest_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim UBTxBilDeflen As Integer, Handle As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  If InStr(UCase(TOWNNAME$), "DUMFRIES") Then
    Dumfriesflag = True
  Else
    Dumfriesflag = False
  End If
  fpcboLogo.AddItem "No"
  fpcboLogo.AddItem "Yes"
  If Exist(UBPath$ + "UBTaxBil.DAT") Then
    ReDim UBTxBilDef(1) As TxBillDefaultsType
    UBTxBilDeflen = Len(UBTxBilDef(1))            'use the length as an error flag
    Handle = FreeFile
    Open UBPath$ + "UBTaxBil.DAT" For Random Shared As Handle Len = UBTxBilDeflen    'open data file
    Get #Handle, 1, UBTxBilDef(1)
    Close Handle
    fpcboLogo.ListIndex = UBTxBilDef(1).dologo
    fpTxtHead1 = QPTrim(UBTxBilDef(1).TxtHead1)
    fpTxtHead2 = QPTrim(UBTxBilDef(1).TxtHead2)
    fptxtOpt1 = QPTrim(UBTxBilDef(1).txtOpt1)
    fpTxtOpt2 = QPTrim(UBTxBilDef(1).TxtOpt2)
    fpTxtOpt3 = QPTrim(UBTxBilDef(1).TxtOpt3)
    fpTxtOpt4 = QPTrim(UBTxBilDef(1).TxtOpt4)
    fptxtPgph(0) = QPTrim(UBTxBilDef(1).txtPgph0)
    fptxtPgph(1) = QPTrim(UBTxBilDef(1).txtPgph1)
    fptxtPgph(2) = QPTrim(UBTxBilDef(1).txtPgph2)
    fptxtPgph(3) = QPTrim(UBTxBilDef(1).txtPgph3)
    fptxtPgph(4) = QPTrim(UBTxBilDef(1).txtPgph4)
    fptxtPgph(5) = QPTrim(UBTxBilDef(1).txtPgph5)
    fptxtPgph(6) = QPTrim(UBTxBilDef(1).txtPgph6)
    fptxtPgph(7) = QPTrim(UBTxBilDef(1).txtPgph7)
    fpTxtOpt5 = QPTrim(UBTxBilDef(1).TxtOpt5)
    fptxtHead4 = QPTrim(UBTxBilDef(1).txtHead4)
    fptxtHead5 = QPTrim(UBTxBilDef(1).txtHead5)
    fptxtHead6 = QPTrim(UBTxBilDef(1).txtHead6)
    fpTxtOpt6 = QPTrim(UBTxBilDef(1).TxtOpt6)
    fpTxtOpt7 = QPTrim(UBTxBilDef(1).TxtOpt7)
   
  End If

End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub


Private Sub fpTxtHead1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtHead2.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      cmdPrint.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtOpt1.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtHead1.SetFocus
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
      fptxtOpt1.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpTxtHead2.SetFocus
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
    fpTxtOpt2.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpcboLogo.SetFocus
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fptxtPgph_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    If Index = 7 Then
      fpTxtOpt5.SetFocus
    Else
      fptxtPgph(Index + 1).SetFocus
    End If
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      If Index = 0 Then
        fpTxtOpt4.SetFocus
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
      fptxtOpt1.SetFocus
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
    fptxtPgph(0).SetFocus
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
    fpTxtOpt6.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtPgph(7).SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt6_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt6_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtHead4.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt5.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead4_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtHead5.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fpTxtOpt6.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead5_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fptxtHead6.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtHead4.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtHead6_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtHead6_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpTxtOpt7.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtHead5.SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpTxtOpt7_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTxtOpt7_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    cmdPrint.SetFocus
    KeyCode = 0
  Else
    If KeyCode = vbKeyUp Then
      fptxtHead6.SetFocus
      KeyCode = 0
    End If
  End If
End Sub

Private Sub TestLtrBillPrint()
  Dim ToPrint As String, UBRptT As Integer, ReportFile As String
  'Dim UBBillLetterlen As Integer
  'Dim UBBillSetuplen As Integer
  
  ReportFile$ = UBPath$ + "UBTstbil.PRN"
  UBRptT = FreeFile
  If Not Dumfriesflag Then
    Open ReportFile$ For Output As UBRptT
      ToPrint$ = "2004-1505"
      ToPrint$ = ToPrint$ + "~" + "JOHN SMITH"
      ToPrint$ = ToPrint$ + "~" + "P O BOX 1520"
      ToPrint$ = ToPrint$ + "~337 HWY NORTH"
      ToPrint$ = ToPrint$ + "~HENDERSONVILLE NC 28793"
      ToPrint$ = ToPrint$ + "~001079847"
      ToPrint$ = ToPrint$ + "~0798-00-49-7585"
      ToPrint$ = ToPrint$ + "~2004 DISCOVERY"
      ToPrint$ = ToPrint$ + "~105,200"
      ToPrint$ = ToPrint$ + "~0"
      ToPrint$ = ToPrint$ + "~0"
      ToPrint$ = ToPrint$ + "~105,200"
      ToPrint$ = ToPrint$ + "~$0.43"
      ToPrint$ = ToPrint$ + "~$452.36"
      Print #UBRptT, ToPrint$
  
      Close
      If fpcboLogo.ListIndex = 1 Then
        If Exist(UBPath$ + "txTNlogo.bmp") Then
          ARptTempTaxBill.Image1.Picture = LoadPicture(UBPath$ + "txTNlogo.bmp")
          ARptTempTaxBill.Image1.Visible = True
        End If
      End If
      ARptTempTaxBill.Head1 = QPTrim(fpTxtHead1)
      ARptTempTaxBill.Head2 = QPTrim(fpTxtHead2)
      ARptTempTaxBill.LblOpt1 = QPTrim(fptxtOpt1)
      ARptTempTaxBill.LblOpt2 = QPTrim(fpTxtOpt2)
      ARptTempTaxBill.LblOpt3 = QPTrim(fpTxtOpt3)
      ARptTempTaxBill.LblOpt4 = QPTrim(fpTxtOpt4)
      ARptTempTaxBill.LblPgph1 = QPTrim(fptxtPgph(0))
      ARptTempTaxBill.LblPgph2 = QPTrim(fptxtPgph(1))
      ARptTempTaxBill.LblPgph3 = QPTrim(fptxtPgph(2))
      ARptTempTaxBill.LblPgph4 = QPTrim(fptxtPgph(3))
      ARptTempTaxBill.LblPgph5 = QPTrim(fptxtPgph(4))
      ARptTempTaxBill.LblPgph6 = QPTrim(fptxtPgph(5))
      ARptTempTaxBill.LblPgph7 = QPTrim(fptxtPgph(6))
      ARptTempTaxBill.LblPgph8 = QPTrim(fptxtPgph(7))
      ARptTempTaxBill.LblOpt5 = QPTrim(fpTxtOpt5)
      ARptTempTaxBill.LblHead4 = QPTrim(fptxtHead4)
      ARptTempTaxBill.LblHead5 = QPTrim(fptxtHead5)
      ARptTempTaxBill.LblHead6 = QPTrim(fptxtHead6)
      ARptTempTaxBill.LblOpt6 = QPTrim(fpTxtOpt6)
      ARptTempTaxBill.LblOpt7 = QPTrim(fpTxtOpt7)
      ARptTempTaxBill.GetName ReportFile$
      ARptTempTaxBill.startrpt
  Else
    Open ReportFile$ For Output As UBRptT
    ToPrint$ = "2006-1505"
    ToPrint$ = ToPrint$ + "~11111~" + "JOHN SMITH"
    ToPrint$ = ToPrint$ + "~" + "P O BOX 1520"
    ToPrint$ = ToPrint$ + "~337 HWY NORTH"
    ToPrint$ = ToPrint$ + "~ANYWHERE VA 28793"
    ToPrint$ = ToPrint$ + "~Description of Property"
    ToPrint$ = ToPrint$ + "~0798-00-49-7585"
    ToPrint$ = ToPrint$ + "~Other description"
    ToPrint$ = ToPrint$ + "~0.1800"
    ToPrint$ = ToPrint$ + "~228,600"
    ToPrint$ = ToPrint$ + "~157,900"
    ToPrint$ = ToPrint$ + "~386,400"
    ToPrint$ = ToPrint$ + "~198.00"
    ToPrint$ = ToPrint$ + "~$545.76"
    Print #UBRptT, ToPrint$

    Close
    If fpcboLogo.ListIndex = 1 Then
      If Exist(UBPath$ + "txTNlogo.bmp") Then
        ARptTempTaxBillDum.Image1.Picture = LoadPicture(UBPath$ + "txTNlogo.bmp")
        ARptTempTaxBillDum.Image1.Visible = True
      End If
    End If
    ARptTempTaxBillDum.Head1 = QPTrim(fpTxtHead1)
    ARptTempTaxBillDum.Head2 = QPTrim(fpTxtHead2)
    ARptTempTaxBillDum.LblOpt1 = QPTrim(fptxtOpt1)
    ARptTempTaxBillDum.LblOpt2 = QPTrim(fpTxtOpt2)
    ARptTempTaxBillDum.LblOpt3 = QPTrim(fpTxtOpt3)
    ARptTempTaxBillDum.LblOpt4 = QPTrim(fpTxtOpt4)
    ARptTempTaxBillDum.LblPgph1 = QPTrim(fptxtPgph(0))
    ARptTempTaxBillDum.LblPgph2 = QPTrim(fptxtPgph(1))
    ARptTempTaxBillDum.LblPgph3 = QPTrim(fptxtPgph(2))
    ARptTempTaxBillDum.LblPgph4 = QPTrim(fptxtPgph(3))
    ARptTempTaxBillDum.LblPgph5 = QPTrim(fptxtPgph(4))
    ARptTempTaxBillDum.LblPgph6 = QPTrim(fptxtPgph(5))
    ARptTempTaxBillDum.LblPgph7 = QPTrim(fptxtPgph(6))
    ARptTempTaxBillDum.LblPgph8 = QPTrim(fptxtPgph(7))
    ARptTempTaxBillDum.LblOpt5 = QPTrim(fpTxtOpt5)
    ARptTempTaxBillDum.LblHead4 = QPTrim(fptxtHead4)
    ARptTempTaxBillDum.LblHead5 = QPTrim(fptxtHead5)
    ARptTempTaxBillDum.LblHead6 = QPTrim(fptxtHead6)
    ARptTempTaxBillDum.LblOpt6 = QPTrim(fpTxtOpt6)
    ARptTempTaxBillDum.LblOpt7 = QPTrim(fpTxtOpt7)
    ARptTempTaxBillDum.GetName ReportFile$
    ARptTempTaxBillDum.startrpt

  End If
End Sub

Private Sub SaveTxDef()
  Dim UBTxBilDef As TxBillDefaultsType
  Dim Handle As Integer, UBTxBilDeflen As Integer
  UBTxBilDeflen = Len(UBTxBilDef)
  Handle = FreeFile
  Open UBPath$ + "UBTaxBil.DAT" For Random Shared As Handle Len = UBTxBilDeflen
    
    UBTxBilDef.TxtHead1 = QPTrim(fpTxtHead1)
    UBTxBilDef.TxtHead2 = QPTrim(fpTxtHead2)
    UBTxBilDef.txtOpt1 = QPTrim(fptxtOpt1)
    UBTxBilDef.TxtOpt2 = QPTrim(fpTxtOpt2)
    UBTxBilDef.TxtOpt3 = QPTrim(fpTxtOpt3)
    UBTxBilDef.TxtOpt4 = QPTrim(fpTxtOpt4)
    UBTxBilDef.txtPgph0 = QPTrim(fptxtPgph(0))
    UBTxBilDef.txtPgph1 = QPTrim(fptxtPgph(1))
    UBTxBilDef.txtPgph2 = QPTrim(fptxtPgph(2))
    UBTxBilDef.txtPgph3 = QPTrim(fptxtPgph(3))
    UBTxBilDef.txtPgph4 = QPTrim(fptxtPgph(4))
    UBTxBilDef.txtPgph5 = QPTrim(fptxtPgph(5))
    UBTxBilDef.txtPgph6 = QPTrim(fptxtPgph(6))
    UBTxBilDef.txtPgph7 = QPTrim(fptxtPgph(7))
    UBTxBilDef.TxtOpt5 = QPTrim(fpTxtOpt5)
    UBTxBilDef.txtHead4 = QPTrim(fptxtHead4)
    UBTxBilDef.txtHead5 = QPTrim(fptxtHead5)
    UBTxBilDef.txtHead6 = QPTrim(fptxtHead6)
    UBTxBilDef.TxtOpt6 = QPTrim(fpTxtOpt6)
    UBTxBilDef.TxtOpt7 = QPTrim(fpTxtOpt7)
    UBTxBilDef.dologo = fpcboLogo.ListIndex
  Put #Handle, 1, UBTxBilDef
  Close Handle
  MsgBox "Save Complete", vbOKOnly, "Saved"
End Sub
   
    
    
    
    
    
