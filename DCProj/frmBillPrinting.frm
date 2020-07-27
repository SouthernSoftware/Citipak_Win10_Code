VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBillPrinting 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Bill Printing"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmBillPrinting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboPrintOrder 
      Height          =   348
      Left            =   5412
      TabIndex        =   9
      Top             =   5544
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
      ColDesigner     =   "frmBillPrinting.frx":08CA
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
      TabIndex        =   15
      Top             =   7752
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
      TabIndex        =   14
      Top             =   7752
      Width           =   1332
   End
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   0
      Left            =   4716
      TabIndex        =   10
      Top             =   6120
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
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
      TabIndex        =   16
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
            TextSave        =   "10:14 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
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
   Begin EditLib.fpDateTime fpCurrDate 
      Height          =   348
      Left            =   5388
      TabIndex        =   6
      Top             =   4260
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin EditLib.fpDateTime fpPrevDate 
      Height          =   348
      Left            =   5388
      TabIndex        =   5
      Top             =   3852
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   1
      Left            =   4716
      TabIndex        =   11
      Top             =   6444
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
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
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   2
      Left            =   4716
      TabIndex        =   12
      Top             =   6768
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
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
   Begin EditLib.fpText fptxtMessage 
      Height          =   324
      Index           =   3
      Left            =   4716
      TabIndex        =   13
      Top             =   7092
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
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
   Begin EditLib.fpDateTime fpPastDueDate 
      Height          =   348
      Left            =   5400
      TabIndex        =   4
      Top             =   3432
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpDateTime fpBillingDate 
      Height          =   348
      Left            =   5400
      TabIndex        =   2
      Top             =   2568
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpDateTime fpDraftDate 
      Height          =   348
      Left            =   5400
      TabIndex        =   7
      Top             =   4680
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin EditLib.fpText fptxtBill1 
      Height          =   324
      Left            =   6384
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2138
      _ExtentY        =   572
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
      MaxLength       =   5
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
   Begin EditLib.fpText fptxtBill2 
      Height          =   324
      Left            =   6384
      TabIndex        =   1
      Top             =   1992
      Visible         =   0   'False
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2138
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
      MaxLength       =   5
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
   Begin EditLib.fpDateTime fpPastDate2 
      Height          =   348
      Left            =   5400
      TabIndex        =   8
      Top             =   5112
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      AllowNull       =   -1  'True
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
      Text            =   ""
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
   Begin EditLib.fpDateTime fpDueDate 
      Height          =   348
      Left            =   5400
      TabIndex        =   3
      Top             =   3000
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date:"
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
      Left            =   3576
      TabIndex        =   32
      Top             =   3024
      Width           =   1644
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2nd Penalty Date:"
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
      Left            =   2784
      TabIndex        =   31
      Top             =   5136
      Width           =   2436
   End
   Begin VB.Label LblApplyDep 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apply Deposit to Final Billing?"
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
      Left            =   2796
      TabIndex        =   30
      Top             =   7824
      Visible         =   0   'False
      Width           =   3324
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      FillColor       =   &H80000005&
      Height          =   972
      Left            =   2496
      Top             =   1464
      Visible         =   0   'False
      Width           =   7236
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Bill Number:"
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
      Left            =   4104
      TabIndex        =   29
      Top             =   1584
      Visible         =   0   'False
      Width           =   2148
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Bill Number:"
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
      Left            =   4056
      TabIndex        =   28
      Top             =   1992
      Visible         =   0   'False
      Width           =   2196
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5220
      Left            =   2496
      Top             =   2424
      Width           =   7236
   End
   Begin VB.Line Line1 
      X1              =   2496
      X2              =   9708
      Y1              =   5976
      Y2              =   5976
   End
   Begin VB.Label Labelm3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 3:"
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
      Left            =   2388
      TabIndex        =   27
      Top             =   6792
      Width           =   2220
   End
   Begin VB.Label Labelfrom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Date:"
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
      Left            =   3588
      TabIndex        =   26
      Top             =   2616
      Width           =   1644
   End
   Begin VB.Label Labelm2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 2:"
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
      Left            =   2436
      TabIndex        =   25
      Top             =   6456
      Width           =   2172
   End
   Begin VB.Label Labelm4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 4:"
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
      Left            =   2508
      TabIndex        =   24
      Top             =   7128
      Width           =   2100
   End
   Begin VB.Label Labelm1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message Line 1:"
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
      Left            =   2364
      TabIndex        =   23
      Top             =   6120
      Width           =   2244
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Draft Date:"
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
      Index           =   4
      Left            =   3372
      TabIndex        =   22
      Top             =   4704
      Width           =   1860
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Curr Read Date:"
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
      Left            =   3084
      TabIndex        =   21
      Top             =   4284
      Width           =   2148
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
      Height          =   372
      Index           =   3
      Left            =   2556
      TabIndex        =   20
      Top             =   5544
      Width           =   2676
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prev Read Date:"
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
      Left            =   3132
      TabIndex        =   19
      Top             =   3876
      Width           =   2100
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   456
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utility Bill Printing"
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
      Left            =   3900
      TabIndex        =   18
      Top             =   696
      Width           =   4428
   End
   Begin VB.Label Labelthru 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Past Due Date:"
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
      Left            =   3372
      TabIndex        =   17
      Top             =   3456
      Width           =   1860
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   336
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
Attribute VB_Name = "frmBillPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CycleFlag As Boolean, OKFlag As Boolean, BadDate As Boolean
Dim ErFlag As Boolean, BLType As Integer, LPIFlag As Boolean
Dim OkiMode As Integer, MaskBill As String ' 1 for not ibm, 2 for ibm
Dim IndianFlag As Boolean, PSAFlag As Boolean, MowFlag As Boolean
Dim UseDraftFlag As Boolean, lpcnt As Integer, PostBar As Boolean
Dim AcctBar As Boolean, MinBalance As Double, UsePrevFlag As Boolean
Dim UseCurrFlag As Boolean, PctAmt As Double, FixAmt As Double
Dim GreaterFlag As Boolean, UseBothFlag As Boolean, UsePctFlag As Boolean
Dim CustPenalty As Double, ReprintFlag As Boolean, PostedFlag As Boolean
Dim FirstBill As Long, LastBill As Long, BFlag As Boolean, BDate As Integer
Dim CustNum As Long, TNum As Long, O As String, M1 As String, ElkFlag As Boolean
Dim M2 As String, M3 As String, M4 As String, BalType As String
Dim NotUpdate As Boolean, Fflag As Boolean, UseDepositFlag As Boolean
Dim Deposit As String, ApplyDepFlag As String, DepFile As Integer
Dim Use2ndPen As Boolean, PenAmt2 As Double, Rteflag As Boolean
Dim TennRdg As Boolean, DueD As String, SpenTn As Boolean, PenTaxFlag As Boolean
  Dim UBSetupreclen As Integer, cnt As Long, ElecRev As Integer
  Dim BillInfoRecLen As Integer, ThisRevCnt As Integer, IndexName As String
  Dim UsingName As Boolean, UsingAcct As Boolean, UsingBook As Boolean
  Dim IdxTypeText As String, PastDateS As String, UBCustRecLen As Integer
  Dim UBOwnerRecLen As Integer, UBBillRecLen As Integer, Handle As Integer
  Dim UBDraftPayLen As Integer, NumOfRecs As Long, IdxRecLen As Integer
  Dim lcnt As Long, UBBill As Integer, UBCust As Integer, UBOwn As Integer
  Dim UBRpt As Integer, UBDraft As Integer, DFFileName As String
  Dim PrintedCnt As Long, NotDone As Boolean, CustAcctNo As Long
  Dim OName As String, Num2Print As Integer, BillDate As Integer
  Dim MtrCnt As Integer, UseEDateFlag As Boolean, PRDate As Integer
  Dim CRDate As Integer, DraftDate As Integer, Message As String
  Dim BillDateS As String, PrevDate As String, DateRead As String
  Dim DaysINRead As Integer, PastDueDate As String, TotalTax As Double
  Dim TaxCnt As Integer, DidADraftFlag As Boolean, BillCopies As Long
  Dim FFFlag As Integer, UBFile As Integer, ReportFile As String
  Dim DueDate As String, BillOrder As String, DueDate2 As Integer
  Dim Msg2 As String, Msg3 As String, Msg4 As String, DraftDateS As String
  Dim PrnCnt As Long, UsageAmt As Long, MaxMeterAmt As Long, Msg1 As String
  Dim Previous As Double, Totalamt As Double, FBillNO As Long, PastDate2 As Integer
  Dim FinalFlag As Boolean, CDeposit As Double, UBRptA As Integer
  Dim PCnt As Integer, MPCnt As Integer, PastDate As Integer, RCnt As Integer
  Dim tmprev As Double, ReqFldsOK As Boolean, ToPrint As String
  Dim FntSize As Integer, ToPrint2 As String, endit As Boolean
  Dim WFoundMtr As Boolean, GFoundMtr As Boolean, EFoundMtr As Boolean
  Dim mChk As Integer, WCurrRead As Long, WPrevRead As Long, WUsageAmt As Long
  Dim GCurrRead As Long, GPrevRead As Long, GUsageAmt As Long, GasTot As Double
  Dim BPrnCnt As Integer, ECurrRead As Long, EPrevRead As Long, EUsageAmt As Long
  Dim Zip As String, ZDigit As String, FoundAMtr As Boolean, Zero As String
  Dim AcctNum As Long, Acct As String, AcctLen As Integer, TBal As Double
  Dim WRevCnt As Integer, Low As Long, High As Long, ThisBillNum As Long
  Dim UBBillSetuplen As Integer, UBTran As Integer, BPrntType As Integer
  Dim UBPFile As Integer, OKgo4it As Boolean, MtrNFlag As Integer, doLogoflag As Integer
  Dim UBBillLetterlen As Integer, TmpAdd As String, BZip As String
  Dim scs As String, Fmt10 As String, Fmt10a As String, Fmt15 As String
  Dim Today As String, BillOutRecLen As Integer, Ext As String
  Dim SCSFileName As String, ChkName As String, IFoundMtr As Boolean
  Dim CubicMtr As Boolean, ICurrRead As Long, IPrevRead As Long, IUsageAmt As Long
  Dim MtrNumb As String, MtrType As String, MtrTyp As Integer, serv As Integer
Dim Bucksport As Boolean, DFoundMtr As Boolean
Dim DCurrRead As Long, DPrevRead As Long, DUsageAmt As Long
'  ReDim MsgText(0 To 5) As String
'  ReDim SFoundMtr(1 To 7) As Boolean
'  ReDim SCurrRead(1 To 7) As Long
'  ReDim SPrevRead(1 To 7) As Long
'  ReDim SUsageAmt(1 To 7) As Long
'  ReDim SMtrType(1 To 7) As String

'   1 Laser CitiPak 3Bill Legal
'   2 Laser 3Bill Letter
'   3 New Standard BarCode - TXT
'   4 New Standard V1 - TXT
'   5 New Standard Rm Stamp - TXT
'   6 Standard 24Line 2Bx - TXT
'   7 Standard 24Line 3Bx - TXT
'   8 Standard 21Line - TXT
'   9 Laser 3Bill Letter /Gas Revs
'  10 Laser 3Bill Legal Preprinted form
'  11 Laser 3Bill Legal(F) Preprinted form (For 1 line higher than 10)
'  12 Standard 21Line w/Loc - TXT (For Robbins)
'  13 Laser 3Bill Letter - Lg Font
'  14 For Spruce Pine 21 line standard 3BX
'  15 for Sunset #4 with mult meters
'  16 Letter Laser Bill - Blank Stock
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'On the Public Subs for different types of bill printing
'Set all flags appropriately.............
'
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'This REPRN is called from Printing All and Reprinting option from
'menu because both options use this form....
'
Public Sub REPRN(RPrn As Boolean, Optional F As Boolean)
  ReprintFlag = RPrn
  Fflag = F
  If RPrn = True Then
    NotUpdate = True
  End If
  BFlag = False
  PostedFlag = False
End Sub
'This is for Posted Bill Printing
Public Sub PBillPrn(CustNum1 As Long, TNum1 As Long, O1 As String, M11 As String, M21 As String, M31 As String, M41 As String, Due As String)
  Dim UBSetupreclen As Integer
  Dim UBBillSetuplen As Integer
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  'get Bill type from setup and store integer
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
  BLType = UBBillSetup(1).Bill
  ReprintFlag = False
  BFlag = False
  PostedFlag = True
  NotUpdate = True
  Use2ndPen = False
  DueD$ = Due
  GetPenaltyVals UBBillSetup()
  If UBBillSetup(1).PostBar = "Y" Then
    PostBar = True
  Else
    PostBar = False
  End If
  If UBBillSetup(1).AcctBar = "Y" Then
    AcctBar = True
  Else
    AcctBar = False
  End If
  If UBBillSetup(1).RtePrint = 1 Then
    Rteflag = True
  Else
    Rteflag = False
  End If
  If UBSetUp(1).BANKDFT = "Y" Then
    UseDraftFlag = True
  End If
  
  CustNum = CustNum1
  TNum = TNum1
  O = O1
  M1 = M11
  M2 = M21
  M3 = M31
  M4 = M41
  If BLType = 1 Or BLType = 2 Or BLType = 9 Or BLType = 10 Or BLType = 11 Or BLType = 13 Or BLType = 16 Or BLType = 17 Or BLType = 19 Or BLType = 20 Then
    PrintUtilBills True
  Else
    PrintUtilBills False
  End If
  DoEvents
  Unload Me
End Sub
'This is for B Status Bill Printing
Public Sub BBillPrn(x As String, Y As String, M11 As String, M21 As String, M31 As String, M41 As String, Due As String)
  Dim UBSetupreclen As Integer
  Dim UBBillSetuplen As Integer
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  'get Bill type from setup and store integer
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
  BLType = UBBillSetup(1).Bill
  ReprintFlag = False
  BFlag = True
  PostedFlag = False
  NotUpdate = True
  Use2ndPen = False
  DueD$ = Due
  'GetPenaltyVals UBBillSetup()
  If UBBillSetup(1).PostBar = "Y" Then
    PostBar = True
  Else
    PostBar = False
  End If
  If UBBillSetup(1).AcctBar = "Y" Then
    AcctBar = True
  Else
    AcctBar = False
  End If
  If UBBillSetup(1).RtePrint = 1 Then
    Rteflag = True
  Else
    Rteflag = False
  End If

  If UBSetUp(1).BANKDFT = "Y" Then
    UseDraftFlag = True
  End If
  BDate = Date2Num(x)
  BalType = Y
  M1 = M11
  M2 = M21
  M3 = M31
  M4 = M41
  If BLType = 1 Or BLType = 2 Or BLType = 9 Or BLType = 10 Or BLType = 11 Or BLType = 13 Or BLType = 16 Or BLType = 17 Or BLType = 19 Or BLType = 20 Then
    PrintUtilBills True
  Else
    PrintUtilBills False
  End If
  DoEvents
  Unload Me
End Sub

Private Sub cmdExit_Click()
  If Not Fflag Then
    Load frmUBPrintBillsMenu
    DoEvents
    frmUBPrintBillsMenu.Show
  Else
    Load frmUBFinalBillPrintMenu
    DoEvents
    frmUBFinalBillPrintMenu.Show
  End If
  Unload Me
  DoEvents
End Sub

Private Sub cmdPrint_Click()
  Dim Today As String, chkthedate As Integer, entdate As Integer
  Dim FntSize As Integer
    Today = Format(Now, "mm/dd/yyyy")
    chkthedate = Date2Num(Today)
    entdate = Date2Num(fpBillingDate)
    If entdate > (chkthedate + 30) Or entdate < (chkthedate - 30) Then
      
      UBLog "Out of Range Date entered Bill, give opt to cancel"
      ReDim MsgText(0 To 5) As String
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      frmMsgDialog.Label(4).FontSize = (FntSize + 2)
      frmMsgDialog.Label(2).FontSize = (FntSize + 2)
      MsgText(0) = "WARNING:"
      MsgText(1) = ""
      MsgText(2) = "Billing Date entered is NOT"
      MsgText(3) = "within monthly date range."
      MsgText(4) = ""
      MsgText(5) = "OK to continue, or Cancel."
      If GetOKorNot(MsgText()) Then
        UBLog "Continue Bill Printing with out of range date-" + fpBillingDate.Text
      Else
        UBLog "Cancel Bill Print so can check date."
        Exit Sub
      End If
    End If
  
  If ReprintFlag = True Then
    FirstBill = Val(fptxtBill1)
    LastBill = Val(fptxtBill2)
    If chkbillnos = False Then
      MsgBox "Invalid bill number selection.", vbOKOnly, "Invalid Selection"
      Exit Sub
    End If
    If FirstBill > LastBill Then
      MsgBox "Invalid bill number selection.", vbOKOnly, "Invalid Selection"
      Exit Sub
    End If
  End If
  DeActivateControls Me, True
  If BLType = 1 Or BLType = 2 Or BLType = 9 Or BLType = 10 Or BLType = 11 Or BLType = 13 Or BLType = 16 Or BLType = 17 Or BLType = 19 Or BLType = 20 Then
    PrintUtilBills True
  Else
    PrintUtilBills False
    ActivateControls Me, True
  End If
  If Fflag Then
    fptxtMessage(0).Enabled = False
  End If
  If ReprintFlag = True Then
    fpBillingDate.Enabled = False
    fpDueDate.Enabled = False
    fpPastDueDate.Enabled = False
    fpPrevDate.Enabled = False
    fpCurrDate.Enabled = False
    fpDraftDate.Enabled = False
    fpPastDate2.Enabled = False
    fpcboPrintOrder.Enabled = False
    fptxtMessage(0).Enabled = False
    fptxtMessage(1).Enabled = False
    fptxtMessage(2).Enabled = False
    fptxtMessage(3).Enabled = False
  End If
End Sub

Private Sub fpBillingDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpBillingDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpDueDate.SetFocus
  End If
End Sub
Private Sub fpDueDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpDueDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpPastDueDate.SetFocus
  End If
End Sub

Private Sub fpCurrDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCurrDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpDraftDate.SetFocus
  End If
End Sub

Private Sub fpDraftDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpPastDueDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpPastDueDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpPrevDate.SetFocus
  End If
End Sub

Private Sub fpPrevDate_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpPrevDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpCurrDate.SetFocus
  End If
End Sub
Private Sub fpDraftDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpPastDate2.SetFocus
  End If
End Sub
Private Sub fpPastDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboPrintOrder.SetFocus
  End If
End Sub
Private Sub fpPastDate2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpcboPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrintOrder.ListDown = True
  End If
  If fpcboPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fptxtMessage(0).Enabled = True Then
        fptxtMessage(0).SetFocus
      Else
        cmdPrint.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpPastDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtBill1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtBill1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      fptxtBill2.SetFocus
  End If
End Sub
Private Sub fptxtBill2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtBill2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
  End If
End Sub

Private Sub fptxtMessage_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If Index < 3 Then
     If fptxtMessage(Index + 1).Visible = True Then
      fptxtMessage(Index + 1).SetFocus
     Else
      cmdPrint.SetFocus
     End If
    Else
      cmdPrint.SetFocus
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
        UBLog "Closed via frmBillPrinting by " + PWUser$
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
      Call cmdPrint_Click
    Case Else:
  End Select
End Sub
Private Function chkbillnos()
  Dim BillInfoRecLen As Integer, UBFile As Integer
  ReDim BillInfoRec(1) As PrintBillInfoType
  BillInfoRecLen = Len(BillInfoRec(1))
  chkbillnos = True
  UBFile = FreeFile
  If Not Fflag Then
    Open UBPath$ + "UBPINFON.DAT" For Random As #UBFile Len = BillInfoRecLen
    Get #UBFile, 1, BillInfoRec(1)
    Close
  Else
    Open UBPath$ + "UBPINFOF.DAT" For Random As #UBFile Len = BillInfoRecLen
    Get #UBFile, 1, BillInfoRec(1)
    Close
  End If
  If Val(fptxtBill1) < BillInfoRec(1).FrstBill Then chkbillnos = False
  If Val(fptxtBill2) > BillInfoRec(1).LastBill Then chkbillnos = False

End Function

Private Sub Form_Load()
  Dim UBSetupreclen As Integer, cnt As Integer
  Dim UBBillSetuplen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  'get Bill type from setup and store integer
  ReDim UBBillSetup(1) As UBBillSetupType
  UBBillSetuplen = Len(UBBillSetup(1))
  LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
  BLType = UBBillSetup(1).Bill
  If UBBillSetup(1).RtePrint = 1 Then
    Rteflag = True
  Else
    Rteflag = False
  End If
  GetPenaltyVals UBBillSetup()
'  OkiMode = 2
'at same time get the OkiMode 1 is not ibm, 2 is ibm
  If UBBillSetup(1).PostBar = "Y" Then
    PostBar = True
  Else
    PostBar = False
  End If
  If UBBillSetup(1).AcctBar = "Y" Then
    AcctBar = True
  Else
    AcctBar = False
  End If
  If UBSetUp(1).BANKDFT = "Y" Then
    UseDraftFlag = True
  End If

  If Not ReprintFlag Then
    Call Pete
    If UBSetUp(1).UseSeq = "Y" Then
      fpcboPrintOrder.AddItem "Sequence Number Order"
    End If
  Else
    Call RePete
  End If
  If Fflag Then
    DepFile = FreeFile
    Open UBPath$ + "UBDEPFLG.DAT" For Random Shared As DepFile Len = 2
    Get DepFile, , UseDepositFlag
    Close DepFile
    If UseDepositFlag <> 0 Then
      Deposit$ = "Y"
      ApplyDepFlag$ = "Y"
      LblApplyDep.Caption = "DEPOSITS APPLIED"
    Else
      Deposit$ = "N"
      ApplyDepFlag$ = " "
      LblApplyDep.Caption = "DEPOSITS NOT APPLIED"
    End If
    LblApplyDep.Visible = True
  End If
    'Section to check for customer modifications
'also need to get the late fee or penalty amount for bills.

'  If InStr(TownName$, "INDIAN TRAIL") Then
'    IndianFlag = True
'  End If
'  If InStr(TownName$, "GILES") Then
'    PSAFlag = True
'  End If
'  If InStr(TownName$, "MOWAS") Then
'    MowFlag = True
'  End If
  If InStr(TOWNNAME$, "SPENCER") Then
    If InStr(UBSetUp(1).DEFSTATE, "TN") Then
      SpenTn = True
    Else
      SpenTn = False
    End If
  End If
  If InStr(TOWNNAME$, "NW CLAY") Then
    PenTaxFlag = True
  Else
    PenTaxFlag = False
  End If
  If InStr(TOWNNAME$, "ELKTON") Then
    ElkFlag = True
  Else
    ElkFlag = False
  End If
  If InStr(TOWNNAME$, "BUCKSPORT") Then
    Bucksport = True
  Else
    Bucksport = False
  End If

  If InStr(TOWNNAME$, "TENNESSEE RIDGE") Then
    TennRdg = True
  Else
    TennRdg = False
  End If
  Select Case BLType
    Case 1, 10, 11, 16, 17, 19, 20:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 50
      Next
'      For cnt = 0 To 3
'        fptxtMessage(cnt).Visible = False
'      Next
'      Labelm1.Visible = False
'      Labelm2.Visible = False
'      Labelm3.Visible = False
'      Labelm4.Visible = False
'    Case 2:
    Case 3:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 30
      Next
    Case 8, 12, 14:
      fptxtMessage(0).Maxlength = 25
      For cnt = 1 To 3
        fptxtMessage(cnt).Visible = False
      Next
      Labelm2.Visible = False
      Labelm3.Visible = False
      Labelm4.Visible = False
    Case 6, 7:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 25
      Next
      Labelm2.Caption = "   1 Continued:"
      Labelm3.Caption = "Message Line 2:"
      Labelm4.Caption = "   2 Continued:"
    Case 4, 5, 15:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 35
      Next
    Case 18:
      For cnt = 0 To 2
        fptxtMessage(cnt).Maxlength = 35
      Next
      Labelm4.Visible = False
      fptxtMessage(3).Visible = False
    Case 2, 9, 13:
      Labelm2.Visible = False
      Labelm3.Visible = False
      Labelm4.Visible = False
      For cnt = 1 To 3
        fptxtMessage(cnt).Visible = False
      Next
      fptxtMessage(0).Maxlength = 30
    Case 98:
      For cnt = 0 To 3
        fptxtMessage(cnt).Maxlength = 22
      Next
    Case Else
  End Select
    If Fflag And Not ReprintFlag Then
     fptxtMessage(0).Text = "Final Billing"
     fptxtMessage(0).Enabled = False
    End If
End Sub
Private Sub Pete()
  fpcboPrintOrder.AddItem "Customer Name Order"
  fpcboPrintOrder.AddItem "Account Number Order"
  If Not Fflag Then
    fpcboPrintOrder.AddItem "Location Number Order"
  Else
    fpcboPrintOrder.AddItem "No Location-Use Account"
  End If
  fpcboPrintOrder.AddItem "Postal Carrier Route Order"
  fpcboPrintOrder.AddItem "ZipCode Order"
  fpcboPrintOrder.AddItem "Zip/Location Order"
  fpcboPrintOrder.ListIndex = 0
  fpBillingDate.Text = Format(Now, "mm/dd/yyyy")
  fpDueDate.Text = Format(Now + 10, "mm/dd/yyyy")
  fpPastDueDate.Text = Format(Now + 10, "mm/dd/yyyy")
  'fpPastDate2.Text = Format(Now + 10, "mm/dd/yyyy")
  If Not BFlag Then
    NotUpdate = False
  End If
  If Fflag Then
    frmBillPrinting.Caption = "Final Bill Printing"
    Label1.Caption = "Final Bill Printing"
  End If
End Sub
Private Sub RePete()
  Dim BillInfoRecLen As Integer, UBFile As Integer
  ReDim BillInfoRec(1) As PrintBillInfoType
  BillInfoRecLen = Len(BillInfoRec(1))

  UBFile = FreeFile
  If Not Fflag Then
    Open UBPath$ + "UBPINFON.DAT" For Random As #UBFile Len = BillInfoRecLen
    Get #UBFile, 1, BillInfoRec(1)
    Close
  Else
    Open UBPath$ + "UBPINFOF.DAT" For Random As #UBFile Len = BillInfoRecLen
    Get #UBFile, 1, BillInfoRec(1)
    Close
  End If
  fptxtBill1.Visible = True
  fptxtBill2.Visible = True
  Label4.Visible = True
  Label6.Visible = True
  Shape4.Visible = True
  fptxtBill1 = QPTrim$(Str$(BillInfoRec(1).FrstBill))
  fptxtBill2 = QPTrim$(Str$(BillInfoRec(1).LastBill))
  fpBillingDate.Text = Num2Date$(BillInfoRec(1).BillDate)
  fpBillingDate.Enabled = False
  fpDueDate.Text = Num2Date$(BillInfoRec(1).DueDate)
  fpDueDate.Enabled = False
  fpPastDueDate.Text = Num2Date$(BillInfoRec(1).PastDate)
  fpPastDueDate.Enabled = False
  fpPastDate2.Text = Num2Date$(BillInfoRec(1).PastDate2)
  fpPastDate2.Enabled = False

  If BillInfoRec(1).PRDate > 0 Then
    fpPrevDate.Text = Num2Date$(BillInfoRec(1).PRDate)
  End If
  fpPrevDate.Enabled = False
  If BillInfoRec(1).CRDate > 0 Then
    fpCurrDate.Text = Num2Date$(BillInfoRec(1).CRDate)
  End If
  fpCurrDate.Enabled = False
  If BillInfoRec(1).DrftDate > 0 Then
    fpDraftDate.Text = Num2Date$(BillInfoRec(1).DrftDate)
  End If
  fpDraftDate.Enabled = False
  fpcboPrintOrder.Text = BillInfoRec(1).PrnOrder
  fpcboPrintOrder.Enabled = False
  fptxtMessage(0).Text = BillInfoRec(1).MsgLine1
  fptxtMessage(0).Enabled = False
  fptxtMessage(1).Text = BillInfoRec(1).MsgLine2
  fptxtMessage(1).Enabled = False
  fptxtMessage(2).Text = BillInfoRec(1).MsgLine3
  fptxtMessage(2).Enabled = False '.ControlType = ControlTypeReadOnly
  fptxtMessage(3).Text = BillInfoRec(1).MsgLine4
  fptxtMessage(3).Enabled = False
  If Fflag Then
    frmBillPrinting.Caption = "Final Bill Printing"
    Label1.Caption = "Final Bill Printing"
  End If

  Close UBFile

End Sub
Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Private Sub PrintUtilBills(Grpt As Boolean)
'  Dim UBSetupreclen As Integer, cnt As Integer, ElecRev As Integer
'  Dim BillInfoRecLen As Integer, ThisRevCnt As Integer, IndexName As String
'  Dim UsingName As Boolean, UsingAcct As Boolean, UsingBook As Boolean
'  Dim IdxTypeText As String, PastDateS As String, UBCustRecLen As Integer
'  Dim UBOwnerRecLen As Integer, UBBillRecLen As Integer, Handle As Integer
'  Dim UBDraftPayLen As Integer, NumOfRecs As Long, IdxRecLen As Integer
'  Dim lcnt As Long, UBBill As Integer, UBCust As Integer, UBOwn As Integer
'  Dim UBRpt As Integer, UBDraft As Integer, DFFileName As String
'  Dim PrintedCnt As Long, NotDone As Boolean, CustAcctNo As Long
'  Dim OName As String, Num2Print As Integer, BillDate As Integer
'  Dim MtrCnt As Integer, UseEDateFlag As Boolean, PRDate As Integer
'  Dim CRDate As Integer, DraftDate As Integer, Message As String
'  Dim BillDateS As String, PrevDate As String, DateRead As String
'  Dim DaysINRead As Integer, PastDueDate As String, TotalTax As Double
'  Dim TaxCnt As Integer, DidADraftFlag As Boolean, BillCopies As Long
'  Dim FFFlag As Integer, UBFile As Integer, ReportFile As String
'  Dim DueDate As String, BillOrder As String, DueDate2 As Integer
'  Dim Msg2 As String, Msg3 As String, Msg4 As String, DraftDateS As String
'  Dim PrnCnt As Long, UsageAmt As Long, MaxMeterAmt As Long, Msg1 As String
'  Dim Previous As Double, Totalamt As Double, FBillNO As Long, PastDate2 As Integer
'  Dim FinalFlag As Boolean, CDeposit As Double, UBRptA As Integer
'  Dim PCnt As Integer, MPCnt As Integer, PastDate As Integer, RCnt As Integer
'  Dim tmprev As Double, ReqFldsOK As Boolean, ToPrint As String
'  Dim FntSize As Integer, ToPrint2 As String, endit As Boolean
'  Dim WFoundMtr As Boolean, GFoundMtr As Boolean, EFoundMtr As Boolean
'  Dim mChk As Integer, WCurrRead As Long, WPrevRead As Long, WUsageAmt As Long
'  Dim GCurrRead As Long, GPrevRead As Long, GUsageAmt As Long, GasTot As Double
'  Dim BPrnCnt As Integer, ECurrRead As Long, EPrevRead As Long, EUsageAmt As Long
'  Dim Zip As String, ZDigit As String, FoundAMtr As Boolean, Zero As String
'  Dim AcctNum As Long, Acct As String, AcctLen As Integer, TBal As Double
'  Dim WRevCnt As Integer, Low As Long, High As Long, ThisBillNum As Long
'  Dim UBBillSetuplen As Integer, UBTran As Integer, BPrntType As Integer
'  Dim UBPFile As Integer, OKgo4it As Boolean, MtrNFlag As Integer, doLogoflag As Integer
'  Dim UBBillLetterlen As Integer, TmpAdd As String, BZip As String
'  Dim scs As String, Fmt10 As String, Fmt10a As String, Fmt15 As String
'  Dim Today As String, BillOutRecLen As Integer, Ext As String
'  Dim SCSFileName As String, ChkName As String, IFoundMtr As Boolean
'  Dim CubicMtr As Boolean, ICurrRead As Long, IPrevRead As Long, IUsageAmt As Long
'  Dim MtrNumb As String, MtrType As String, MtrTyp As Integer, serv As Integer
  ReDim MsgText(0 To 5) As String
  ReDim SFoundMtr(1 To 7) As Boolean
  ReDim SCurrRead(1 To 7) As Long
  ReDim SPrevRead(1 To 7) As Long
  ReDim SUsageAmt(1 To 7) As Long
  ReDim SMtrType(1 To 7) As String
 'Stop
  ToPrint$ = ""
  ToPrint2$ = "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ "
  ToPrint2$ = ToPrint2$ + "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ "
  ToPrint2$ = ToPrint2$ + "~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ "
If BLType = 16 Or BLType = 19 Then
    ReDim UBBillLetter(1) As UBBillLetterType
    UBBillLetterlen = Len(UBBillLetter(1))
    LoadUBBillLetterFile UBBillLetter(), UBBillLetterlen
    MtrNFlag = UBBillLetter(1).MtrNumFlag  '1=mtrserial, 2=mtrID
    doLogoflag = UBBillLetter(1).IncLogoFlag  '0=noprint, 1=print
End If
If BLType = 98 Then
  If Fflag Then
    scs$ = "F"
  ElseIf BFlag Then
    scs$ = "B"
  Else
    scs$ = "N"
  End If
    
  CrLf$ = Chr$(13) + Chr$(10)
  Fmt10$ = "##########"
  Fmt10a$ = "#######.##"
  Fmt15$ = "############.##"
  Today$ = Date$
  ReDim PrintRec(1) As BillOutRecType
  BillOutRecLen = Len(PrintRec(1))
  Ext$ = ".BTF"
  SCSFileName$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2) + scs$
  For cnt = 1 To 9
    ChkName$ = SCSFileName$ + QPTrim$(Str$(cnt)) + Ext$
    If Exist(ChkName$) = False Then
      SCSFileName$ = ChkName$
      Exit For
    End If
  Next
End If


'Don't let in here unless regular or reprint
If Not PostedFlag And Not BFlag Then
  If Not ReprintFlag Then
    If Not Fflag Then
      UBLog " IN: Bill printing."
    Else
      UBLog " IN: Final Bill Printing."
    End If
    LPIFlag = False
    If Not Fflag Then
      If Not ChkBillFile% Then
        frmMsgDialog.RetLabel = "-2"
        UBLog "ERROR: NO BILL FILE!"
        FntSize = frmMsgDialog.Label(3).FontSize
        frmMsgDialog.Label(1).FontSize = (FntSize + 2)
        frmMsgDialog.Label(3).FontSize = (FntSize + 2)
        MsgText(0) = "ERROR:"
        MsgText(1) = ""
        MsgText(2) = "NO BILL FILE!"
        MsgText(3) = ""
        MsgText(4) = ""
        MsgText(5) = ""
        GetOKorNot MsgText(), True
        ActivateControls Me, True
        GoTo ExitPrintBill
      End If
    End If
    GoSub CheckReqFields
  ElseIf ReprintFlag Then
    If Not Fflag Then
      UBLog " IN: Bill REprinting."
    Else
      UBLog " IN: Final Bill REprinting."
    End If
    LPIFlag = False
  End If
  FrmShowPctComp.Label1 = "Creating Bills"
  FrmShowPctComp.Show , Me

  ReDim BillInfoRec(1) As PrintBillInfoType
  BillInfoRecLen = Len(BillInfoRec(1))
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))
  LoadUBSetUpFile UBSetUpRec(), UBSetupreclen

  For ThisRevCnt = 1 To 15
    If InStr(UBSetUpRec(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") Then
      ElecRev = ThisRevCnt
      Exit For
    End If
  Next
  If Not ReprintFlag Then
    Select Case fpcboPrintOrder.ListIndex
    Case 0
      IndexName$ = NameIndexFile
      UsingName = True
      OKFlag = True
    Case 1
      IndexName$ = ""
      UsingAcct = True
      OKFlag = True
    Case 2
      If Not Fflag Then
        IndexName$ = BookIndexFile
        UsingBook = True
        OKFlag = True
      Else
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      End If
    Case 3, 4
      IdxTypeText$ = "Zip-Code"
      If fpcboPrintOrder.ListIndex = 3 Then
        IdxTypeText$ = "Postal Route"
        MakePostalIndex IdxTypeText$
        IndexName$ = TempIndexName
        OKFlag = True
      ElseIf (fpcboPrintOrder.ListIndex = 4 And PSAFlag) Or (fpcboPrintOrder.ListIndex = 4 And MowFlag) Then
        If MowFlag Then
          MakeMowZipCodeIndex IdxTypeText$
        Else
          MakeZipCodeIndex IdxTypeText$
        End If
        IndexName$ = TempIndexName
        OKFlag = True
      Else
        MakeMowZipCodeIndex IdxTypeText$
        'MakePostalIndex IdxTypeText$
        IndexName$ = TempIndexName
        OKFlag = True
      End If
    Case 5
      If Not Fflag Then
        IdxTypeText$ = "Zip-Location"
        MakeZipLocationIndex IdxTypeText$
      Else
        IdxTypeText$ = "Zip-Code"
        MakeMowZipCodeIndex IdxTypeText$
      End If
      IndexName$ = TempIndexName
      OKFlag = True
    Case 6
      IdxTypeText$ = "Sequence Number"
      MakeSequenceIndex IdxTypeText$, Me
      IndexName$ = TempIndexName
      OKFlag = True
    End Select
  End If
  PastDateS$ = fpPastDueDate
  'do bill printing
  '**************************************************************************
  If Fflag Then FinalFlag = True
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBOwnerRec(1) As UBOwnerRecType
  UBOwnerRecLen = Len(UBOwnerRec(1))

  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))

  ReDim UBDraftPayRec(1) As UBDraftPayRecType
  UBDraftPayLen = Len(UBDraftPayRec(1))
  If Not ReprintFlag Then
    BillInfoRec(1).BillDate = Date2Num(fpBillingDate)
    BillInfoRec(1).DueDate = Date2Num(fpDueDate)
    BillInfoRec(1).PastDate = Date2Num(fpPastDueDate)
    BillInfoRec(1).PRDate = Date2Num(fpPrevDate)
    BillInfoRec(1).CRDate = Date2Num(fpCurrDate)
    BillInfoRec(1).DrftDate = Date2Num(fpDraftDate)
    BillInfoRec(1).PastDate2 = Date2Num(fpPastDate2)
    BillInfoRec(1).PrnOrder = QPTrim$(fpcboPrintOrder.Text)
    BillInfoRec(1).MsgLine1 = QPTrim$(fptxtMessage(0).Text)
    BillInfoRec(1).MsgLine2 = QPTrim$(fptxtMessage(1).Text)
    BillInfoRec(1).MsgLine3 = QPTrim$(fptxtMessage(2).Text)
    BillInfoRec(1).MsgLine4 = QPTrim$(fptxtMessage(3).Text)
    'GOSUB UpdateInfoRec
      
    If Fflag Then
'      If fpcboApplyDep.ListIndex = 1 Then
'        ApplyDepFlag$ = "Y"
'      Else
'        ApplyDepFlag$ = " "
'      End If
      UBBillRec(1).ApplyDepFlag = ApplyDepFlag$
    End If
  '  If UsingAcct Then             'load the index
  '    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  '  Else
  '    NumOfRecs = FileSize(IndexName$) \ 4
  '    ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
  '    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  '  End If
    If UsingAcct Then
      NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
    Else          'load the index
      UBLog "Loading index file: " + IndexName$
      IdxRecLen = 4
      NumOfRecs = FileSize(IndexName$) \ 4
      ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
      Handle = FreeFile
      Open IndexName$ For Random Shared As Handle Len = IdxRecLen
      For lcnt& = 1 To NumOfRecs
        Get #Handle, lcnt&, IndexArray(lcnt&)
      Next
      Close Handle
      'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
    End If
  
  Else
    UBBill = FreeFile
    If Not Fflag Then
      Open UBPath$ + "UBBILLS.DAT" For Random Shared As UBBill Len = UBBillRecLen
    Else
      Open UBPath$ + UBFinBillsFile For Random Shared As UBBill Len = UBBillRecLen
    End If
    NumOfRecs = LOF(UBBill) \ UBBillRecLen
    ReDim RePrintIdx(1 To NumOfRecs) As RePrintIndexType
    For cnt = 1 To NumOfRecs
      Get UBBill, cnt, UBBillRec(1)
        If UBBillRec(1).ActiveFlag Then
          RePrintIdx(cnt).BillNum = UBBillRec(1).BillNumber
          RePrintIdx(cnt).BillRec = cnt
        End If
    Next
   Close UBBill
   Low = LBound(RePrintIdx)
   High = UBound(RePrintIdx)
   BillQSort RePrintIdx(), Low, High
  End If
  UBBill = FreeFile
  If Not Fflag Then
    Open UBPath$ + "UBBILLS.DAT" For Random Shared As UBBill Len = UBBillRecLen
  Else
    Open UBPath$ + UBFinBillsFile For Random Shared As UBBill Len = UBBillRecLen
  End If
  
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBOwn = FreeFile
  Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBOwn Len = UBOwnerRecLen

  UBRpt = FreeFile
  If Not BLType = 98 Then
    If Not ReprintFlag And Not Fflag Then
      ReportFile$ = UBPath$ + "UBBILLS.PRN"
    ElseIf Fflag And Not ReprintFlag Then
      ReportFile$ = UBPath$ + "UBFBILLS.PRN"
    ElseIf ReprintFlag And Not Fflag Then
      ReportFile$ = UBPath$ + "UBBILLSR.PRN"
    ElseIf ReprintFlag And Fflag Then
      ReportFile$ = UBPath$ + "UBFBILLR.PRN"
    End If
    Open ReportFile$ For Output As UBRpt
  Else
    Open SCSFileName$ For Random Shared As UBRpt Len = BillOutRecLen
  End If
  If UseDraftFlag Then
    DraftDateS$ = fpDraftDate
    DFFileName$ = "DF" + Left$(DraftDateS$, 2) + Mid$(DraftDateS$, 4, 2) + Right$(DraftDateS$, 2) + ".DAT"
    DFFileName$ = UBPath$ + DFFileName$

    UBDraft = FreeFile
    Open DFFileName$ For Random Shared As UBDraft Len = UBDraftPayLen
  End If
  If Not ReprintFlag Then
    If Fflag Then
      UBLog "Printing Final ub bills to disk."
    Else
      UBLog "Printing utility bills to disk."
    End If
  Else
    If Fflag Then
      UBLog "Preparing Reprint Final bill info."
    Else
      UBLog "Preparing Reprint bills information."
    End If
  End If
  'ShowProcessingScrn "Creating Utility Bills."

  '-----------------------------------------
  PrintedCnt = 0
  NotDone = True
  If Not ReprintFlag Then
    For cnt = 1 To NumOfRecs
      If UsingAcct Then
        CustAcctNo& = cnt
      Else
        CustAcctNo& = IndexArray(cnt).RecNum
      End If
      FrmShowPctComp.ShowPctComp cnt, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        If Not Fflag Then
          UBLog "ABORTED: Bill printing, AFTER START."
        Else
          UBLog "ABORTED: Final Bill printing, AFTER START."
        End If
        FrmShowPctComp.Out = False
        GoTo ExitPrintBill
      End If
  
      Get UBCust, CustAcctNo&, UBCustRec(1)
      Num2Print = UBCustRec(1).BILLCOPY
      'change 2/10/05 so if set billcopy to 0 no bill will print
      If Num2Print < 1 Then Num2Print = 0      '1
        If Fflag Then
          Get UBBill, CustAcctNo&, UBBillRec(1)
          If Not UBBillRec(1).ActiveFlag Then GoTo SkipEm
        End If
  '121598 Finished adding bill to owner
      If UBCustRec(1).BillTo = "O" Then   'if they want to send the bill
        'temp copy the owners info to the customers rec
        Get UBOwn, CustAcctNo&, UBOwnerRec(1)
        OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
        UBCustRec(1).CustName = OName$
        UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
        UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
        UBCustRec(1).CITY = UBOwnerRec(1).CITY
        UBCustRec(1).STATE = UBOwnerRec(1).STATE
        UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
      End If
  'look here when try to figure out way to give trans to 0 charg bill
      'Num2Print = UBCustRec(1).BILLCOPY
      'If Num2Print < 1 Then Num2Print = 1
      If Not Fflag Then
        If UBCustRec(1).Status <> "F" And UBCustRec(1).Status <> "I" And UBCustRec(1).Status <> " " Then
          Get UBBill, CustAcctNo&, UBBillRec(1)
          OKgo4it = True
        Else
          OKgo4it = False
        End If
      Else
        If UBBillRec(1).ActiveFlag Then
         'If Not UBCustRec(1).CurrBalance <> 0 Then
          'OKgo4it = False
          'GoTo SkipEm
         'End If
         OKgo4it = True
        Else
          OKgo4it = False
        End If
      End If
      If OKgo4it = True Then
        UBBillRec(1).TransDate = BillDate
        UBBillRec(1).TransDesc = "UTILITY BILL"
        If UBBillRec(1).ActiveFlag Then
          If UBBillRec(1).Transamt <> 0 Or UBCustRec(1).PrevBalance Then
            PrintedCnt = PrintedCnt + 1
            UBBillRec(1).BillNumber = PrintedCnt
            'see if anything in customer billing msg
            'if so then use instead on message on screen(depends on bill)
            Message$ = QPTrim$(UBCustRec(1).BILLCMNT)
            '****************
            '02-05-97 added code to get a valid meter read date,  maybe?  ;-]
            '         from one of the meters
            For MtrCnt = 1 To 7
              If UBBillRec(1).MtrTypes(MtrCnt) > 0 Then
                UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
                UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
                Exit For
              End If
            Next
            '02-05-97 this is to add a read date to a bill that has no metered
            '         services
            If UBBillRec(1).ReadDate <= 0 Then
              UBBillRec(1).ReadDate = BillDate - 30
            End If
            If UBBillRec(1).PrevDate <= 0 Then
              UBBillRec(1).PrevDate = UBBillRec(1).ReadDate - 30
            End If
  
            If UseEDateFlag Then
              UBBillRec(1).PrevDate = PRDate
              UBBillRec(1).ReadDate = CRDate
            End If
  
            UBBillRec(1).BillDate = BillDate
            If Use2ndPen Then
              If (Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance) > 0) And (Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance + UBBillRec(1).Transamt) > 0) Then
                UBBillRec(1).PastDueDate = PastDate2
              Else
                UBBillRec(1).PastDueDate = PastDate
              End If
            Else
              UBBillRec(1).PastDueDate = PastDate
            End If
            UBBillRec(1).DraftDate = DraftDate
            UBBillRec(1).BillMsg = Message$
  
            'these are for reprinting bills
            'UBBillRec(1)CustLocation = CustAcctNo&
            UBBillRec(1).CustAcctNo = CustAcctNo&
            BillDateS$ = Num2Date$(UBBillRec(1).BillDate)
  
            'if they entered a previous read date
            If BillInfoRec(1).PRDate > 0 Then
              PrevDate$ = Num2Date$(BillInfoRec(1).PRDate)
              PastDateS$ = Num2Date$(BillInfoRec(1).PRDate)
              DateRead$ = Num2Date$(BillInfoRec(1).CRDate)
              DaysINRead = BillInfoRec(1).CRDate - BillInfoRec(1).PRDate
            Else
              PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
              PastDateS$ = Num2Date$(UBBillRec(1).PrevDate)
              DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
              DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
            End If
  
  'IF DaysINRead > 59 THEN STOP
            DueDate$ = Num2Date$(BillInfoRec(1).DueDate)
            PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)
  
            TotalTax# = 0
            For TaxCnt = 1 To 14 'MaxRevsCnt
              TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
            Next
            If Fflag Then
              If ApplyDepFlag$ = "Y" Then
                CDeposit# = UBCustRec(1).DepositAmt
              Else
                CDeposit# = 0
              End If
              UBBillRec(1).ApplyDepFlag = ApplyDepFlag$
            End If
            
            Put UBBill, CustAcctNo&, UBBillRec(1)
            
            DidADraftFlag = False
            If UseDraftFlag And UBCustRec(1).USEDRAFT = "Y" And UBCustRec(1).PreNoteFlag Then
              UBDraftPayRec(1).CustAcctNum = CustAcctNo&
              UBDraftPayRec(1).DraftAmt = UBBillRec(1).Transamt
              Put UBDraft, , UBDraftPayRec(1)
              DidADraftFlag = True
            End If
'            If Fflag Then
'               'Custom Mod Here For Lilesville, NC
'              If Lilesville > 0 Then
'                If UBCustRec(1).Serv(1).RATECODE = "WIN " Or UBCustRec(1).Serv(1).RATECODE = "WOUT" Then
'                  TenPercentAmount# = (UBBillRec(1).Transamt - UBBillRec(1).RevAmt(1)) + (UBBillRec(1).RevAmt(1) * 1.1111)
'                Else
'                  TenPercentAmount# = UBBillRec(1).Transamt
'                End If
'              End If
'            'End Lilesville Custom Mod
'           End If
  'NOTE: Maccelsfield bill has best code to parse out meter readings
            For BillCopies = 1 To Num2Print
              GoSub PrintThemOne
  
            Next
          Else
            Put UBBill, CustAcctNo&, UBBillRec(1)
          End If
        Else
          'mod for cleveland***
          If UBBillRec(1).NONProfit = "Y" Then
            DOChurchTrans UBBillRec(), UBCustRec()
          End If
          '***
        End If
      End If
SkipEm:
    Next
  End If    'end if not reprint
End If    'end if not posted
If ReprintFlag Then
  UBPFile = FreeFile
  If Not Fflag Then
    Open UBPath$ + "UBPINFON.DAT" For Random As #UBPFile Len = BillInfoRecLen
    Get #UBPFile, 1, BillInfoRec(1)
    Close UBPFile
  Else
    Open UBPath$ + "UBPINFOF.DAT" For Random As #UBPFile Len = BillInfoRecLen
    Get #UBPFile, 1, BillInfoRec(1)
    Close UBPFile
  End If
  cnt = 0
  Do
    cnt = cnt + 1
    If cnt > NumOfRecs Then
      FrmShowPctComp.ShowPctComp cnt, cnt
      Exit Do
    End If
    ThisBillNum = RePrintIdx(cnt).BillNum
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      If Not Fflag Then
        UBLog "ABORTED: Bill Reprinting, AFTER START."
      Else
        UBLog "ABORTED: Final Bill Reprinting, AFTER START."
      End If
      FrmShowPctComp.Out = False
      GoTo ExitPrintBill
    End If

    If ThisBillNum >= FirstBill And ThisBillNum <= LastBill Then
      PrintedCnt = ThisBillNum
      CustAcctNo& = RePrintIdx(cnt).BillRec

      Get UBBill, CustAcctNo&, UBBillRec(1)
      Get UBCust, CustAcctNo&, UBCustRec(1)
'121598 Added bill to owner
      If UBCustRec(1).BillTo = "O" Then   'if bill owner, temp copy info
        Get UBOwn, CustAcctNo&, UBOwnerRec(1)
        OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
        UBCustRec(1).CustName = OName$
        UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
        UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
        UBCustRec(1).CITY = UBOwnerRec(1).CITY
        UBCustRec(1).STATE = UBOwnerRec(1).STATE
        UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
      End If
      Message$ = QPTrim$(UBCustRec(1).BILLCMNT)
      BillDateS$ = Num2Date$(UBBillRec(1).BillDate)
      DueDate$ = Num2Date$(BillInfoRec(1).DueDate)
      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      PastDateS$ = Num2Date$(UBBillRec(1).PrevDate)
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
      PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)
      'UBBillRec(1).BillMsg = Message$
      PRDate = Date2Num(fpPrevDate)
      CRDate = Date2Num(fpCurrDate)
    
      If (CRDate > 0) And (PRDate > 0) Then
        UseEDateFlag = True
      Else
        UseEDateFlag = False
      End If

      'if they entered a previous read date
      If BillInfoRec(1).PRDate > 0 Then
        PrevDate$ = Num2Date$(BillInfoRec(1).PRDate)
        PastDateS$ = Num2Date$(BillInfoRec(1).PRDate)
        DateRead$ = Num2Date$(BillInfoRec(1).CRDate)
        DaysINRead = BillInfoRec(1).CRDate - BillInfoRec(1).PRDate
      Else
        PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
        PastDateS$ = Num2Date$(UBBillRec(1).PrevDate)
        DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
        DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
      End If

      TotalTax# = 0
      For TaxCnt = 1 To 14  'MaxRevsCnt
        TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
      Next
      
      If ApplyDepFlag$ = "Y" Then
        CDeposit# = UBCustRec(1).DepositAmt
      Else
        CDeposit# = 0
      End If
      
      DidADraftFlag = False
      If UseDraftFlag And UBCustRec(1).USEDRAFT = "Y" And UBCustRec(1).PreNoteFlag Then
        DidADraftFlag = True
      End If
      Msg1$ = QPTrim$(BillInfoRec(1).MsgLine1)
      Msg2$ = QPTrim$(BillInfoRec(1).MsgLine2)
      Msg3$ = QPTrim$(BillInfoRec(1).MsgLine3)
      Msg4$ = QPTrim$(BillInfoRec(1).MsgLine4)

      Num2Print = UBCustRec(1).BILLCOPY
      'change on 2/10/05 so if set billcopy to 0 no bill will print
      If Num2Print < 1 Then Num2Print = 0   '1
'NOTE: Maccelsfield bill has best code to parse out meter readings
      For BillCopies = 1 To Num2Print
        GoSub PrintThemOne
      Next
    End If
  '  ShowPctComp cnt, NumOfRecs
  Loop
End If  'End Reprinted
If PostedFlag Then
  ReDim BillInfoRec(1) As PrintBillInfoType
  BillInfoRecLen = Len(BillInfoRec(1))

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))

  ReDim UBOwnerRec(1) As UBOwnerRecType
  UBOwnerRecLen = Len(UBOwnerRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, CustNum&, UBCustRec(1)
  Close UBCust

'  Select Case UBCustRec(1).BillTo
'  Case "O"
'    OFlag = True
'  End Select
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))
  LoadUBSetUpFile UBSetUpRec(), UBSetupreclen

  For ThisRevCnt = 1 To 15
    If InStr(UBSetUpRec(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") Then
      ElecRev = ThisRevCnt
      Exit For
    End If
  Next

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBBillRecLen
  Get #UBTran, TNum&, UBBillRec(1)
  If UBBillRec(1).TransType = TranUtilityBill Then
    Msg1$ = M1$
    Msg2$ = M2$
    Msg3$ = M3$
    Msg4$ = M4$
  
    PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
    PastDateS$ = Num2Date$(UBBillRec(1).PrevDate)
    DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
    DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
    DueDate$ = DueD$
    BillDateS$ = Num2Date$(UBBillRec(1).BillDate)
    CustAcctNo& = CustNum&
  'This section attempts to get the previous balance at the time the bill
  'was originally printed it is only temporary.
    UBCustRec(1).PrevBalance = 0
    UBCustRec(1).CurrBalance = 0
    If Round#(UBBillRec(1).RunBalance) <> Round#(UBBillRec(1).Transamt) Then
      UBCustRec(1).CurrBalance = Round#(UBBillRec(1).RunBalance - UBBillRec(1).Transamt)
    End If
''*********************************
    If O = "O" Then
      UBOwn = FreeFile
      Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBOwn Len = UBOwnerRecLen
      Get UBOwn, CustNum&, UBOwnerRec(1)
      OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
      UBCustRec(1).CustName = OName$
      UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
      UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
      UBCustRec(1).CITY = UBOwnerRec(1).CITY
      UBCustRec(1).STATE = UBOwnerRec(1).STATE
      UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
      Close UBOwn
    End If

    UBRpt = FreeFile
    ReportFile$ = UBPath$ + "UBBILLPP.PRN"
    Open ReportFile$ For Output As UBRpt
    PrintedCnt = 1
    endit = True
    GoSub PrintThemOne
    Close
'
'  If Not AbortFlag Then
'    PrintRptFile "Posted Bill Reprinting ", "UBBILLPP.PRN", 1, RetCode, 1
'  End If
'
'Reprint1Exit:
'  UBLog "OUT: Posted Bill Reprinting."
  End If
End If 'End Posted
If BFlag Then
    Select Case BalType$
    Case "B"
      BPrntType = 1
    Case "C"
      BPrntType = 2
    Case "A"
      BPrntType = 3
    End Select

'  NoUpDate = True

  LPIFlag = False
  FrmShowPctComp.Label1 = "Creating Bills"
  FrmShowPctComp.Show , Me

  ReDim BillInfoRec(1) As PrintBillInfoType
  BillInfoRecLen = Len(BillInfoRec(1))

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupreclen      'load setup file
  Msg1$ = M1
  Msg2$ = M2
  Msg3$ = M3
  Msg4$ = M4
  DueDate$ = DueD$
 ' FirstTime = True
 ' PastDay = Today + 10
  'do bill printing here
  '**************************************************************************
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))

  ReDim UBOwnerRec(1) As UBOwnerRecType
  UBOwnerRecLen = Len(UBOwnerRec(1))

  UBOwn = FreeFile
  Open UBPath$ + "UBOWNER.DAT" For Random Shared As UBOwn Len = UBOwnerRecLen

  NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBRpt = FreeFile
  If Not BLType = 98 Then
    ReportFile$ = UBPath$ + "UBBILLB.PRN"
    Open ReportFile$ For Output As UBRpt
  Else
    Open SCSFileName$ For Random Shared As UBRpt Len = BillOutRecLen
  End If
  UBLog "Printing utility bills to disk."
  'ShowProcessingScrn "Creating Utility Bills."

  '-----------------------------------------
  PrintedCnt = 0
  NotDone = True

  For cnt = 1 To NumOfRecs
    CustAcctNo& = cnt
    TBal# = 0
      FrmShowPctComp.ShowPctComp cnt, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        UBLog "ABORTED: B Bill printing, AFTER START."
        FrmShowPctComp.Out = False
        GoTo ExitPrintBill
      End If
    Get UBCust, CustAcctNo&, UBCustRec(1)

    If UBCustRec(1).Status = "B" Then
      TBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      If TBal# <> 0 Then
        ReDim UBBillRec(1) As UBTransRecType
        Select Case BPrntType
        Case 1  'credit bills
          If TBal# > 0 Then
            GoSub PRintOne
          End If
        Case 2  'balance bills
          If TBal# < 0 Then
            GoSub PRintOne
          End If
        Case 3  'all
          GoSub PRintOne
        End Select
      End If
    End If
'    ShowPctComp cnt, NumOfRecs
  Next 'Next BFlag
  GoTo bskipem

PRintOne:

  If UBCustRec(1).BillTo = "O" Then
    Get UBOwn, CustAcctNo&, UBOwnerRec(1)
    OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
    UBCustRec(1).CustName = OName$
    UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
    UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
    UBCustRec(1).CITY = UBOwnerRec(1).CITY
    UBCustRec(1).STATE = UBOwnerRec(1).STATE
    UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
  End If

  For RCnt = 1 To 15
    If UBCustRec(1).CurrRevAmts(RCnt) <> 0 Then
      UBBillRec(1).RevAmt(RCnt) = UBCustRec(1).CurrRevAmts(RCnt)
    End If
  Next

  Num2Print = UBCustRec(1).BILLCOPY
  'change on 2/10/05 so if set billcopy to 0 no bill will print
  If Num2Print < 1 Then Num2Print = 0     '1

  PrintedCnt = PrintedCnt + 1
  UBBillRec(1).BillNumber = PrintedCnt

  'Look for a valid meter read date,  maybe?
  'from one of the meters
  For MtrCnt = 1 To 7
    If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrType)) > 0 Then
      UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
      UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
      'UBBillRec(1).CurRead(1) = UBCustRec(1).LocMeters(MtrCnt).CurRead
      'UBBillRec(1).PrevRead(1) = UBCustRec(1).LocMeters(MtrCnt).PrevRead
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      UBBillRec(1).MtrTypes(1) = 1
      Exit For
    End If
  Next

  UBBillRec(1).CustAcctNo = CustAcctNo&
  UBBillRec(1).BillDate = BDate
  UBBillRec(1).PastDueDate = UBBillRec(1).BillDate

  BillDateS$ = Num2Date$(UBBillRec(1).BillDate)
  PastDueDate$ = BillDateS$

'NOTE: Maccelsfield bill has best code to parse out meter readings
  For BillCopies = 1 To Num2Print
    GoSub PrintThemOne
  Next

  Return

bskipem:

  Close
  'end bill printing
  GoTo BExitPrintBill:
  '******************
BExitPrintBill:
  UBLog "OUT: B-Status Bill Printing."
End If 'End BFlag Bills
  If Not LPIFlag And BLType <> 1 And BLType <> 2 And BLType <> 9 And BLType <> 10 And BLType <> 11 And BLType <> 13 And BLType <> 16 And BLType <> 17 And BLType <> 19 And BLType <> 20 Then
      If InStr(TOWNNAME$, "PEACHLAND") Then
        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
      Else
        Print #UBRpt, Chr$(27); Chr$(50);           'set printer in 6 lines per in
      End If
  End If

  If FFFlag And BLType <> 1 And BLType <> 2 And BLType <> 9 And BLType <> 10 And BLType <> 11 And BLType <> 13 And BLType <> 16 And BLType <> 17 And BLType <> 19 And BLType <> 20 Then
    Print #UBRpt, Chr$(12);
  End If
   If BLType = 1 Then
    endit = True
    'GoSub Dblcheck1
  End If
  If BLType = 2 Or BLType = 13 Then
    endit = True
    GoSub Dblcheck
  End If
  If BLType = 9 Then
    endit = True
    GoSub Dblcheck2
  End If

  Close
  If NotUpdate = False Then
    UBLog "Finished printing to disk."
  
    BillInfoRec(1).FrstBill = 1
    BillInfoRec(1).LastBill = PrintedCnt
    BillInfoRec(1).BillDate = Date2Num(fpBillingDate)
    BillInfoRec(1).DueDate = Date2Num(fpDueDate)
    BillInfoRec(1).PastDate = Date2Num(fpPastDueDate)
    BillInfoRec(1).PRDate = Date2Num(fpPrevDate)
    BillInfoRec(1).CRDate = Date2Num(fpCurrDate)
    BillInfoRec(1).DrftDate = Date2Num(fpDraftDate)
    BillInfoRec(1).PrnOrder = QPTrim$(fpcboPrintOrder.Text)
  
    BillInfoRec(1).MsgLine1 = QPTrim$(fptxtMessage(0).Text)
    BillInfoRec(1).MsgLine2 = QPTrim$(fptxtMessage(1).Text)
    BillInfoRec(1).MsgLine3 = QPTrim$(fptxtMessage(2).Text)
    BillInfoRec(1).MsgLine4 = QPTrim$(fptxtMessage(3).Text)
  
    UBFile = FreeFile
    If Not Fflag Then
      Open UBPath$ + "UBPINFON.DAT" For Random As #UBFile Len = BillInfoRecLen
      Put #UBFile, 1, BillInfoRec(1)
      Close
      UBLog "Updated: Bill Information File."
    ElseIf Fflag Then
      Open UBPath$ + "UBPINFOF.DAT" For Random As #UBFile Len = BillInfoRecLen
      Put #UBFile, 1, BillInfoRec(1)
      Close
      UBLog "Updated: Final Bill Information File."
    End If
    Erase UBCustRec, UBBillRec, BillInfoRec, IndexArray
  ElseIf ReprintFlag Then
    Erase UBCustRec, UBBillRec, RePrintIdx, IndexArray
  ElseIf PostedFlag Or BFlag Then
    Erase UBCustRec, UBBillRec, BillInfoRec, IndexArray
  End If

    'PrintRptFile "Utility Bill Printing ", "UBBILLS.PRN", 1, RetCode, 1
  DoBLMask
  Select Case BLType
  Case 1:
    ReDim UBBillSetup(1) As UBBillSetupType
    UBBillSetuplen = Len(UBBillSetup(1))
    LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
    
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If PostBar = False Then ARptBillLaserLegal.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLegal.Barcode2.Visible = False
    ARptBillLaserLegal.Head1 = QPTrim(UBBillSetup(1).BL1Head1)
    ARptBillLaserLegal.Head2 = QPTrim(UBBillSetup(1).BL1Head2)
    ARptBillLaserLegal.Head3 = QPTrim(UBBillSetup(1).BL1Head3)
    ARptBillLaserLegal.LblOpt1 = QPTrim(UBBillSetup(1).BL1Opt1)
    ARptBillLaserLegal.LblOpt2 = QPTrim(UBBillSetup(1).BL1Opt2)
    ARptBillLaserLegal.LblOpt3 = QPTrim(UBBillSetup(1).BL1Opt3)
    ARptBillLaserLegal.LblOpt4 = QPTrim(UBBillSetup(1).BL1Opt4)
    ARptBillLaserLegal.LblOpt5 = QPTrim(UBBillSetup(1).BL1Opt5)
    ARptBillLaserLegal.LblOpt6 = QPTrim(UBBillSetup(1).BL1Opt6)
    ARptBillLaserLegal.LblOpt7 = QPTrim(UBBillSetup(1).BL1Opt7)
    ARptBillLaserLegal.LblOpt8 = QPTrim(UBBillSetup(1).BL1Opt8)
    ARptBillLaserLegal.LblOpt9 = QPTrim(UBBillSetup(1).BL1Opt9)
    ARptBillLaserLegal.LblOpt10 = QPTrim(UBBillSetup(1).BL1Opt10)
    If UBBillSetup(1).Permit = "Y" Then
      ARptBillLaserLegal.LblPermit1 = QPTrim(UBBillSetup(1).BL1Permit1)
      ARptBillLaserLegal.LblPermit2 = QPTrim(UBBillSetup(1).BL1Permit2)
      ARptBillLaserLegal.LblPermit3 = QPTrim(UBBillSetup(1).BL1Permit3)
      ARptBillLaserLegal.LblPermit4 = QPTrim(UBBillSetup(1).BL1Permit4)
      ARptBillLaserLegal.LblPermit5 = QPTrim(UBBillSetup(1).BL1Permit5)
    Else
      ARptBillLaserLegal.LblPermit1.Visible = False
      ARptBillLaserLegal.LblPermit2.Visible = False
      ARptBillLaserLegal.LblPermit3.Visible = False
      ARptBillLaserLegal.LblPermit4.Visible = False
      ARptBillLaserLegal.LblPermit5.Visible = False
      ARptBillLaserLegal.Shape1.Visible = False
    End If
    ARptBillLaserLegal.GetName ReportFile$
    ARptBillLaserLegal.startrpt
  Case 2, 9:
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    ARptBillLaser.GetName ReportFile$
    ARptBillLaser.startrpt
  Case 13:
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    ARptBillLaserF.GetName ReportFile$
    ARptBillLaserF.startrpt
  Case 10:
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If PostBar = False Then ARptBillLaserLegal2.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLegal2.Barcode2.Visible = False
    ARptBillLaserLegal2.GetName ReportFile$
    ARptBillLaserLegal2.startrpt
  Case 11:
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If PostBar = False Then ARptBillLaserLegal2F.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLegal2F.Barcode2.Visible = False
    ARptBillLaserLegal2F.GetName ReportFile$
    ARptBillLaserLegal2F.startrpt
  Case 3, 4, 5, 6, 7, 8, 12, 14, 15, 18:
      ViewPrint ReportFile$, "Utility Bill Printing", False, , True, MaskBill$, True
'    Else if no mask
'      ViewPrint ReportFile$, "Utility Bill Printing"
  Case 16:
    ' do stuff for laser letter bill
    ReDim UBBillSetup(1) As UBBillSetupType
    UBBillSetuplen = Len(UBBillSetup(1))
    LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
    ReDim UBBillLetter(1) As UBBillLetterType
    UBBillLetterlen = Len(UBBillLetter(1))
    LoadUBBillLetterFile UBBillLetter(), UBBillLetterlen

    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If UBBillLetter(1).IncLogoFlag = 1 Then
      If Exist(UBPath$ + "UBTNlogo.bmp") Then
        ARptBillLaserLetterForm.Image1.Picture = LoadPicture(UBPath$ + "\UBTNlogo.bmp")
        ARptBillLaserLetterForm.Image1.Visible = True
      End If
    End If
    If PostBar = False Then ARptBillLaserLetterForm.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLetterForm.Barcode2.Visible = False
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
  Case 17:
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If PostBar = False Then ARptBillLaserLegal3F.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLegal3F.Barcode2.Visible = False
    ARptBillLaserLegal3F.GetName ReportFile$
    ARptBillLaserLegal3F.startrpt
  Case 19:
    ' do stuff for laser letter bill preprinted
    ReDim UBBillSetup(1) As UBBillSetupType
    UBBillSetuplen = Len(UBBillSetup(1))
    LoadUBBillSetUpFile UBBillSetup(), UBBillSetuplen
    ReDim UBBillLetter(1) As UBBillLetterType
    UBBillLetterlen = Len(UBBillLetter(1))
    LoadUBBillLetterFile UBBillLetter(), UBBillLetterlen

    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If PostBar = False Then ARptBillLaserLetterPrePrinted.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLetterPrePrinted.Barcode2.Visible = False
    If Fflag Then ARptBillLaserLetterPrePrinted.lblfinal.Visible = True
    ARptBillLaserLetterPrePrinted.LblPgph1 = QPTrim(UBBillLetter(1).MsgPgph1)
    ARptBillLaserLetterPrePrinted.LblPgph2 = QPTrim(UBBillLetter(1).MsgPgph2)
    ARptBillLaserLetterPrePrinted.LblPgph3 = QPTrim(UBBillLetter(1).MsgPgph3)
    ARptBillLaserLetterPrePrinted.LblPgph4 = QPTrim(UBBillLetter(1).MsgPgph4)
    ARptBillLaserLetterPrePrinted.LblPgph5 = QPTrim(UBBillLetter(1).MsgPgph5)
    ARptBillLaserLetterPrePrinted.GetName ReportFile$
    ARptBillLaserLetterPrePrinted.startrpt
  Case 20:
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmBillPrinting
    If PostBar = False Then ARptBillLaserLegal2Big.Barcode1.Visible = False
    If AcctBar = False Then ARptBillLaserLegal2Big.Barcode2.Visible = False
    ARptBillLaserLegal2Big.GetName ReportFile$
    ARptBillLaserLegal2Big.startrpt
  Case 98:
    MsgBox "File " + SCSFileName$ + " has been created.", vbOKOnly, "Process Complete"
  Case Else
  End Select
  ' **************************************************************************
  'end bill printing
  GoTo ExitPrintBill:
  '******************

  '******************
CheckReqFields:

  BillDate = Date2Num(fpBillingDate)
  DueDate2 = Date2Num(fpDueDate)
  PastDate = Date2Num(fpPastDueDate)
  DraftDate = Date2Num(fpDraftDate)
  PastDate2 = Date2Num(fpPastDate2)
  DueDate$ = fpDueDate
  BillOrder$ = QPTrim$(Left$(fpcboPrintOrder.Text, 1))

  Msg1$ = QPTrim$(fptxtMessage(0).Text)
  Msg2$ = QPTrim$(fptxtMessage(1).Text)
  Msg3$ = QPTrim$(fptxtMessage(2).Text)
  Msg4$ = QPTrim$(fptxtMessage(3).Text)

  PRDate = Date2Num(fpPrevDate)
  CRDate = Date2Num(fpCurrDate)

  If (CRDate > 0) And (PRDate > 0) Then
    UseEDateFlag = True
  Else
    UseEDateFlag = False
  End If
  If PRDate > 0 And CRDate < 0 Then
    MsgText(2) = "Invalid Reading Date."
  ElseIf CRDate > 0 And PRDate < 0 Then
    MsgText(2) = "Invalid Reading Date."
  ElseIf BillDate = -32768 Then
    MsgText(2) = "Invalid Billing Date."
  ElseIf PastDate < BillDate Then
    MsgText(2) = "Invalid Past Due Date."
  ElseIf (UseDraftFlag And DraftDate = -32768) Or (UseDraftFlag And DraftDate < BillDate) Then
    MsgText(2) = "Invalid Draft Date."
  ElseIf (Use2ndPen And PastDate2 = -32768) Or (Use2ndPen And PastDate2 < BillDate) Then
    MsgText(2) = "Invalid 2nd Penalty Date."
  ElseIf Len(BillOrder$) = 0 Then
    MsgText(2) = "Invalid Printing Order."
  Else
    ReqFldsOK = True
  End If
  If Not ReqFldsOK Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO BILL FILE!"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
   'MsgText(2) = "NO BILL FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    ActivateControls Me, True
    GoTo ExitPrintBill
  End If

  If UseDraftFlag Then
    DraftDateS$ = fpDraftDate
    DFFileName$ = "DF" + Left$(DraftDateS$, 2) + Mid$(DraftDateS$, 4, 2) + Right$(DraftDateS$, 2) + ".DAT"
    DFFileName$ = UBPath$ + DFFileName$
  End If
  UBLog "BillPrn-" + fpBillingDate + "," + fpPastDueDate + "," + fpPrevDate + "," + fpCurrDate + "," + fpDraftDate + "," + fpPastDate2

  Return

PrintThemOne:
'BillTypes
   Select Case BLType
    Case 1, 10, 11, 17, 20: 'Laser 3bill legal 1)full format and 10)preprinted stock and 11)F,17)3F
      GoSub PrnLaserLegal
    Case 2, 13:
      GoSub PrnLaser '3-bill letter SIZE format
    Case 3:
      If ElkFlag Then
        PrnBill3ForElkton UBCustRec(), UBBillRec(), UBSetUpRec()
      Else
        GoSub PrnNewStandBarCode
      End If
    Case 4, 18:
      If Bucksport = True Then
        PrnBill4ForBucksp UBCustRec(), UBBillRec(), UBSetUpRec()
      Else
        GoSub PrintNewStandV1
      End If
    Case 5:
      GoSub PrintNewStandRmStamp
    Case 6:
      GoSub PrnStand24L2Bx
    Case 7:
      GoSub PrnStand24L3Bx
    Case 8, 12, 14:
      GoSub PrnStand21Line
    Case 9:
      GoSub PrnLaserCashion 'gas revs
    Case 15:
      GoSub PrintNewStandV1SB
    Case 16:
       PrnLaserLetterForm UBCustRec(), UBBillRec(), UBSetUpRec()
    Case 19:
       PrnLaserLetterPrePrinted UBCustRec(), UBBillRec(), UBSetUpRec()
    Case 98:
       CreateSCSFileTransfer UBCustRec(), UBBillRec(), UBSetUpRec()
    Case Else
   End Select
   
   PrnCnt = PrnCnt + 1

Return

'@@@@@@@@@@@   BILLS    @@@@@@@@@@@@@@@@@@@@@@@@@
'''''
PrnLaserLegal:  'Use billing date and past due date
'Utility Bill Laser(Uses Blank Stock)BAR CODE PRINTABLE
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  ZDigit$ = GetZipEDigit$(Zip$)
  Zip$ = Zip$ + ZDigit$
  lpcnt = 8
  FoundAMtr = False
  
  
    For mChk = 1 To 7
      SFoundMtr(mChk) = False
    Next
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
          SCurrRead&(mChk) = UBBillRec(1).CurRead(mChk)
          SPrevRead&(mChk) = UBBillRec(1).PrevRead(mChk)
          SUsageAmt&(mChk) = SCurrRead&(mChk) - SPrevRead&(mChk)
          'SUsageAmt&(mChk) = Round(SUsageAmt&(mChk) * UBCustRec(1).LocMeters(mChk).MTRMulti)
          If SUsageAmt&(mChk) < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(SPrevRead&(mChk))) - 1)
            SUsageAmt&(mChk) = (MaxMeterAmt& - SPrevRead&(mChk)) + SCurrRead&(mChk)
          End If
          SUsageAmt&(mChk) = Round(SUsageAmt&(mChk) * UBCustRec(1).LocMeters(mChk).MTRMulti)
          Select Case UBBillRec(1).MtrTypes(mChk)
          Case 1
            SMtrType(mChk) = "W"
          Case 2
            SMtrType(mChk) = "S"
          Case 3
            SMtrType(mChk) = "C"
          Case 4
            SMtrType(mChk) = "E"
          Case 5
            SMtrType(mChk) = "D"
          Case 6
            SMtrType(mChk) = "G"
          Case 7
            SMtrType(mChk) = "T"
          Case 9
            SMtrType(mChk) = "I"
          End Select
          SFoundMtr(mChk) = True
      End If
    Next
    For mChk = 1 To 7
      If SFoundMtr(mChk) = False Then
        FoundAMtr = False
      Else
        FoundAMtr = True
        Exit For
      End If
    Next
    
'  For mChk = 1 To 7
'    If UBBillRec(1).MtrTypes(mChk) > 0 Then
'      FoundAMtr = True
'      Exit For
'    End If
'  Next
'  Select Case UBCustRec(1).LocMeters(1).MTRMulti
'  Case 10
'    Zero$ = "0"
'  Case 100
'    Zero$ = "00"
'  Case 1000
'    Zero$ = "000"
'  Case Else
'    Zero$ = ""
'  End Select
  If UseEDateFlag = False Then
    If FoundAMtr = False Then
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    End If
  End If
  AcctNum = UBBillRec(1).CustAcctNo
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
  If Previous# <> 0 Then
    lpcnt = lpcnt - 1
  End If
  If TotalTax > 0 Then
    lpcnt = lpcnt - 1
  End If

  If FinalFlag And CDeposit# Then
    Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
  End If
'Need to get Penalty calculated based on defaults from setup screen
  If Totalamt# > 0 And Not FinalFlag Then
    If UBCustRec(1).LATEFEE = "Y" Then
        If TotalTax > 0 Then
          If PenTaxFlag Then
            CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
          Else
            CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
          End If
        Else
          CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
        End If
    Else
      CustPenalty# = 0
    End If
  Else
    CustPenalty# = 0
  End If
'-=-=-=-=-=-=-=-=-=-=-=-
  AcctNum = UBBillRec(1).CustAcctNo
  Acct$ = QPTrim$(Str$(AcctNum))
  Select Case AcctNum
  Case Is < 10
    Acct$ = "00" + Acct$
  Case Is < 100
    Acct$ = "0" + Acct$
  End Select
  ToPrint$ = Using("########", (FBillNO& + PrintedCnt))
  If Not BFlag Then
    ToPrint$ = ToPrint$ + "~" + PrevDate$ + "~" + DateRead$
     'Only Print Days if Greater than 0
     If DaysINRead > 0 Then
       ToPrint$ = ToPrint$ + "~" + Using("####", DaysINRead)
     Else
       ToPrint$ = ToPrint$ + "~ "
     End If
  Else
    ToPrint$ = ToPrint$ + "~ ~ "
    ToPrint$ = ToPrint$ + "~ "
  End If
  PCnt = 0
  

    For WRevCnt = 1 To lpcnt - 1
      PCnt = PCnt + 1
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        ToPrint$ = ToPrint$ + "~" + Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3)
'        If UBBillRec(1).CurRead(WRevCnt) > 0 Then
'          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'          If UsageAmt& < 0 Then
'            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
'            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'          End If
'          ToPrint$ = ToPrint$ + "~" + Using("#########", UBBillRec(1).PrevRead(WRevCnt)) + Zero$
'          ToPrint$ = ToPrint$ + "~" + Using("#########", UBBillRec(1).CurRead(WRevCnt)) + Zero$
'          ToPrint$ = ToPrint$ + "~" + Using("#######", UsageAmt&) + Zero$
'        Else
        For mChk = 1 To 7
          If Not BFlag Then

          If UBBillRec(1).MtrTypes(mChk) > 0 Then
          
            If UBCustRec(1).serv(WRevCnt).RMtrType = SMtrType(mChk) Then
              ToPrint$ = ToPrint$ + "~" + Using("##########", SPrevRead&(mChk))
              ToPrint$ = ToPrint$ + "~" + Using("##########", SCurrRead&(mChk))
              ToPrint$ = ToPrint$ + "~" + Using("#######", SUsageAmt&(mChk))
              Exit For
'            Else
'              ToPrint$ = ToPrint$ + "~ ~ ~ "
'              Exit For
            End If
          Else
            ToPrint$ = ToPrint$ + "~ ~ ~ "
            Exit For
          End If
          Else
            ToPrint$ = ToPrint$ + "~ ~ ~ "
            Exit For
          End If

        Next
        ToPrint$ = ToPrint$ + "~" + Using("#####.##", UBBillRec(1).RevAmt(WRevCnt))
      Else
        ToPrint$ = ToPrint$ + "~ ~ ~ ~ ~ "
      End If
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        ToPrint$ = ToPrint$ + "~ ~ ~ ~" + "   Other:"
        ToPrint$ = ToPrint$ + "~" + Using("#####.##", tmprev#)
      Else
        ToPrint$ = ToPrint$ + "~ ~ ~ ~ ~ "
      End If

    If TotalTax# > 0 Then
      ToPrint$ = ToPrint$ + "~ ~ ~ ~" + "     TAX:" + "~" + Using("$###,###.##", TotalTax#)
'    Else
'      Print #UBRpt, ""
    End If
    If Previous# <> 0 Then
      ToPrint$ = ToPrint$ + "~ ~ ~ ~" + "Previous:" + "~" + Using("$###,###.##", Previous#)
'    Else
'      Print #UBRpt,
    End If
    '" Current:"
    ToPrint$ = ToPrint$ + "~" + Using("$###,###.##", UBBillRec(1).Transamt)

    If FinalFlag And CDeposit# Then
      ToPrint$ = ToPrint$ + "~" + " Deposit:" + "~" + Using("$###,###.##", -UBCustRec(1).DepositAmt)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    ToPrint$ = ToPrint$ + "~" + Num2Date$(UBBillRec(1).BillDate) + "~" + Num2Date$(UBBillRec(1).PastDueDate)
    ToPrint$ = ToPrint$ + "~" + Msg1$
    ToPrint$ = ToPrint$ + "~" + Msg2$
    ToPrint$ = ToPrint$ + "~" + Msg3$
    If DidADraftFlag Then
      ToPrint$ = ToPrint$ + "~" + "DRAFT NOTICE DO NOT PAY!!"
    ElseIf Len(Message$) > 0 Then
      ToPrint$ = ToPrint$ + "~" + Message$
    Else
      ToPrint$ = ToPrint$ + "~" + Msg4$
    End If

    If Totalamt# < 0 And FinalFlag Then
      If BLType = 1 Then
        ToPrint$ = ToPrint$ + "~" + "REFUND DUE" + "~" + Using("$#,###,###.##", Abs(Totalamt#))
      Else
        ToPrint$ = ToPrint$ + "~" + "REFUND DUE" + "~" + Using("$#,###,###.##", Totalamt#)
      End If
    ElseIf Totalamt# < 0 And Not FinalFlag Then
      ToPrint$ = ToPrint$ + "~" + "CREDIT BAL" + "~" + Using("$#,###,###.##", Totalamt#)
    Else
      ToPrint$ = ToPrint$ + "~" + "TOTAL DUE" + "~" + Using("$#,###,###.##", Totalamt#)
    End If
    ToPrint$ = ToPrint$ + "~" + Acct$ 'Using("##########", UBBillRec(1).CustAcctNo)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ServAddr, 26)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 29)
    ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).ADDR1)
    ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).ADDR2)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CITY, 14) + " " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
'    If FinalFlag Then
'      ToPrint$ = ToPrint$ + "~" + Using("#######.##", Round#(Totalamt#))
'    Else
    If Not Totalamt# > 0 Then
      ToPrint$ = ToPrint$ + "~ "
    Else
      ToPrint$ = ToPrint$ + "~" + Using("$#,###,###.##", Round#(Totalamt# + CustPenalty#))
    End If
    ToPrint$ = ToPrint$ + "~" + Zip$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    Print #UBRpt, ToPrint$
    ToPrint$ = ""
Return

PrnNewStandBarCode: 'use billing date and pastduedate
'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    lpcnt = 8
    If Not LPIFlag Then
      LPIFlag = -2
      If InStr(TOWNNAME$, "PEACHLAND") Then
        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
      Else
        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
      End If
    End If                                                 ':  M
    FoundAMtr = False

    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        FoundAMtr = True
        Exit For
      End If
    Next

    Select Case UBCustRec(1).LocMeters(1).MTRMulti
    Case 10
      Zero$ = "0"
    Case 100
      Zero$ = "00"
    Case 1000
      Zero$ = "000"
    Case Else
      Zero$ = ""
    End Select
    If UseEDateFlag = False Then
      If FoundAMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    AcctNum = UBBillRec(1).CustAcctNo

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If

    If FinalFlag And CDeposit# Then
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
    
'Need to get Penalty calculated based on defaults from setup screen
    If Totalamt# > 0 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
        If TotalTax > 0 Then
          CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
        Else
          CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
        End If
      Else
        CustPenalty# = 0
      End If
    Else
      CustPenalty# = 0
    End If
'-=-=-=-=-=-=-=-=-=-=-=-
    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, "~"; Tab(50); Using("########", (FBillNO& + PrintedCnt))
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(18); PrevDate$; Tab(33); DateRead$;
     'Only Print Days if Greater than 0
     If DaysINRead > 0 Then
       Print #UBRpt, "      "; Using("####", DaysINRead)
     Else
       Print #UBRpt, " "
     End If

    Print #UBRpt, " "
    Print #UBRpt, " "

    PCnt = 0
    For WRevCnt = 1 To lpcnt - 1
      PCnt = PCnt + 1
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
        If UBBillRec(1).CurRead(WRevCnt) > 0 Then
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          If UsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          End If
          Print #UBRpt, Tab(4); Using("#########", UBBillRec(1).PrevRead(WRevCnt));
          Print #UBRpt, Tab(13); Using("#########", UBBillRec(1).CurRead(WRevCnt));
          Print #UBRpt, Tab(23); Using("#######", UsageAmt&);
          Print #UBRpt, Zero$;
        End If
        Print #UBRpt, Tab(33); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));
      End If
      Select Case PCnt
      Case 1
        Print #UBRpt, Tab(45); Using("##########", UBBillRec(1).CustAcctNo)
      Case 5
        Print #UBRpt, Tab(45); Left$(UBCustRec(1).ServAddr, 26)
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, "Other";
        Print #UBRpt, Tab(33); Using("#####.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If

    If TotalTax# > 0 Then
      Print #UBRpt, Tab(14); "     TAX:"; Tab(31); Using("$###,###.##", TotalTax#)
'    Else
'      Print #UBRpt, ""
    End If
    If Previous# <> 0 Then
      Print #UBRpt, Tab(14); "Previous:"; Tab(31); Using("$###,###.##", Previous#)
'    Else
'      Print #UBRpt,
    End If
    Print #UBRpt, Tab(14); " Current:"; Tab(31); Using("$###,###.##", UBBillRec(1).Transamt)

    If FinalFlag And CDeposit# Then
      Print #UBRpt, Tab(14); " Deposit:"; Tab(31); Using("$###,###.##", -UBCustRec(1).DepositAmt);
    End If
    Print #UBRpt, Tab(45); Num2Date$(UBBillRec(1).BillDate); Tab(60); Num2Date$(UBBillRec(1).PastDueDate)
    Print #UBRpt, Tab(2); Msg1$
    Print #UBRpt, Tab(2); Msg2$
    Print #UBRpt, Tab(2); Msg3$;
    If Totalamt# < 0 And FinalFlag Then
      Print #UBRpt, Tab(35); "Refund:"; Tab(44); Using("$###,###.##", Abs(Totalamt#))
    Else
      Print #UBRpt, Tab(36); "Total:"; Tab(44); Using("$###,###.##", Totalamt#)
    End If

    If DidADraftFlag Then
      Print #UBRpt, Tab(2); "DRAFT NOTICE DO NOT PAY!!" ';
    ElseIf Len(Message$) > 0 Then
      Print #UBRpt, Tab(2); Message$ ';
    Else
      Print #UBRpt, Tab(2); Msg4$ ';
    End If

'    If Totalamt# < 0 And FinalFlag Then
'      Print #UBRpt, Tab(35); "Refund:"; Tab(44); Using("$###,###.##", Abs(Totalamt#))
'    Else
'      Print #UBRpt, Tab(36); "Total:"; Tab(44); Using("$###,###.##", Totalamt#)
'    End If

    Print #UBRpt, " "
')))))))))))))))))))))))
    If AcctBar = True Then
'*************For Okidata to print Bar code
      Print #UBRpt, Tab(55); Chr$(27); Chr$(16); "A"; 'String$(50, " ")
      Print #UBRpt, Chr$(8);
      Print #UBRpt, Chr$(2); "0";
      Print #UBRpt, "0"; Chr$(2);
      Print #UBRpt, Chr$(1); Chr$(1);
      Print #UBRpt, Chr$(1); Chr$(2);
      Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
'**************************
    Else
      Print #UBRpt, " "
    End If
'))))))))))))))))))))))))))
    Print #UBRpt, Tab(22); Left$(UBCustRec(1).CustName, 29)
    Print #UBRpt, Tab(22); QPTrim(UBCustRec(1).ADDR1)
    Print #UBRpt, Using("##########", UBBillRec(1).CustAcctNo); Tab(22); QPTrim(UBCustRec(1).ADDR2);
    Print #UBRpt, Tab(55); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(22); Left$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
    Print #UBRpt, " "
    Print #UBRpt, Using("$###,###.##", Totalamt#)
    'Print #UBRpt,
    Print #UBRpt, " "
    Print #UBRpt, " "
    If Totalamt# > 0 Then
'      Print #UBRpt, Using("#######.##", Round#(Totalamt#));
'    Else
      Print #UBRpt, Using("$###,###.##", Round#(Totalamt# + CustPenalty#));
    End If
    If PostBar = True Then
      Print #UBRpt, Tab(22); Chr$(27); Chr$(16); "C"; Chr$(Len(Zip$)); Zip$
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, " "
    Print #UBRpt, "~"
Return

PrnLaser: '7 'standard Laser w/possible 2read lines, Water/Elec
           'One message line and draft message line
           'use read date and due date
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
   
    lpcnt = 7
    WFoundMtr = False
    EFoundMtr = False
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        Select Case UBBillRec(1).MtrTypes(mChk)
        Case 1, 2, 3, 7
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          End If
          WFoundMtr = True
        Case 4, 5
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          If EUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
          End If
          EFoundMtr = True
        End Select
      End If
    Next

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round(UBBillRec(1).Transamt + Previous#)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If
    If FinalFlag And CDeposit# Then
      lpcnt = lpcnt - 1
    End If
    ToPrint$ = ToPrint$ + Using("####", FBillNO& + PrintedCnt)
    ToPrint$ = ToPrint$ + "~" + Left$(DateRead$, 2) + "~" + Mid$(DateRead$, 4, 2) + "~" + Right$(DateRead$, 2)
    ToPrint$ = ToPrint$ + "~" + Left$(DueDate$, 2) + "~" + Mid$(DueDate$, 4, 2) + "~" + Right$(DueDate$, 2) + "~" + DueDate$
    If EFoundMtr Then
      ToPrint$ = ToPrint$ + "~" + Using("########", EPrevRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", ECurrRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", EUsageAmt&)
      ToPrint$ = ToPrint$ + "~" + " E"
    Else
      ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    End If
    If WFoundMtr Then
      ToPrint$ = ToPrint$ + "~" + Using("########", WPrevRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", WCurrRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", WUsageAmt&)
      ToPrint$ = ToPrint$ + "~" + " W"
    Else
      ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    End If

    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ServAddr, 24)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ADDR1, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ADDR2, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CITY, 14) + " " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
    ToPrint$ = ToPrint$ + "~" + Str(CustAcctNo&)
'    If UBBillRec(1).RevAmt(1) <> 0 Then
'      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(1).REVNAME + "~" + Using("######.##", UBBillRec(1).RevAmt(1))
'    Else
'      ToPrint$ = ToPrint$ + "~ ~ "
'    End If
'    If UBBillRec(1).RevAmt(2) <> 0 Then
'      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(2).REVNAME + "~" + Using("######.##", UBBillRec(1).RevAmt(2))
'    Else
'      ToPrint$ = ToPrint$ + "~ ~ "
'    End If
'    If UBBillRec(1).RevAmt(3) <> 0 Then
'      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(3).REVNAME + "~" + Using("######.##", UBBillRec(1).RevAmt(3))
'    Else
'      ToPrint$ = ToPrint$ + "~ ~ "
'    End If
    For PCnt = 1 To lpcnt - 1
      If UBBillRec(1).RevAmt(PCnt) <> 0 Then
        ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(PCnt).RevName + "~" + Using("######.##", UBBillRec(1).RevAmt(PCnt))
      Else
        ToPrint$ = ToPrint$ + "~ ~ "
      End If
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        ToPrint$ = ToPrint$ + "~" + "Other" + "~" + Using("######.##", tmprev#)
      Else
        ToPrint$ = ToPrint$ + "~ ~ "
      End If
    If FinalFlag And CDeposit# Then
      ToPrint$ = ToPrint$ + "~" + "Deposit:" + "~" + Using("######.##", -UBCustRec(1).DepositAmt)
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
    If TotalTax# > 0 Then
      ToPrint$ = ToPrint$ + "~" + "TAX:" + "~" + Using("######.##", TotalTax#)
    End If

    If Previous# <> 0 Then
      ToPrint$ = ToPrint$ + "~" + "Previous:" + "~" + Using("######.##", Previous#)
    End If
    If DidADraftFlag Then
      ToPrint$ = ToPrint$ + "~" + "Account Will Be Drafted"
    Else
      ToPrint$ = ToPrint$ + "~ "
    End If
    ToPrint$ = ToPrint$ + "~" + Using("######.##", Totalamt#)
    If Len(Message$) > 0 Then
      ToPrint$ = ToPrint$ + "~" + Message$ + "~"
    Else
      ToPrint$ = ToPrint$ + "~" + Msg1$ + "~"
    End If
    BPrnCnt = BPrnCnt + 1
Dblcheck:
    If BPrnCnt = 3 Then
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    ElseIf BPrnCnt = 1 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$ + ToPrint2$
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    ElseIf BPrnCnt = 2 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    End If

Return

' Cashion   Laser Bills 3 per page
PrnLaserCashion:  ' 9 'special code for cashion for gas revenues...
                  'use read date and duedate
    WFoundMtr = False
    GFoundMtr = False
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        Select Case UBBillRec(1).MtrTypes(mChk)
        Case 1, 2, 3, 7
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          End If
          WFoundMtr = True
        Case 6
          GCurrRead& = UBBillRec(1).CurRead(mChk)
          GPrevRead& = UBBillRec(1).PrevRead(mChk)
          GUsageAmt& = GCurrRead& - GPrevRead&
          If GUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(GPrevRead&)) - 1)
            GUsageAmt& = (MaxMeterAmt& - GPrevRead&) + GCurrRead&
          End If
          GFoundMtr = True
        End Select
      End If
    Next
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round(UBBillRec(1).Transamt + Previous#)
'    If BPrnCnt > 0 Then
'      Print #UBRpt,
'
'      'PRINT #UBRpt,
'    Else
'      Print #UBRpt, Chr$(27); "&l3a"; Chr$(27); "&l6D"
'      Print #UBRpt,
'      'CASHION vvvvvv
'      'PRINT #UBRpt, CHR$(27); "&l3a"; CHR$(27); "&l6D";
'    End If
'show 2 reads
    ToPrint$ = ToPrint$ + Using("####", FBillNO& + PrintedCnt)
    ToPrint$ = ToPrint$ + "~" + Left$(DateRead$, 2) + "~" + Mid$(DateRead$, 4, 2) + "~" + Right$(DateRead$, 2)
    ToPrint$ = ToPrint$ + "~" + Left$(DueDate$, 2) + "~" + Mid$(DueDate$, 4, 2) + "~" + Right$(DueDate$, 2) + "~" + DueDate$
    If GFoundMtr Then
      ToPrint$ = ToPrint$ + "~" + Using("########", GPrevRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", GCurrRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", GUsageAmt&)
      ToPrint$ = ToPrint$ + "~" + " G"
    Else
      ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    End If
    If WFoundMtr Then
      ToPrint$ = ToPrint$ + "~" + Using("########", WPrevRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", WCurrRead&)
      ToPrint$ = ToPrint$ + "~" + Using("########", WUsageAmt&)
      ToPrint$ = ToPrint$ + "~" + " W"
    Else
      ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
    End If

'    IF GFoundMtr THEN
'      PRINT #UBRpt, USING "########"; GPrevRead&;
'      PRINT #UBRpt, TAB(9); USING "########"; GCurrRead&;
'      PRINT #UBRpt, TAB(18); USING "########"; GUsageAmt&;
'      PRINT #UBRpt, " G";
'    END IF

    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ServAddr, 24)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CustName, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ADDR1, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).ADDR2, 25)
    ToPrint$ = ToPrint$ + "~" + Left$(UBCustRec(1).CITY, 14) + " " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
    ToPrint$ = ToPrint$ + "~" + Str(CustAcctNo&)
    If UBBillRec(1).RevAmt(1) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(1).RevName + "~" + Using("######.##", UBBillRec(1).RevAmt(1))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If UBBillRec(1).RevAmt(2) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(2).RevName + "~" + Using("######.##", UBBillRec(1).RevAmt(2))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
'Specific for Cashion rev's 3,4,8,9
    GasTot# = 0
    GasTot# = Round#(UBBillRec(1).RevAmt(3) + UBBillRec(1).TaxAmt(3))
    GasTot# = Round#(GasTot# + UBBillRec(1).RevAmt(4) + UBBillRec(1).TaxAmt(4))
    GasTot# = Round#(GasTot# + UBBillRec(1).RevAmt(8) + UBBillRec(1).TaxAmt(8))
    GasTot# = Round#(GasTot# + UBBillRec(1).RevAmt(9) + UBBillRec(1).TaxAmt(9))

    If GasTot# <> 0 Then       'gas res.
      ToPrint$ = ToPrint$ + "~" + "GAS"
      ToPrint$ = ToPrint$ + "~" + Using("######.##", GasTot#)
    Else
    ToPrint$ = ToPrint$ + "~ ~ "
    End If
  'The code for GasTot above takes place of rev 3 and 4
    'IF UBBillRec(1).RevAmt(3) <> 0 THEN       'gas res.
    '  PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(3).RevName;
    '  PRINT #UBRpt, TAB(18); USING "######.##"; Round#(UBBillRec(1).RevAmt(3) + UBBillRec(1).TaxAmt(3));
    'ELSEIF UBBillRec(1).RevAmt(4) <> 0 THEN   'gas com.
    '  PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(4).RevName;
    '  PRINT #UBRpt, TAB(18); USING "######.##"; Round#(UBBillRec(1).RevAmt(4)+ UBBillRec(1).TaxAmt(4));
    'END IF


    If UBBillRec(1).RevAmt(5) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(5).RevName
      ToPrint$ = ToPrint$ + "~" + Using("######.##", UBBillRec(1).RevAmt(5))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
'special code for Cashion only
'Since Rev 3 and 4 are together can use rev 6 and deposit line
'Only shows 5 rev lines with prev and optional deposit/final
    If UBBillRec(1).RevAmt(6) <> 0 Then
      ToPrint$ = ToPrint$ + "~" + UBSetUpRec(1).Revenues(6).RevName
      ToPrint$ = ToPrint$ + "~" + Using("######.##", UBBillRec(1).RevAmt(6))
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If FinalFlag And CDeposit# Then
      ToPrint$ = ToPrint$ + "~" + "Deposit:" + "~" + Using("######.##", -UBCustRec(1).DepositAmt)
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If Previous# <> 0 Then
      ToPrint$ = ToPrint$ + "~" + "Previous:" + "~" + Using("######.##", Previous#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If DidADraftFlag Then
      ToPrint$ = ToPrint$ + "~" + "Account Will Be Drafted"
    Else
      ToPrint$ = ToPrint$ + "~ "
    End If
    ToPrint$ = ToPrint$ + "~" + Using("######.##", Totalamt#)
    If Len(Message$) > 0 Then
      ToPrint$ = ToPrint$ + "~" + Message$ + "~"
    Else
      ToPrint$ = ToPrint$ + "~" + Msg1$ + "~"
    End If
    BPrnCnt = BPrnCnt + 1
Dblcheck2:
    If BPrnCnt = 3 Then
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    ElseIf BPrnCnt = 1 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$ + ToPrint2$
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    ElseIf BPrnCnt = 2 And endit = True Then
      ToPrint$ = ToPrint$ + ToPrint2$
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
      BPrnCnt = 0
    End If

Return
PrintNewStandV1: 'use billing date, due and pastdue
    If InStr(TOWNNAME$, "RICH CREEK") Then
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      If Not LPIFlag Then
        LPIFlag = -2
        Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
        'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
        ' put printer in     8 lpi             12 cpi  oki mode
      End If
    End If
    If BLType = 4 Then
      lpcnt = 11
    ElseIf BLType = 18 Then
      lpcnt = 10
    End If
    FoundAMtr = False
    WFoundMtr = False
    EFoundMtr = False
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        Select Case UBBillRec(1).MtrTypes(mChk)
        Case 1, 2, 3, 7, 9
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
           End If
           'Added this to show zeros for mult on consumption amts
          WUsageAmt& = Round(WUsageAmt& * UBCustRec(1).LocMeters(mChk).MTRMulti)
          WFoundMtr = True
        Case 4, 5
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          If EUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
            End If
            'Added this to show zeros for mult on consumption amts
          EUsageAmt& = Round(EUsageAmt& * UBCustRec(1).LocMeters(mChk).MTRMulti)
          EFoundMtr = True
        End Select
      End If
    Next

'    For mChk = 1 To 7
'      If UBBillRec(1).MtrTypes(mChk) > 0 Then
''This code from Saltville
''        WPrevRead& = UBBillRec(1).PrevRead(1)
''        WCurrRead& = UBBillRec(1).CurRead(1)
''        WUsageAmt& = WCurrRead& - WPrevRead&
''        If WUsageAmt& < 0 Then
''          MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
''          WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
''        End If
'        FoundAMtr = True
'        Exit For
'      End If
'    Next
    If WFoundMtr = False And EFoundMtr = False Then
      FoundAMtr = False
    Else
      FoundAMtr = True
    End If
    If UseEDateFlag = False Then
      If FoundAMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If
    If FinalFlag And CDeposit# Then
      lpcnt = lpcnt - 1
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
'       If Totalamt# > 0 Then
'      'TenPct# = 2
'      'TenPct# = 0
'      'TenPct# = 10
'      'TenPct# = Round#(TotalAmt# * .1)
'      TenPct# = Round#(UBBillRec(1).Transamt * 0.1)
'    Else
'      TenPct# = 0
'    End If
    If Totalamt# > 0 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
        If TennRdg Then
          If TotalTax > 0 Then
            CalcPenamt Round(UBBillRec(1).RevAmt(1) + UBBillRec(1).RevAmt(2)), Previous, Round(Totalamt - TotalTax)
          Else
            CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
          End If
        ElseIf SpenTn Then
            CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
        Else
          If TotalTax > 0 Then
            CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
          Else
            CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
          End If
        End If
      Else
        CustPenalty# = 0
      End If
    Else
      CustPenalty# = 0
    End If
    AcctNum = UBBillRec(1).CustAcctNo

    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, "~"; Tab(50); Using("########", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    If BLType = 18 Then
      Print #UBRpt, " "
    End If
    Print #UBRpt, " "
    Print #UBRpt, Using("##########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(15); Left$(UBCustRec(1).ServAddr, 31); Tab(50); Using("########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(62); Num2Date$(UBBillRec(1).BillDate)
    Print #UBRpt, " "
    If Not BFlag Then
      Print #UBRpt, Tab(50); DueDate$; Tab(64); Using("#####.##", Totalamt#)
      Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate); Tab(15); PrevDate$; Tab(26); DateRead$;
     'Only Print Days if Greater than 0
      If DaysINRead > 0 Then
        Print #UBRpt, Tab(39); Using("####", DaysINRead)
      Else
        Print #UBRpt, " "
      End If
    Else
      Print #UBRpt, Tab(50); DueDate$; Tab(64); Using("#####.##", Totalamt#)
      Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate);
      Print #UBRpt, " "
    End If
    If BLType = 18 Then
      Print #UBRpt, " "
    End If
    Print #UBRpt, Tab(50); Num2Date$(UBBillRec(1).PastDueDate);
    If Not Totalamt# > 0 Then
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt#))
    Else
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt# + CustPenalty#))
    End If
    If Not BLType = 18 Then
      Print #UBRpt, " " 'line 13
     Else
      Print #UBRpt,
     End If
'    Print #UBRpt, " "
'    Print #UBRpt, " "
'    Print #UBRpt, Tab(47); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
'Saltville code to summarize first two meter info
'    PCnt = 0
'    For WRevCnt = 1 To 7
'      PCnt = PCnt + 1
'      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
'        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).REVNAME, 5);
'        If WRevCnt < 3 Then
'          Print #UBRpt, Tab(7); Using("##########", CurrRead&);
'          Print #UBRpt, Tab(17); Using("##########", PrevRead&);
'          Print #UBRpt, Tab(28); Using("#######", UsageAmt&);
'        End If
'      Print #UBRpt, Tab(36); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));
'      End If
      
    PCnt = 0
    For WRevCnt = 1 To lpcnt - 1
      PCnt = PCnt + 1  'Printable lines
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
'        If UBBillRec(1).CurRead(WRevCnt) > 0 Then
'          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'          If UsageAmt& < 0 Then
'            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
'            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'          End If
'          Print #UBRpt, Tab(7); Using("##########", UBBillRec(1).CurRead(WRevCnt));
'          Print #UBRpt, Tab(17); Using("##########", UBBillRec(1).PrevRead(WRevCnt));
'          Print #UBRpt, Tab(28); Using("#######", UsageAmt&);
      

        Select Case PCnt
        Case 1, 2  'water/sewer
          If WFoundMtr Then
            If Not BFlag Then
              Print #UBRpt, Tab(7); Using("##########", WCurrRead&);
              Print #UBRpt, Tab(17); Using("##########", WPrevRead&);
              Print #UBRpt, Tab(28); Using("#######", WUsageAmt&);
            End If
         End If
        End Select
'       End If
       Print #UBRpt, Tab(36); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));

      End If
      Select Case PCnt
      Case 2
        'print bar code acct
        ')))))))))))))))))))))))
    If AcctBar = True Then
'*************For Okidata to print Bar code
      Print #UBRpt, Tab(47); Chr$(27); Chr$(16); "A";
      Print #UBRpt, Chr$(8);
      Print #UBRpt, Chr$(2); "0";
      Print #UBRpt, "0"; Chr$(2);
      Print #UBRpt, Chr$(1); Chr$(1);
      Print #UBRpt, Chr$(1); Chr$(2);
      Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
'**************************
    Else
      Print #UBRpt, " "
    End If
'))))))))))))))))))))))))))

      Case 5
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CustName, 29)
      Case 6
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR1)
      Case 7
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR2)
      Case 8
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CITY, 14); " "; QPTrim$(UBCustRec(1).STATE); " "; Left$(UBCustRec(1).ZIPCODE, 5)
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, "                     Other:  ";
        Print #UBRpt, Tab(36); Using("#####.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If
    If TotalTax# > 0 Then
      Print #UBRpt, "                       TAX:  "; Tab(34); Using("$###,###.##", TotalTax#)
'    Else
'      Print #UBRpt, " "
    End If
     If Previous# <> 0 Then
      Print #UBRpt, "                  Previous:  "; Tab(34); Using("$###,###.##", Previous#)
'    Else
'      Print #UBRpt, " "
    End If
    Print #UBRpt, "                   Current:  "; Tab(34); Using("$###,###.##", UBBillRec(1).Transamt)
    Print #UBRpt, Tab(32); "--------------"
'line 24
    If FinalFlag And CDeposit# Then
      Print #UBRpt, "                   Deposit:  "; Tab(34); Using("$###,###.##", -UBCustRec(1).DepositAmt)
'    Else
'      Print #UBRpt, " "
    End If

    If Totalamt# < 0 And FinalFlag Then
      Print #UBRpt, "                Refund Due:  "; Tab(34); Using("$###,###.##", Abs(Totalamt#))
    Else
      Print #UBRpt, "                     Total:  "; Tab(34); Using("$###,###.##", Totalamt#)
    End If
    If Not BLType = 18 Then
      Print #UBRpt, " "
    End If
    Print #UBRpt, Tab(3); Msg1$; Tab(47); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(3); Msg2$;
    
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT NOTICE DO NOT PAY!!"
    Else
      Print #UBRpt, " "
    End If
    If BLType = 18 Then
      If Len(Message$) > 0 Then
        Print #UBRpt, Tab(3); Message$;
      Else
        Print #UBRpt, Tab(3); Msg3$;
      End If
    Else
      Print #UBRpt, Tab(3); Msg3$;
    End If
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    Else
      Print #UBRpt, " "
    End If
    If Not BLType = 18 Then
      If Len(Message$) > 0 Then
        Print #UBRpt, Tab(3); Message$;
      Else
        Print #UBRpt, Tab(3); Msg4$;
      End If
      If Rteflag Then
        Print #UBRpt, Tab(47); "Route: "; UBCustRec(1).POSTRTE
      Else
        Print #UBRpt, " "
      End If
    End If
    If Not BLType = 18 Then
      Print #UBRpt, " "
    End If
    Print #UBRpt, "~"
Return
PrintNewStandRmStamp: 'use billingdate, due and pastdue
    If Not LPIFlag Then
      LPIFlag = -2
      OkiMode = LPIFlag
      'PRINT #UBRpt, CHR$(27); "&k4S"; CHR$(27); "&l8D";
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      'put printer in     8 lpi             12 cpi  oki mode
    End If
    FoundAMtr = False
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        FoundAMtr = True
        Exit For
      End If
    Next
    If UseEDateFlag = False Then
      If FoundAMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If FinalFlag And CDeposit# Then
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
    If Totalamt# > 0 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
        If TotalTax > 0 Then
          CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
        Else
          CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
        End If
      Else
        CustPenalty# = 0
      End If
'      TenPct# = Round#(Totalamt# * 0.1)
'    Else
'      TenPct# = 0
    Else
      CustPenalty# = 0
    End If

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, "~"; Tab(50); Using("########", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Using("##########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(15); Left$(UBCustRec(1).ServAddr, 26); Tab(50); Using("########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(62); Num2Date$(UBBillRec(1).BillDate)
    Print #UBRpt, " "

    Print #UBRpt, Tab(50); DueDate$; Tab(64); Using("#####.##", Totalamt#)
    Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate); Tab(15); PrevDate$; Tab(26); DateRead$;
     'Only Print Days if Greater than 0
     If DaysINRead > 0 Then
       Print #UBRpt, Tab(40); Using("####", DaysINRead)
     Else
       Print #UBRpt, " "
     End If

    'Print #UBRpt, Tab(50); QPTrim$(Message$);
'unrem
    If Not Totalamt# > 0 Then
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt#))
    Else
      Print #UBRpt, Tab(50); Num2Date$(UBBillRec(1).PastDueDate); Tab(64); Using("#####.##", Round#(Totalamt# + CustPenalty#))
    End If

    Print #UBRpt, " "
    'PRINT #UBRpt, STRING$(50, " "); CHR$(27); CHR$(16); "A";
    'PRINT #UBRpt, CHR$(8);
    'PRINT #UBRpt, CHR$(2); CHR$(0);
    'PRINT #UBRpt, CHR$(0); CHR$(2);
    'PRINT #UBRpt, CHR$(1); CHR$(1);
    'PRINT #UBRpt, CHR$(1); CHR$(2);
    'PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(AcctLen); Acct$

    PCnt = 0
    For WRevCnt = 1 To 6
      PCnt = PCnt + 1
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        If UBBillRec(1).CurRead(WRevCnt) > 0 Then
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          If UsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          End If
          Print #UBRpt, Tab(7); Using("##########", UBBillRec(1).CurRead(WRevCnt));
          Print #UBRpt, Tab(17); Using("##########", UBBillRec(1).PrevRead(WRevCnt));
          Print #UBRpt, Tab(28); Using("#######", UsageAmt&);
        End If
        Print #UBRpt, Tab(36); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));
      End If
      Select Case PCnt
      Case 2
        'print bar code acct
        ')))))))))))))))))))))))
        If AcctBar = True Then
      '*************For Okidata to print Bar code
          Print #UBRpt, Tab(47); Chr$(27); Chr$(16); "A";
          Print #UBRpt, Chr$(8);
          Print #UBRpt, Chr$(2); "0";
          Print #UBRpt, "0"; Chr$(2);
          Print #UBRpt, Chr$(1); Chr$(1);
          Print #UBRpt, Chr$(1); Chr$(2);
          Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
      '**************************
        Else
          Print #UBRpt, " "
        End If
      '))))))))))))))))))))))))))
      Case 5
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CustName, 29)
      Case 6
        Print #UBRpt, Tab(47); UBCustRec(1).ADDR1
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
      For PCnt = 7 To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, "                     Other:  ";
        Print #UBRpt, Tab(36); Using("#####.##", tmprev#);
      End If
      Print #UBRpt, Tab(47); UBCustRec(1).ADDR2
    If TotalTax# > 0 Then
      Print #UBRpt, "                       TAX:  "; Tab(34); Using("$###,###.##", TotalTax#);
    End If
    Print #UBRpt, Tab(47); Left$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; Left$(UBCustRec(1).ZIPCODE, 5)
    If Previous# <> 0 Then
      Print #UBRpt, "                  Previous:  "; Tab(34); Using("$###,###.##", Previous#);
    End If
    Print #UBRpt, " "
    Print #UBRpt, "                   Current:  "; Tab(34); Using("$###,###.##", UBBillRec(1).Transamt)
    Print #UBRpt, Tab(32); "--------------"

    If Totalamt# < 0 And FinalFlag Then
      Print #UBRpt, "                Refund Due:  "; Tab(34); Using("$###,###.##", Abs(Totalamt#))
    Else
      Print #UBRpt, "                     Total:  "; Tab(34); Using("$###,###.##", Totalamt#)
    End If
    If FinalFlag And CDeposit# Then
      Print #UBRpt, "                   Deposit:  "; Tab(34); Using("$###,###.##", -UBCustRec(1).DepositAmt)
    Else
      Print #UBRpt, " "
    End If

    Print #UBRpt, Tab(3); Msg1$; Tab(47); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(3); Msg2$;
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT NOTICE DO NOT PAY!!"
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, Tab(3); Msg3$;
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    Else
      Print #UBRpt, " "
    End If
    If Len(Message$) > 0 Then
      Print #UBRpt, Tab(3); Message$
    Else
      Print #UBRpt, Tab(3); Msg4$
    End If
    Print #UBRpt, Tab(3); "~"
Return
PrnStand24L2Bx: '6  'use read date and due date
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  ZDigit$ = GetZipEDigit$(Zip$)
  Zip$ = Zip$ + ZDigit$
   
    lpcnt = 6
    WFoundMtr = False
    EFoundMtr = False
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        Select Case UBBillRec(1).MtrTypes(mChk)
        Case 1, 2, 3, 7
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          End If
          WFoundMtr = True
        Case 4, 5
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          If EUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
          End If
          EFoundMtr = True
        End Select
      End If
    Next

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round(UBBillRec(1).Transamt + Previous#)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If
    If FinalFlag And CDeposit# Then
      lpcnt = lpcnt - 1
    End If
    If UseEDateFlag = False Then
      If WFoundMtr = False And EFoundMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    Zero$ = "0"

    AcctNum = UBBillRec(1).CustAcctNo
    If FinalFlag And CDeposit# Then
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If

    If Totalamt# > 0 And BLType = 7 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
        If TotalTax > 0 Then
          CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
        Else
          CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
        End If
      Else
        CustPenalty# = 0
      End If
    Else
      CustPenalty# = 0
    End If

    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, "~"; Using("#####", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(4); Left$(DateRead$, 2); Tab(9); Mid$(DateRead$, 4, 2); Tab(14); Right$(DateRead$, 2);
    Print #UBRpt, Tab(18); Left$(DueDate$, 2); Tab(23); Mid$(DueDate$, 4, 2); Tab(28); Right$(DueDate$, 2);
    Print #UBRpt, " "
    Print #UBRpt, Tab(40); DueDate$
    Print #UBRpt, " "
    'Print #UBRpt, " "
    If WFoundMtr Then
      Print #UBRpt, Tab(2); Using("#########", WPrevRead&);
      Print #UBRpt, Tab(12); Using("#########", WCurrRead&);
      Print #UBRpt, Tab(22); Using("########", WUsageAmt&)
    Else
      Print #UBRpt, " "
    End If
    If EFoundMtr Then
      Print #UBRpt, Tab(2); Using("#########", EPrevRead&);
      Print #UBRpt, Tab(12); Using("#########", ECurrRead&);
      Print #UBRpt, Tab(22); Using("########", EUsageAmt&);
    End If
    Print #UBRpt, Tab(35); Left$(QPTrim$(UBCustRec(1).ServAddr), 24)
    Print #UBRpt, " "
    Print #UBRpt, Tab(45); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).CustName), 25)
    PCnt = 0
    For WRevCnt = 1 To lpcnt
      PCnt = PCnt + 1
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, Tab(3); UBSetUpRec(1).Revenues(WRevCnt).RevName; Tab(20); Using("######.##", UBBillRec(1).RevAmt(WRevCnt));
      End If
    
    Select Case PCnt:
    Case 1:
      Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).ADDR1), 25)
    Case 2:
      Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).ADDR2), 25)
    Case 3:
      Print #UBRpt, Tab(34); Left$(QPTrim$(UBCustRec(1).CITY), 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
'    Case 4:
'      Print #UBRpt,
    Case Else:
      Print #UBRpt, " " 'TAB(34); STRING$(24, "-")
    End Select
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, Tab(3); "   Other:";
        Print #UBRpt, Tab(20); Using("######.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If
    If TotalTax# <> 0 Then
      Print #UBRpt, Tab(3); "     TAX:"; Tab(20); Using("######.##", TotalTax#)
    End If
    If Previous# <> 0 Then
      Print #UBRpt, Tab(3); "Previous:"; Tab(20); Using("######.##", Previous#)
'    Else
'      Print #UBRpt, " "
    End If

    If FinalFlag And CDeposit# Then
      Print #UBRpt, Tab(4); "Deposit:"; Tab(20); Using("######.##", -UBCustRec(1).DepositAmt)
      'Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
'        Else
'      Print #UBRpt, " "
    End If

    Print #UBRpt,
    Print #UBRpt, Tab(8); CustAcctNo&; Tab(21); Using("#####.##", Totalamt#);
    If BLType = 6 Then
      Print #UBRpt, Tab(38); CustAcctNo&; Tab(51); Using("#####.##", Totalamt#)
    Else
      If Totalamt# > 0 Then
        Print #UBRpt, Tab(35); CustAcctNo&; Tab(42); Using("#####.##", Totalamt#); Tab(52); Using("#####.##", Totalamt# + CustPenalty#)
      Else
        Print #UBRpt, Tab(35); CustAcctNo&; Tab(42); Using("#####.##", Totalamt#)
      End If
    End If
    Print #UBRpt, " "; Msg1$; " "; Msg2$
    If Len(Message$) > 0 Then
      Print #UBRpt, " "; Message$;
    Else
      Print #UBRpt, " "; Msg3$;
    End If
    If DidADraftFlag Then
      Print #UBRpt, " "; "DRAFT NOTICE DO NOT PAY!!"
    Else
      Print #UBRpt, " "; Msg4$
    End If
    Print #UBRpt, "~"
Return
PrnStand24L3Bx:
  GoSub PrnStand24L2Bx
Return

'Prints in 10cpi
'Will list 1 thru 5 revenues then 6 to 15 totaled under other
PrnStand21Line:  '8 and 12 and 14 'Old Standard 21 Line
                  'use read date and past due
    If UBBillRec(1).CurRead(1) >= 0 And UBBillRec(1).PrevRead(1) >= 0 Then
      UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
      If UsageAmt& < 0 Then
        MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(1))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
      End If
    Else
      UsageAmt& = 0
    End If
'Added this to show zeros for mult on consumption amts
    UsageAmt& = Round(UsageAmt& * UBCustRec(1).LocMeters(1).MTRMulti)
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round(UBBillRec(1).Transamt + Previous#)

    Print #UBRpt, "~"; Using("#####", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    'PRINT #UBRpt, TAB(3); LEFT$(BillDate$, 2); TAB(8); MID$(BillDate$, 4, 2); TAB(13); RIGHT$(BillDate$, 2);
    Print #UBRpt, Tab(3); Left$(DateRead$, 2); Tab(8); Mid$(DateRead$, 4, 2); Tab(13); Right$(DateRead$, 2);
    Print #UBRpt, Tab(17); Left$(PastDueDate$, 2); Tab(22); Mid$(PastDueDate$, 4, 2); Tab(27); Right$(PastDueDate$, 2);
    If BLType = 12 Then
      Print #UBRpt, Tab(32); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Else
      Print #UBRpt, Tab(40); PastDueDate$
    End If
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(2); Using("#########", UBBillRec(1).PrevRead(1));
    Print #UBRpt, Tab(12); Using("#########", UBBillRec(1).CurRead(1));
    Print #UBRpt, Tab(22); Using("########", UsageAmt&);
    Print #UBRpt, Tab(35); Left$(QPTrim(UBCustRec(1).ServAddr), 24)
    Print #UBRpt, " "
    Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).CustName), 25)

    If UBBillRec(1).RevAmt(1) <> 0 Then
      Print #UBRpt, Tab(3); UBSetUpRec(1).Revenues(1).RevName; Tab(20); Using("######.##", UBBillRec(1).RevAmt(1));
    End If
    Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).ADDR1), 25)

    If UBBillRec(1).RevAmt(2) <> 0 Then
      Print #UBRpt, Tab(3); UBSetUpRec(1).Revenues(2).RevName; Tab(20); Using("######.##", UBBillRec(1).RevAmt(2));
    End If
    Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).ADDR2), 25)

    If UBBillRec(1).RevAmt(3) <> 0 Then
      Print #UBRpt, Tab(3); UBSetUpRec(1).Revenues(3).RevName;
      Print #UBRpt, Tab(20); Using("######.##", UBBillRec(1).RevAmt(3));
    End If
    Print #UBRpt, Tab(34); Left$(QPTrim(UBCustRec(1).CITY), 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE

    If UBBillRec(1).RevAmt(4) <> 0 Then
      Print #UBRpt, Tab(3); UBSetUpRec(1).Revenues(4).RevName;
      Print #UBRpt, Tab(20); Using("######.##", UBBillRec(1).RevAmt(4));
    End If
    Print #UBRpt, Tab(34); String$(24, "-")

    If UBBillRec(1).RevAmt(5) <> 0 Then
      Print #UBRpt, Tab(3); UBSetUpRec(1).Revenues(5).RevName;
      Print #UBRpt, Tab(20); Using("######.##", UBBillRec(1).RevAmt(5));
    End If
    If DidADraftFlag Then
      Print #UBRpt, Tab(34); "Account Will Be Drafted"
    Else
      If Len(Message$) > 0 Then
        Print #UBRpt, Tab(34); Message$
      Else
        Print #UBRpt, Tab(34); Msg1$
      End If
    End If

    'Print #UBRpt, Tab(34); Message$
'if not this then will get a line below the previous....
    If Not FinalFlag And Not CDeposit# Then
      tmprev# = 0
      For PCnt = 6 To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, Tab(3); "Other";
        Print #UBRpt, Tab(20); Using("######.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If
    End If
    If Previous# <> 0 Then
      Print #UBRpt, Tab(3); "Previous:"; Tab(20); Using("######.##", Previous#)
    Else
      Print #UBRpt, " "
    End If

    If FinalFlag And CDeposit# Then
      Print #UBRpt, Tab(4); "Deposit:"; Tab(20); Using("######.##", -UBCustRec(1).DepositAmt);
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
      Print #UBRpt, " " 'TAB(50); USING "######.##"; TotalAmt#
    ElseIf FinalFlag And Not CDeposit# Then
      Print #UBRpt, " "
    End If
    Print #UBRpt, " "
    Print #UBRpt, Tab(5); CustAcctNo&; Tab(20); Using("######.##", Totalamt#);
    If BLType = 14 Then
'      Print #UBRpt, Tab(38); CustAcctNo&; Tab(51); Using("#####.##", Totalamt#)
'    Else
      If Totalamt# > 0 Then
        Print #UBRpt, Tab(35); CustAcctNo&; Tab(42); Using("#####.##", Totalamt#); Tab(52); Using("#####.##", Totalamt# + CustPenalty#)
      Else
        Print #UBRpt, Tab(35); CustAcctNo&; Tab(42); Using("#####.##", Totalamt#)
      End If
    Else
      Print #UBRpt, Tab(37); CustAcctNo&; Tab(50); Using("######.##", Totalamt#)
    End If
    Print #UBRpt, "~"
Return
    '***
'This is for Sunset Beach So Rev Irrigation will show reads
PrintNewStandV1SB:   'Bill 15
                    'use billing date, due and pastdue
    If Not LPIFlag Then
      LPIFlag = -2
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    End If
    lpcnt = 11
    FoundAMtr = False
    For mChk = 1 To 7
      SFoundMtr(mChk) = False
    Next
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
          SCurrRead&(mChk) = UBBillRec(1).CurRead(mChk)
          SPrevRead&(mChk) = UBBillRec(1).PrevRead(mChk)
          SUsageAmt&(mChk) = SCurrRead&(mChk) - SPrevRead&(mChk)
          'SUsageAmt&(mChk) = Round(SUsageAmt&(mChk) * UBCustRec(1).LocMeters(mChk).MTRMulti)
          If SUsageAmt&(mChk) < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(SPrevRead&(mChk))) - 1)
            SUsageAmt&(mChk) = (MaxMeterAmt& - SPrevRead&(mChk)) + SCurrRead&(mChk)
          End If
          SUsageAmt&(mChk) = Round(SUsageAmt&(mChk) * UBCustRec(1).LocMeters(mChk).MTRMulti)
          Select Case UBBillRec(1).MtrTypes(mChk)
          Case 1
            SMtrType(mChk) = "W"
          Case 2
            SMtrType(mChk) = "S"
          Case 3
            SMtrType(mChk) = "C"
          Case 4
            SMtrType(mChk) = "E"
          Case 5
            SMtrType(mChk) = "D"
          Case 6
            SMtrType(mChk) = "G"
          Case 7
            SMtrType(mChk) = "T"
          Case 9
            SMtrType(mChk) = "I"
          End Select
          SFoundMtr(mChk) = True
      End If
    Next
    For mChk = 1 To 7
      If SFoundMtr(mChk) = False Then
        FoundAMtr = False
      Else
        FoundAMtr = True
        Exit For
      End If
    Next
    If UseEDateFlag = False Then
      If FoundAMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If
    If FinalFlag And CDeposit# Then
      lpcnt = lpcnt - 1
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
'       If Totalamt# > 0 Then
'      'TenPct# = 2
'      'TenPct# = 0
'      'TenPct# = 10
'      'TenPct# = Round#(TotalAmt# * .1)
'      TenPct# = Round#(UBBillRec(1).Transamt * 0.1)
'    Else
'      TenPct# = 0
'    End If
    If Totalamt# > 0 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
        If TotalTax > 0 Then
          CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
        Else
          CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
        End If
      Else
        CustPenalty# = 0
      End If
    Else
      CustPenalty# = 0
    End If

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, "~"; Tab(50); Using("########", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Using("##########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(15); Left$(QPTrim(UBCustRec(1).ServAddr), 31); Tab(50); Using("########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(62); Num2Date$(UBBillRec(1).BillDate)
    Print #UBRpt, " "

    Print #UBRpt, Tab(50); DueDate$; Tab(64); Using("#####.##", Totalamt#)
    If Not BFlag Then
      Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate); Tab(15); PrevDate$; Tab(26); DateRead$;
       'Only Print Days if Greater than 0
       If DaysINRead > 0 Then
         Print #UBRpt, Tab(40); Using("####", DaysINRead)
       Else
         Print #UBRpt, " "
       End If
    Else
      Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate);
      Print #UBRpt, " "
    End If

    Print #UBRpt, Tab(50); Num2Date$(UBBillRec(1).PastDueDate);
    If Not Totalamt# > 0 Then
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt#))
    Else
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt# + CustPenalty#))
    End If

    Print #UBRpt, " " 'line 13
'    Print #UBRpt, " "
'    Print #UBRpt, " "
'    Print #UBRpt, Tab(47); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
'Saltville code to summarize first two meter info
'    PCnt = 0
'    For WRevCnt = 1 To 7
'      PCnt = PCnt + 1
'      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
'        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).REVNAME, 5);
'        If WRevCnt < 3 Then
'          Print #UBRpt, Tab(7); Using("##########", CurrRead&);
'          Print #UBRpt, Tab(17); Using("##########", PrevRead&);
'          Print #UBRpt, Tab(28); Using("#######", UsageAmt&);
'        End If
'      Print #UBRpt, Tab(36); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));
'      End If
      
    PCnt = 0
    For WRevCnt = 1 To lpcnt - 1
      PCnt = PCnt + 1  'Printable lines
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
'        If UBBillRec(1).CurRead(WRevCnt) > 0 Then
'          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'          If UsageAmt& < 0 Then
'            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
'            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'          End If
'          Print #UBRpt, Tab(7); Using("##########", UBBillRec(1).CurRead(WRevCnt));
'          Print #UBRpt, Tab(17); Using("##########", UBBillRec(1).PrevRead(WRevCnt));
'          Print #UBRpt, Tab(28); Using("#######", UsageAmt&);
        For mChk = 1 To 7
          If UBBillRec(1).MtrTypes(mChk) > 0 Then
            If UBCustRec(1).serv(WRevCnt).RMtrType = SMtrType(mChk) Then
              If Not BFlag Then
                Print #UBRpt, Tab(7); Using("##########", SCurrRead&(mChk));
                Print #UBRpt, Tab(17); Using("##########", SPrevRead&(mChk));
                Print #UBRpt, Tab(28); Using("#######", SUsageAmt&(mChk));
              End If
              Exit For
            End If
          End If
        Next
'        Select Case PCnt
'        Case 1, 2  'water/sewer
'          If WFoundMtr Then
'            Print #UBRpt, Tab(7); Using("##########", WCurrRead&);
'            Print #UBRpt, Tab(17); Using("##########", WPrevRead&);
'            Print #UBRpt, Tab(28); Using("#######", WUsageAmt&);
'          End If
'        End Select
'       End If
       Print #UBRpt, Tab(36); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));

      End If
      Select Case PCnt
      Case 2
        'print bar code acct
        ')))))))))))))))))))))))
    If AcctBar = True Then
'*************For Okidata to print Bar code
      Print #UBRpt, Tab(47); Chr$(27); Chr$(16); "A";
      Print #UBRpt, Chr$(8);
      Print #UBRpt, Chr$(2); "0";
      Print #UBRpt, "0"; Chr$(2);
      Print #UBRpt, Chr$(1); Chr$(1);
      Print #UBRpt, Chr$(1); Chr$(2);
      Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
'**************************
    Else
      Print #UBRpt, " "
    End If
'))))))))))))))))))))))))))

      Case 5
        Print #UBRpt, Tab(47); Left$(QPTrim(UBCustRec(1).CustName), 29)
      Case 6
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR1)
      Case 7
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR2)
      Case 8
        Print #UBRpt, Tab(47); Left$(QPTrim(UBCustRec(1).CITY), 14); " "; QPTrim$(UBCustRec(1).STATE); " "; Left$(UBCustRec(1).ZIPCODE, 5)
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, "                     Other:  ";
        Print #UBRpt, Tab(36); Using("#####.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If
    If TotalTax# > 0 Then
      Print #UBRpt, "                       TAX:  "; Tab(34); Using("$###,###.##", TotalTax#)
'    Else
'      Print #UBRpt, " "
    End If
     If Previous# <> 0 Then
      Print #UBRpt, "                  Previous:  "; Tab(34); Using("$###,###.##", Previous#)
'    Else
'      Print #UBRpt, " "
    End If
    Print #UBRpt, "                   Current:  "; Tab(34); Using("$###,###.##", UBBillRec(1).Transamt)
    Print #UBRpt, Tab(32); "--------------"
'line 24
    If FinalFlag And CDeposit# Then
      Print #UBRpt, "                   Deposit:  "; Tab(34); Using("$###,###.##", -UBCustRec(1).DepositAmt)
'    Else
'      Print #UBRpt, " "
    End If

    If Totalamt# < 0 And FinalFlag Then
      Print #UBRpt, "                Refund Due:  "; Tab(34); Using("$###,###.##", Abs(Totalamt#))
    Else
      Print #UBRpt, "                     Total:  "; Tab(34); Using("$###,###.##", Totalamt#)
    End If
    Print #UBRpt, " "
    Print #UBRpt, Tab(3); Msg1$; Tab(47); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(3); Msg2$;
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT NOTICE DO NOT PAY!!"
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, Tab(3); Msg3$;
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    Else
      Print #UBRpt, " "
    End If
    If Len(Message$) > 0 Then
      Print #UBRpt, Tab(3); Message$
    Else
      Print #UBRpt, Tab(3); Msg4$
    End If
    Print #UBRpt, " "
    Print #UBRpt, "~"
Return


GetOut:

ExitPrintBill:
  UBLog "OUT: Bill Printing."

End Sub
Private Sub PrnLaserLetterForm(UBCustRec() As NewUBCustRecType, UBBillRec() As UBTransRecType, UBSetUpRec() As UBSetupRecType)  'Bill 16
'Utility Bill LaserLetterFormat(Uses Blank Stock)BAR CODE PRINTABLE
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  'ZDigit$ = GetZipEDigit$(Zip$)
  Zip$ = Zip$ '+ ZDigit$
  FoundAMtr = False
  For mChk = 1 To 7
    If UBBillRec(1).MtrTypes(mChk) > 0 Then
      FoundAMtr = True
      Exit For
    End If
  Next
  Select Case UBCustRec(1).LocMeters(1).MTRMulti
  Case 10
    Zero$ = "0"
  Case 100
    Zero$ = "00"
  Case 1000
    Zero$ = "000"
  Case Else
    Zero$ = ""
  End Select
  If UseEDateFlag = False Then
    If FoundAMtr = False Then
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    End If
  End If
  AcctNum = UBBillRec(1).CustAcctNo
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
'  If Previous# <> 0 Then
'  End If
'  If TotalTax > 0 Then
'    lpcnt = lpcnt - 1
'  End If

  If FinalFlag And CDeposit# Then
    Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
  End If
'Need to get Penalty calculated based on defaults from setup screen
  If Totalamt# > 0 And Not FinalFlag Then
    If UBCustRec(1).LATEFEE = "Y" Then
      If TotalTax > 0 Then
        CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
      Else
        CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
      End If
    Else
      CustPenalty# = 0
    End If
  Else
    CustPenalty# = 0
  End If
'-=-=-=-=-=-=-=-=-=-=-=-
  AcctNum = UBBillRec(1).CustAcctNo
  Acct$ = QPTrim$(Str$(AcctNum))
  Select Case AcctNum
  Case Is < 10
    Acct$ = "00" + Acct$
  Case Is < 100
    Acct$ = "0" + Acct$
  End Select
  ToPrint$ = Using("########", (FBillNO& + PrintedCnt))
  ToPrint$ = ToPrint$ + "~" + PrevDate$ + "~" + DateRead$
   'Only Print Days if Greater than 0
   If DaysINRead > 0 Then
     ToPrint$ = ToPrint$ + "~" + Using("####", DaysINRead)
   Else
     ToPrint$ = ToPrint$ + "~ "
   End If
  PCnt = 0
  For PCnt = 1 To 7
        If UBBillRec(1).CurRead(PCnt) > 0 Then
          UsageAmt& = UBBillRec(1).CurRead(PCnt) - UBBillRec(1).PrevRead(PCnt)
          If UsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(PCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(PCnt)) + UBBillRec(1).CurRead(PCnt)
          End If
          UsageAmt& = Round(UsageAmt& * UBCustRec(1).LocMeters(PCnt).MTRMulti)
          If MtrNFlag = 2 Then
            ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).LocMeters(PCnt).MtrIDNO)
          Else
            ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).LocMeters(PCnt).MtrNum)
          End If
          ToPrint$ = ToPrint$ + "~" + Using("#########", UBBillRec(1).PrevRead(PCnt))
          ToPrint$ = ToPrint$ + "~" + Using("#########", UBBillRec(1).CurRead(PCnt))
          ToPrint$ = ToPrint$ + "~" + Using("#######", UsageAmt&)
        Else
          ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
        End If
  Next
    For WRevCnt = 1 To 14
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBSetUpRec(1).Revenues(WRevCnt).RevName)
        ToPrint$ = ToPrint$ + "~" + Using("#####.##", UBBillRec(1).RevAmt(WRevCnt))
      Else
        ToPrint$ = ToPrint$ + "~ ~ "
      End If
    Next
    ToPrint$ = ToPrint$ + "~ ~ "
'    tmprev# = 0
'      For PCnt = lpcnt To 15
'        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
'          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
'        End If
'      Next
'      If tmprev# <> 0 Then
'        ToPrint$ = ToPrint$ + "~ ~ ~ ~" + "   Other:"
'        ToPrint$ = ToPrint$ + "~" + Using("#####.##", tmprev#)
'      Else
'        ToPrint$ = ToPrint$ + "~ ~ ~ ~ ~ "
'      End If

    If TotalTax# > 0 Then
      ToPrint$ = ToPrint$ + "~Tax:" + "~" + Using("$###,###.##", TotalTax#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If Previous# <> 0 Then
      ToPrint$ = ToPrint$ + "~Previous:" + "~" + Using("$###,###.##", Previous#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    ToPrint$ = ToPrint$ + "~Current:" + "~" + Using("$###,###.##", UBBillRec(1).Transamt)

    If FinalFlag And CDeposit# Then
      ToPrint$ = ToPrint$ + "~Deposit:" + "~" + Using("$###,###.##", -UBCustRec(1).DepositAmt)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If Totalamt# < 0 And FinalFlag Then
      ToPrint$ = ToPrint$ + "~REFUND DUE" + "~" + Using("$#,###,###.##", Abs(Totalamt#))
    ElseIf Totalamt# < 0 And Not FinalFlag Then
      ToPrint$ = ToPrint$ + "~CREDIT BAL" + "~" + Using("$#,###,###.##", Totalamt#)
    Else
      ToPrint$ = ToPrint$ + "~TOTAL DUE" + "~" + Using("$#,###,###.##", Totalamt#)
    End If
    ToPrint$ = ToPrint$ + "~" + Num2Date$(UBBillRec(1).BillDate) + "~" + Num2Date$(UBBillRec(1).PastDueDate)
    ToPrint$ = ToPrint$ + "~" + Msg1$
    ToPrint$ = ToPrint$ + "~" + Msg2$
    ToPrint$ = ToPrint$ + "~" + Msg3$
    ToPrint$ = ToPrint$ + "~" + Msg4$
    ToPrint$ = ToPrint$ + "~" + Message$
    If DidADraftFlag Then
      ToPrint$ = ToPrint$ + "~" + "DRAFT NOTICE DO NOT PAY!! Draft Date-" + Num2Date$(UBBillRec(1).DraftDate)
    Else
      ToPrint$ = ToPrint$ + "~ "
    End If

    ToPrint$ = ToPrint$ + "~" + Acct$ 'Using("##########", UBBillRec(1).CustAcctNo)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ServAddr)

'This swapping of address lines and names is for proper mailing address
'printing for the bar coded discounts.
    If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).CustName)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ADDR2)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ADDR1)
     ' TmpAdd$ = QPTrim$(UBCustRec(1).ADDR2)
    Else
      ToPrint$ = ToPrint$ + "~ ~" + QPTrim$(UBCustRec(1).CustName) + "~ " + QPTrim$(UBCustRec(1).ADDR1)
     ' TmpAdd$ = QPTrim$(UBCustRec(1).ADDR1)
    End If
    ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).CITY) + " " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
'    If FinalFlag Then
'      ToPrint$ = ToPrint$ + "~" + Using("#######.##", Round#(Totalamt#))
'    Else
    If Not Totalamt# > 0 Then
      ToPrint$ = ToPrint$ + "~ "
    Else
      ToPrint$ = ToPrint$ + "~" + Using("$#,###,###.##", Round#(Totalamt# + CustPenalty#))
    End If
'    TmpAdd$ = Val(TmpAdd$)
    If PostBar = True Then
      TmpAdd$ = QPTrim$(UBCustRec(1).DPCode)
      BZip$ = Zip$ + TmpAdd$
    Else
      BZip$ = Zip$
    End If
    ToPrint$ = ToPrint$ + "~" + Zip$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    ToPrint$ = ToPrint$ + "~" + BZip$
    Print #UBRpt, ToPrint$
    ToPrint$ = ""
End Sub
'This bill uses billdate, duedate and pastduedate
Private Sub PrnLaserLetterPrePrinted(UBCustRec() As NewUBCustRecType, UBBillRec() As UBTransRecType, UBSetUpRec() As UBSetupRecType) 'Bill 19
'Utility Bill LaserLetterFormat(Uses Preprinted Stock)BAR CODE PRINTABLE
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  'ZDigit$ = GetZipEDigit$(Zip$)
  Zip$ = Zip$ '+ ZDigit$
  FoundAMtr = False
  For mChk = 1 To 7
    If UBBillRec(1).MtrTypes(mChk) > 0 Then
      FoundAMtr = True
      Exit For
    End If
  Next
  Select Case UBCustRec(1).LocMeters(1).MTRMulti
  Case 10
    Zero$ = "0"
  Case 100
    Zero$ = "00"
  Case 1000
    Zero$ = "000"
  Case Else
    Zero$ = ""
  End Select
  If UseEDateFlag = False Then
    If FoundAMtr = False Then
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    End If
  End If
  AcctNum = UBBillRec(1).CustAcctNo
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
'  If Previous# <> 0 Then
'  End If
'  If TotalTax > 0 Then
'    lpcnt = lpcnt - 1
'  End If

  If FinalFlag And CDeposit# Then
    Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
  End If
'Need to get Penalty calculated based on defaults from setup screen
  If Totalamt# > 0 And Not FinalFlag Then
    If UBCustRec(1).LATEFEE = "Y" Then
      If TotalTax > 0 Then
        CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
      Else
        CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
      End If
    Else
      CustPenalty# = 0
    End If
  Else
    CustPenalty# = 0
  End If
'-=-=-=-=-=-=-=-=-=-=-=-
  AcctNum = UBBillRec(1).CustAcctNo
  Acct$ = QPTrim$(Str$(AcctNum))
  Select Case AcctNum
  Case Is < 10
    Acct$ = "00" + Acct$
  Case Is < 100
    Acct$ = "0" + Acct$
  End Select
  ToPrint$ = Using("########", (FBillNO& + PrintedCnt))
  ToPrint$ = ToPrint$ + "~" + PrevDate$ + "~" + DateRead$
   'Only Print Days if Greater than 0
   If DaysINRead > 0 Then
     ToPrint$ = ToPrint$ + "~" + Using("####", DaysINRead)
   Else
     ToPrint$ = ToPrint$ + "~ "
   End If
  PCnt = 0
  For PCnt = 1 To 7
        If UBBillRec(1).CurRead(PCnt) > 0 Then
          UsageAmt& = UBBillRec(1).CurRead(PCnt) - UBBillRec(1).PrevRead(PCnt)
          If UsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(PCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(PCnt)) + UBBillRec(1).CurRead(PCnt)
          End If
          UsageAmt& = Round(UsageAmt& * UBCustRec(1).LocMeters(PCnt).MTRMulti)
          If MtrNFlag = 2 Then
            ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).LocMeters(PCnt).MtrIDNO)
          Else
            ToPrint$ = ToPrint$ + "~" + QPTrim(UBCustRec(1).LocMeters(PCnt).MtrNum)
          End If
          ToPrint$ = ToPrint$ + "~" + Using("#########", UBBillRec(1).PrevRead(PCnt))
          ToPrint$ = ToPrint$ + "~" + Using("#########", UBBillRec(1).CurRead(PCnt))
          ToPrint$ = ToPrint$ + "~" + Using("#######", UsageAmt&)
        Else
          ToPrint$ = ToPrint$ + "~ ~ ~ ~ "
        End If
  Next
    For WRevCnt = 1 To 14
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBSetUpRec(1).Revenues(WRevCnt).RevName)
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).serv(WRevCnt).Ratecode)
        ToPrint$ = ToPrint$ + "~" + Using("#####.##", UBBillRec(1).RevAmt(WRevCnt))
      Else
        ToPrint$ = ToPrint$ + "~ ~ ~ "
      End If
    Next
    ToPrint$ = ToPrint$ + "~ ~ ~ "
'    tmprev# = 0
'      For PCnt = lpcnt To 15
'        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
'          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
'        End If
'      Next
'      If tmprev# <> 0 Then
'        ToPrint$ = ToPrint$ + "~ ~ ~ ~" + "   Other:"
'        ToPrint$ = ToPrint$ + "~" + Using("#####.##", tmprev#)
'      Else
'        ToPrint$ = ToPrint$ + "~ ~ ~ ~ ~ "
'      End If

'    If TotalTax# > 0 Then
'      ToPrint$ = ToPrint$ + "~Tax:" + "~" + Using("$###,###.##", TotalTax#)
'    Else
'      ToPrint$ = ToPrint$ + "~ ~ "
'    End If
    If Previous# < 0 Then
      ToPrint$ = ToPrint$ + "~Prev Cred:" + "~" + Using("$###,###.##", Previous#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If Previous# > 0 Then
      ToPrint$ = ToPrint$ + "~Previous:" + "~" + Using("$###,###.##", Previous#)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    ToPrint$ = ToPrint$ + "~Current:" + "~" + Using("$###,###.##", UBBillRec(1).Transamt)
    
    If FinalFlag And CDeposit# Then
      ToPrint$ = ToPrint$ + "~Deposit:" + "~" + Using("$###,###.##", -UBCustRec(1).DepositAmt)
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    If Totalamt# < 0 And FinalFlag Then
      ToPrint$ = ToPrint$ + "~REFUND DUE" + "~" + Using("$#,###,###.##", Abs(Totalamt#))
    ElseIf Totalamt# < 0 And Not FinalFlag Then
      ToPrint$ = ToPrint$ + "~CREDIT BAL" + "~" + Using("$#,###,###.##", Totalamt#)
    Else
      ToPrint$ = ToPrint$ + "~TOTAL DUE" + "~" + Using("$#,###,###.##", Totalamt#)
    End If
    ToPrint$ = ToPrint$ + "~" + Num2Date$(UBBillRec(1).BillDate) + "~" + Num2Date$(UBBillRec(1).PastDueDate)
    ToPrint$ = ToPrint$ + "~" + Msg1$
    ToPrint$ = ToPrint$ + "~" + Msg2$
    ToPrint$ = ToPrint$ + "~" + Msg3$
    ToPrint$ = ToPrint$ + "~" + Msg4$
    ToPrint$ = ToPrint$ + "~" + Message$
    If DidADraftFlag Then
      ToPrint$ = ToPrint$ + "~" + "DRAFT NOTICE DO NOT PAY!! Draft Date-" + Num2Date$(UBBillRec(1).DraftDate)
    Else
      ToPrint$ = ToPrint$ + "~ "
    End If
   
    ToPrint$ = ToPrint$ + "~" + Acct$ 'Using("##########", UBBillRec(1).CustAcctNo)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ServAddr)
    
'This swapping of address lines and names is for proper mailing address
'printing for the bar coded discounts.
    If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).CustName)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ADDR2)
      ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).ADDR1)
     ' TmpAdd$ = QPTrim$(UBCustRec(1).ADDR2)
    Else
      ToPrint$ = ToPrint$ + "~ ~" + QPTrim$(UBCustRec(1).CustName) + "~ " + QPTrim$(UBCustRec(1).ADDR1)
     ' TmpAdd$ = QPTrim$(UBCustRec(1).ADDR1)
    End If
    ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).CITY) + " " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
'    If FinalFlag Then
'      ToPrint$ = ToPrint$ + "~" + Using("#######.##", Round#(Totalamt#))
'    Else
    If Not Totalamt# > 0 Then
      ToPrint$ = ToPrint$ + "~ "
    Else
      ToPrint$ = ToPrint$ + "~" + Using("$#,###,###.##", Round#(Totalamt# + CustPenalty#))
    End If
'    TmpAdd$ = Val(TmpAdd$)
    If PostBar = True Then
      TmpAdd$ = QPTrim$(UBCustRec(1).DPCode)
      BZip$ = Zip$ + TmpAdd$
    Else
      BZip$ = Zip$
    End If
    ToPrint$ = ToPrint$ + "~" + Zip$ + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    ToPrint$ = ToPrint$ + "~" + BZip$ + "~" + DueDate$
    Print #UBRpt, ToPrint$
    ToPrint$ = ""

End Sub
Private Sub PrnBill3ForElkton(UBCustRec() As NewUBCustRecType, UBBillRec() As UBTransRecType, UBSetUpRec() As UBSetupRecType)  'Bill 16
'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL
    Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    lpcnt = 8
    If Not LPIFlag Then
      LPIFlag = -2
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    End If
    EFoundMtr = False
    DFoundMtr = False
    WFoundMtr = False
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        Select Case UBBillRec(1).MtrTypes(mChk)
        Case 1, 2, 3, 7
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          End If
          WFoundMtr = True
        Case 4
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          If EUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
          End If
          EFoundMtr = True
        Case 5
          DCurrRead& = UBBillRec(1).CurRead(mChk)
          DPrevRead& = UBBillRec(1).PrevRead(mChk)
          DUsageAmt& = DCurrRead& - DPrevRead&
          If DUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(DPrevRead&)) - 1)
            DUsageAmt& = (MaxMeterAmt& - DPrevRead&) + DCurrRead&
          End If
          DFoundMtr = True
        End Select
      End If
    Next
    FoundAMtr = False
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        FoundAMtr = True
        Exit For
      End If
    Next
'-=-=-=-=-=-=-=-=-=-=-=-
    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    If UseEDateFlag = False Then
      If FoundAMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    AcctNum = UBBillRec(1).CustAcctNo

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If

    If FinalFlag And CDeposit# Then
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
    TotalTax# = 0
    For TaxCnt = 1 To 15  'MaxRevsCnt
      TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
    Next
    For TaxCnt = 9 To 14
       TotalTax# = Round(TotalTax# + UBBillRec(1).RevAmt(TaxCnt))
    Next
'Need to get Penalty calculated based on defaults from setup screen
    If Totalamt# > 0 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
'        If TotalTax > 0 Then
'          CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
'        Else
          CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
'        End If
      Else
        CustPenalty# = 0
      End If
    Else
      CustPenalty# = 0
    End If

   
'***************************************
    Print #UBRpt, "~"; Tab(50); Using("########", (FBillNO& + PrintedCnt))
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Tab(18); PrevDate$; Tab(33); DateRead$;
     'Only Print Days if Greater than 0
     If DaysINRead > 0 Then
       Print #UBRpt, "      "; Using("####", DaysINRead)
     Else
       Print #UBRpt, " "
     End If

    Print #UBRpt, " "
    Print #UBRpt, " "

    PCnt = 0
    For WRevCnt = 1 To 7
      PCnt = PCnt + 1
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
        Select Case PCnt
        Case 1, 2
          If WFoundMtr Then
            Print #UBRpt, Tab(4); Using("##########", WPrevRead&);
            Print #UBRpt, Tab(14); Using("##########", WCurrRead&);
            Print #UBRpt, Tab(24); Using("#######", WUsageAmt&);
          End If
        Case 3  'electric
          If EFoundMtr Then
            Print #UBRpt, Tab(4); Using("##########", EPrevRead&);
            Print #UBRpt, Tab(14); Using("##########", ECurrRead&);
            Print #UBRpt, Tab(24); Using("#######", EUsageAmt&);
          End If
        Case 4
          If DFoundMtr Then
            Print #UBRpt, Tab(4); Using("##########", DPrevRead&);
            Print #UBRpt, Tab(14); Using("##########", DCurrRead&);
            Print #UBRpt, Tab(24); Using("#######", DUsageAmt&);
          End If
        End Select
        Print #UBRpt, Tab(33); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));
      End If
      Select Case PCnt
      Case 1
        Print #UBRpt, Tab(45); Using("##########", UBBillRec(1).CustAcctNo)
      Case 5
        Print #UBRpt, Tab(45); Left$(UBCustRec(1).ServAddr, 26)
      Case 2, 3, 4     'was else
        Print #UBRpt,
      End Select
    Next
    tmprev# = 0
      For PCnt = 6 To 8
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, "Other";
        Print #UBRpt, Tab(33); Using("#####.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If

    If TotalTax# > 0 Then
      Print #UBRpt, Tab(14); "     TAX:"; Tab(31); Using("$###,###.##", TotalTax#)
    Else
      Print #UBRpt, ""
    End If
    If Previous# <> 0 Then
      Print #UBRpt, Tab(14); "Previous:"; Tab(31); Using("$###,###.##", Previous#)
    Else
     Print #UBRpt,
    End If
    Print #UBRpt, Tab(14); " Current:"; Tab(31); Using("$###,###.##", UBBillRec(1).Transamt)

    If FinalFlag And CDeposit# Then
      Print #UBRpt, Tab(14); " Deposit:"; Tab(31); Using("$###,###.##", -UBCustRec(1).DepositAmt);
    End If
    Print #UBRpt, Tab(45); Num2Date$(UBBillRec(1).BillDate); Tab(60); Num2Date$(UBBillRec(1).PastDueDate)
    Print #UBRpt, Tab(2); Msg1$
    Print #UBRpt, Tab(2); Msg2$
    Print #UBRpt, Tab(2); Msg3$;
    If Totalamt# < 0 And FinalFlag Then
      Print #UBRpt, Tab(35); "Refund:"; Tab(44); Using("$###,###.##", Abs(Totalamt#))
    Else
      Print #UBRpt, Tab(36); "Total:"; Tab(44); Using("$###,###.##", Totalamt#)
    End If

    If DidADraftFlag Then
      Print #UBRpt, Tab(2); "DRAFT NOTICE DO NOT PAY!!" ';
    ElseIf Len(Message$) > 0 Then
      Print #UBRpt, Tab(2); Message$ ';
    Else
      Print #UBRpt, Tab(2); Msg4$ ';
    End If

    Print #UBRpt, " "
')))))))))))))))))))))))
    If AcctBar = True Then
'*************For Okidata to print Bar code
      Print #UBRpt, Tab(55); Chr$(27); Chr$(16); "A"; 'String$(50, " ")
      Print #UBRpt, Chr$(8);
      Print #UBRpt, Chr$(2); "0";
      Print #UBRpt, "0"; Chr$(2);
      Print #UBRpt, Chr$(1); Chr$(1);
      Print #UBRpt, Chr$(1); Chr$(2);
      Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
'**************************
    Else
      Print #UBRpt, " "
    End If
'))))))))))))))))))))))))))
    Print #UBRpt, Tab(22); Left$(UBCustRec(1).CustName, 29)
    Print #UBRpt, Tab(22); QPTrim(UBCustRec(1).ADDR1)
    Print #UBRpt, Using("##########", UBBillRec(1).CustAcctNo); Tab(22); QPTrim(UBCustRec(1).ADDR2);
    Print #UBRpt, Tab(55); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(22); Left$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
    Print #UBRpt, " "
    Print #UBRpt, Using("$###,###.##", Totalamt#)
    'Print #UBRpt,
    Print #UBRpt, " "
    Print #UBRpt, " "
    If Totalamt# > 0 Then
'      Print #UBRpt, Using("#######.##", Round#(Totalamt#));
'    Else
      Print #UBRpt, Using("$###,###.##", Round#(Totalamt# + CustPenalty#));
    End If
    If PostBar = True Then
      Print #UBRpt, Tab(22); Chr$(27); Chr$(16); "C"; Chr$(Len(Zip$)); Zip$
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, " "
    Print #UBRpt, "~"
End Sub
Private Sub PrnBill4ForBucksp(UBCustRec() As NewUBCustRecType, UBBillRec() As UBTransRecType, UBSetUpRec() As UBSetupRecType)
'This bill prints only 23 lines so there is room at bottom for PO to print their barcode
    If Not LPIFlag Then
      LPIFlag = -2
      Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    End If
    lpcnt = 11
    FoundAMtr = False
    WFoundMtr = False
    EFoundMtr = False
'this is same for all since meter types all same
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        Select Case UBBillRec(1).MtrTypes(mChk)
        Case 1, 2, 3, 7, 9
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
           End If
           'Added this to show zeros for mult on consumption amts
          WUsageAmt& = Round(WUsageAmt& * UBCustRec(1).LocMeters(mChk).MTRMulti)
          WFoundMtr = True
        Case 4, 5
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          If EUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
            End If
            'Added this to show zeros for mult on consumption amts
          EUsageAmt& = Round(EUsageAmt& * UBCustRec(1).LocMeters(mChk).MTRMulti)
          EFoundMtr = True
        End Select
      End If
    Next
    If WFoundMtr = False And EFoundMtr = False Then
      FoundAMtr = False
    Else
      FoundAMtr = True
    End If
    If UseEDateFlag = False Then
      If FoundAMtr = False Then
        'if no metered services then adjust read dates to billdate
        'and billdate - 30
        DateRead$ = Num2Date$(UBBillRec(1).BillDate)
        PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      End If
    End If
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If Previous# <> 0 Then
      lpcnt = lpcnt - 1
    End If
    If TotalTax > 0 Then
      lpcnt = lpcnt - 1
    End If
    If FinalFlag And CDeposit# Then
      lpcnt = lpcnt - 1
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If
    If Totalamt# > 0 And Not FinalFlag Then
      If UBCustRec(1).LATEFEE = "Y" Then
          If TotalTax > 0 Then
            CalcPenamt Round(UBBillRec(1).Transamt - TotalTax), Previous, Round(Totalamt - TotalTax)
          Else
            CalcPenamt UBBillRec(1).Transamt, Previous, Totalamt
          End If
      Else
        CustPenalty# = 0
      End If
    Else
      CustPenalty# = 0
    End If
    AcctNum = UBBillRec(1).CustAcctNo

    Acct$ = QPTrim$(Str$(AcctNum))
    Select Case AcctNum
    Case Is < 10
      Acct$ = "00" + Acct$
    Case Is < 100
      Acct$ = "0" + Acct$
    End Select
    AcctLen = Len(Acct$)

    Print #UBRpt, "~"; Tab(50); Using("########", FBillNO& + PrintedCnt)
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, " "
    Print #UBRpt, Using("##########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(15); Left$(UBCustRec(1).ServAddr, 31); Tab(50); Using("########", UBBillRec(1).CustAcctNo);
    Print #UBRpt, Tab(62); Num2Date$(UBBillRec(1).BillDate)
    Print #UBRpt, " "
    If Not BFlag Then
      Print #UBRpt, Tab(50); DueDate$; Tab(64); Using("#####.##", Totalamt#)
      Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate); Tab(15); PrevDate$; Tab(26); DateRead$;
     'Only Print Days if Greater than 0
      If DaysINRead > 0 Then
        Print #UBRpt, Tab(39); Using("####", DaysINRead)
      Else
        Print #UBRpt, " "
      End If
    Else
      Print #UBRpt, Tab(50); DueDate$; Tab(64); Using("#####.##", Totalamt#)
      Print #UBRpt, Tab(3); Num2Date$(UBBillRec(1).BillDate);
      Print #UBRpt, " "
    End If
    Print #UBRpt, Tab(50); Num2Date$(UBBillRec(1).PastDueDate);
    If Not Totalamt# > 0 Then
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt#))
    Else
      Print #UBRpt, Tab(64); Using("#####.##", Round#(Totalamt# + CustPenalty#))
    End If
      Print #UBRpt,
    PCnt = 0
    For WRevCnt = 1 To lpcnt - 1
      PCnt = PCnt + 1  'Printable lines
      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
        Print #UBRpt, " "; Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);

        Select Case PCnt
        Case 1, 2  'water/sewer
          If WFoundMtr Then
            If Not BFlag Then
              Print #UBRpt, Tab(7); Using("##########", WCurrRead&);
              Print #UBRpt, Tab(17); Using("##########", WPrevRead&);
              Print #UBRpt, Tab(28); Using("#######", WUsageAmt&);
            End If
         End If
        End Select
'       End If
       Print #UBRpt, Tab(36); Using("#####.##", UBBillRec(1).RevAmt(WRevCnt));

      End If
      Select Case PCnt
      Case 2
        'print bar code acct
        ')))))))))))))))))))))))
    If AcctBar = True Then
'*************For Okidata to print Bar code
      Print #UBRpt, Tab(47); Chr$(27); Chr$(16); "A";
      Print #UBRpt, Chr$(8);
      Print #UBRpt, Chr$(2); "0";
      Print #UBRpt, "0"; Chr$(2);
      Print #UBRpt, Chr$(1); Chr$(1);
      Print #UBRpt, Chr$(1); Chr$(2);
      Print #UBRpt, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); Acct$
'**************************
    Else
      Print #UBRpt, " "
    End If
'))))))))))))))))))))))))))

      Case 5
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CustName, 29)
      Case 6
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR1)
      Case 7
        Print #UBRpt, Tab(47); QPTrim$(UBCustRec(1).ADDR2)
      Case 8
        Print #UBRpt, Tab(47); Left$(UBCustRec(1).CITY, 14); " "; QPTrim$(UBCustRec(1).STATE); " "; Left$(UBCustRec(1).ZIPCODE, 5)
      Case Else
        Print #UBRpt, " "
      End Select
    Next
    tmprev# = 0
      For PCnt = lpcnt To 15
        If UBBillRec(1).RevAmt(PCnt) <> 0 Then
          tmprev# = tmprev# + UBBillRec(1).RevAmt(PCnt)
        End If
      Next
      If tmprev# <> 0 Then
        Print #UBRpt, "                     Other:  ";
        Print #UBRpt, Tab(36); Using("#####.##", tmprev#)
      Else
        Print #UBRpt, " "
      End If
    If TotalTax# > 0 Then
      Print #UBRpt, "                       TAX:  "; Tab(34); Using("$###,###.##", TotalTax#)
'    Else
'      Print #UBRpt, " "
    End If
     If Previous# <> 0 Then
      Print #UBRpt, "                  Previous:  "; Tab(34); Using("$###,###.##", Previous#)
'    Else
'      Print #UBRpt, " "
    End If
    Print #UBRpt, "                   Current:  "; Tab(34); Using("$###,###.##", UBBillRec(1).Transamt)
    Print #UBRpt, Tab(32); "--------------"
'line 24
    If FinalFlag And CDeposit# Then
      Print #UBRpt, "                   Deposit:  "; Tab(34); Using("$###,###.##", -UBCustRec(1).DepositAmt)
'    Else
'      Print #UBRpt, " "
    End If

    If Totalamt# < 0 And FinalFlag Then
      Print #UBRpt, "                Refund Due:  "; Tab(34); Using("$###,###.##", Abs(Totalamt#))
    Else
      Print #UBRpt, "                     Total:  "; Tab(34); Using("$###,###.##", Totalamt#)
    End If
    Print #UBRpt, Tab(3); Msg1$; Tab(47); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    Print #UBRpt, Tab(3); Msg2$;
    
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT NOTICE DO NOT PAY!!"
    Else
      Print #UBRpt, " "
    End If
    Print #UBRpt, Tab(3); Msg3$;
   
    If DidADraftFlag Then
      Print #UBRpt, Tab(47); "DRAFT DATE: "; Num2Date$(UBBillRec(1).DraftDate)
    Else
      Print #UBRpt, " "
    End If
      If Len(Message$) > 0 Then
        Print #UBRpt, Tab(3); Message$;
      Else
        Print #UBRpt, Tab(3); Msg4$;
      End If
      If Rteflag Then
        Print #UBRpt, Tab(47); "Route: "; UBCustRec(1).POSTRTE
      Else
        Print #UBRpt, " "
      End If
      Print #UBRpt, " "
      Print #UBRpt, " "
    Print #UBRpt, "~"

End Sub
Private Sub CreateSCSFileTransfer(UBCustRec() As NewUBCustRecType, UBBillRec() As UBTransRecType, UBSetUpRec() As UBSetupRecType)
    WFoundMtr = False
    IFoundMtr = False
    CubicMtr = False

    WCurrRead& = 0
    WPrevRead& = 0
    WUsageAmt& = 0
    ICurrRead& = 0
    IPrevRead& = 0
    IUsageAmt& = 0
    MtrNumb$ = ""
    MtrType$ = ""
    For mChk = 1 To 7
      If UBBillRec(1).MtrTypes(mChk) > 0 Then
        MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MtrNum)
        If Len(MtrNumb$) = 0 Then
          MtrNumb$ = "???"
        End If
        MtrTyp = UBBillRec(1).MtrTypes(mChk)
        Select Case MtrTyp
        Case 1
          IFoundMtr = True
          ICurrRead& = UBBillRec(1).CurRead(mChk)
          IPrevRead& = UBBillRec(1).PrevRead(mChk)
          IUsageAmt& = ICurrRead& - IPrevRead&
          If IUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(IPrevRead&)) - 1)
            IUsageAmt& = (MaxMeterAmt& - IPrevRead&) + ICurrRead&
          End If
          If UBCustRec(1).LocMeters(mChk).MtrUnit = "C" Then
            CubicMtr = True
            IUsageAmt& = IUsageAmt& * 7.481
          End If
          If MtrTyp = 1 Then
            MtrType$ = "W"
          ElseIf MtrTyp = 2 Then
            MtrType$ = "S"
          ElseIf MtrTyp = 3 Then
            MtrType$ = "C"
          ElseIf MtrTyp = 7 Then
            MtrType$ = "T"
          End If
        Case 2, 3, 7
          WFoundMtr = True
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          If WUsageAmt& < 0 Then
            MaxMeterAmt& = 10& ^ (Len(Str$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          End If
          If UBCustRec(1).LocMeters(mChk).MtrUnit = "C" Then
            CubicMtr = True
            WUsageAmt& = WUsageAmt& * 7.481
          End If
          If MtrTyp = 1 Then
            MtrType$ = "W"
          ElseIf MtrTyp = 2 Then
            MtrType$ = "S"
          ElseIf MtrTyp = 3 Then
            MtrType$ = "C"
          ElseIf MtrTyp = 7 Then
            MtrType$ = "T"
          End If
        End Select
      End If
    Next

    If WFoundMtr = False And IFoundMtr = False Then
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
      DaysINRead = 30
    End If

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    Totalamt# = Round#(Previous# + UBBillRec(1).Transamt)
    If FinalFlag And CDeposit# Then
      lpcnt = lpcnt - 1
      Totalamt# = Round#(Totalamt# - UBCustRec(1).DepositAmt)
    End If

'          If Totalamt# > 0 Then
'            'TenPct# = 0
'            If DaysINRead < 1 Then DaysINRead = 1
'            AvgCst# = Round#(Totalamt# / DaysINRead)
'          Else
'            'TenPct# = 0
'            AvgCst# = 0
'          End If

    ReDim PrintRec(1) As BillOutRecType

    PrintRec(1).AcctNo = Using("########", Str$(CustAcctNo&))
    PrintRec(1).LocationNum = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    RSet PrintRec(1).CustName = QPTrim$(UBCustRec(1).CustName)
    RSet PrintRec(1).ADDR1 = QPTrim$(UBCustRec(1).ADDR1)
    RSet PrintRec(1).ADDR2 = QPTrim$(UBCustRec(1).ADDR2)
    RSet PrintRec(1).ServAddr = QPTrim$(UBCustRec(1).ServAddr)
    RSet PrintRec(1).CITY = QPTrim$(UBCustRec(1).CITY)
    RSet PrintRec(1).STATE = QPTrim$(UBCustRec(1).STATE)
    RSet PrintRec(1).ZIPCODE = QPTrim$(UBCustRec(1).ZIPCODE)
    If Fflag Then
      PrintRec(1).BillType = "F"
      PrintRec(1).DepAppAmt = Using(Fmt10a$, Str$(CDeposit#))
    Else
      PrintRec(1).BillType = "N"
      PrintRec(1).DepAppAmt = ""
    End If

    PrintRec(1).PrevDue = Using(Fmt15$, Str$(Previous#))
    PrintRec(1).CurrDue = Using(Fmt15$, Str$(UBBillRec(1).Transamt))
    PrintRec(1).TotalDue = Using(Fmt15$, Str$(Totalamt#))
    PrintRec(1).CurrDate = DateRead$
    PrintRec(1).PrevDate = PrevDate$

    If WFoundMtr Then
      PrintRec(1).CurrRead = Using(Fmt10$, Str$(WCurrRead&))
      PrintRec(1).PrevRead = Using(Fmt10$, Str$(WPrevRead&))
      PrintRec(1).Consump = Using(Fmt10$, Str$(WUsageAmt&))
      PrintRec(1).ServDays = Using("####", Str$(DaysINRead))
    Else
      PrintRec(1).CurrRead = ""
      PrintRec(1).PrevRead = ""
      PrintRec(1).Consump = ""
      PrintRec(1).ServDays = ""
    End If
    If IFoundMtr = False Then
      PrintRec(1).ICurrRead = ""
      PrintRec(1).IPrevRead = ""
      PrintRec(1).IConsump = ""
      PrintRec(1).IServDays = ""
    Else
      PrintRec(1).ICurrRead = Using(Fmt10$, Str$(ICurrRead&))
      PrintRec(1).IPrevRead = Using(Fmt10$, Str$(IPrevRead&))
      PrintRec(1).IConsump = Using(Fmt10$, Str$(IUsageAmt&))
      PrintRec(1).IServDays = Using("####", Str$(DaysINRead))
    End If

    For serv = 1 To 15
      PrintRec(1).ServInfo(serv).ServText = QPTrim$(UBSetUpRec(1).Revenues(serv).RevName)
      PrintRec(1).ServInfo(serv).ServAmt = Using(Fmt10a$, Str$(UBBillRec(1).RevAmt(serv)))
    Next

    PrintRec(1).MtrType = MtrType$
    If CubicMtr Then
      PrintRec(1).MtrUnit = "C"
    Else
      PrintRec(1).MtrUnit = "G"
    End If

    PrintRec(1).BillDate = BillDateS$
    PrintRec(1).PastDueDate = PastDueDate$
    If DidADraftFlag Then
      PrintRec(1).DraftDate = DraftDateS$
    Else
      PrintRec(1).DraftDate = ""
    End If
    If Len(Message$) > 0 Then
      RSet PrintRec(1).MsgLine1 = Message$
    Else
      RSet PrintRec(1).MsgLine1 = Msg1$
    End If
    RSet PrintRec(1).MsgLine2 = Msg2$
    RSet PrintRec(1).MsgLine3 = Msg3$
    RSet PrintRec(1).MsgLine4 = Msg4$
    LSet PrintRec(1).MtrNumb = MtrNumb$

    PrintRec(1).CrLf = CrLf$
    Put #UBRpt, , PrintRec(1)
End Sub
Private Sub GetPenaltyVals(UBBillSetup() As UBBillSetupType)
    MinBalance# = UBBillSetup(1).MinBalance
    If MinBalance# < 0 Then
      MinBalance# = 0
    End If
    If QPTrim$(UBBillSetup(1).ChargeOn) = "Current Balance" Then  'Applying to Current
      UsePrevFlag = False
      UseCurrFlag = True
    ElseIf QPTrim$(UBBillSetup(1).ChargeOn) = "Previous Balance" Then 'Applying to Previous
      UsePrevFlag = True
      UseCurrFlag = False
    Else   'Apply to total
      UsePrevFlag = True
      UseCurrFlag = True
    End If
    'Get percent or fixed amount
    PctAmt# = UBBillSetup(1).PctCharge
    FixAmt# = UBBillSetup(1).AmtCharge
    PenAmt2# = UBBillSetup(1).AmtChge2
    If PenAmt2# > 0 And Not Fflag Then
      Use2ndPen = True
    Else
      PenAmt2# = 0
      Use2ndPen = False
    End If
    If Len(QPTrim(UBBillSetup(1).GreatLess)) <> 0 Then
      If QPTrim$(UBBillSetup(1).GreatLess) = "G" Then
        GreaterFlag = True
      Else
        GreaterFlag = False
      End If
      PctAmt# = PctAmt# / 100
      UseBothFlag = True
    Else
      If PctAmt# > 0 Then
        PctAmt# = PctAmt# / 100
        FixAmt# = 0
        UsePctFlag = True
      Else
        PctAmt# = 0
        UsePctFlag = False
      End If
      UseBothFlag = False
    End If

End Sub
Private Sub CalcPenamt(Curr As Double, Prev As Double, Tot As Double)
  Dim PenBal As Double, CustPctPenalty As Double, CustFixPenalty As Double
  If Tot > 0 And Not BFlag Then  'if they have any balance but not B status
    If Use2ndPen And Prev > 0 Then
      CustPenalty# = PenAmt2#
      GoTo SkipEm
    End If

   ' If Curr >= MinBalance# Or Prev > MinBalance# Then
      If UseBothFlag Then             'both an amount and percent
        If UsePrevFlag And Not UseCurrFlag Then       'use prev not curr
          If Curr < 0 Then
            PenBal# = Tot
          Else
            PenBal# = Prev
          End If
        ElseIf UseCurrFlag And Not UsePrevFlag Then   'use curr not
          If Prev < 0 Then
            PenBal# = Tot
          Else
            PenBal# = Curr
          End If
        ElseIf UsePrevFlag And UseCurrFlag Then       'use curr and
          PenBal# = Tot
        End If
        CustPctPenalty# = Round#(PenBal# * PctAmt#)
        CustFixPenalty# = FixAmt#
        If PenBal# <= MinBalance# Then    'if cust had p
          CustPenalty# = 0
          GoTo SkipEm
        End If
        If GreaterFlag Then
          If CustPctPenalty# >= CustFixPenalty# Then
            CustPenalty# = CustPctPenalty#
          Else
            CustPenalty# = CustFixPenalty#
          End If
        Else          'nope want whichever is less
          If CustPctPenalty# >= CustFixPenalty# Then
            CustPenalty# = CustFixPenalty#
          Else
            CustPenalty# = CustPctPenalty#
          End If
        End If
      ElseIf UsePctFlag Then          'if they want a percent penalty
        If UsePrevFlag And Not UseCurrFlag Then       'using prev not curr
               '030398 Modified to consider a credit in cur or prev balances
          If Curr < 0 Then
            PenBal# = Tot
          Else
            PenBal# = Prev
          End If
          If PenBal# <= MinBalance# Then    'if cust had prev bal
            CustPenalty# = 0
            GoTo SkipEm
          End If
          CustPenalty# = Round#(PenBal# * PctAmt#)
        ElseIf UseCurrFlag And Not UsePrevFlag Then   'using curr not prev
               '030398 Modified to consider a credit in cur or prev balances
          If Not TennRdg Then
            If Prev < 0 Then
              PenBal# = Tot
            Else
              PenBal# = Curr
            End If
          Else
            PenBal# = Curr
          End If
''                    'code added to exclude tax
''                    '092898 Said they didn't take partial payments - Not!
''                    If TennFlag Then            'AND UBCustRec(1).TaxExpt <> "Y" then
''                      GoSub GetTennRidgeLastBill
''                    End If
''                    If CashFlag Then
''                      GoSub GetCashionLastBill
''                    End If

          If PenBal# <= MinBalance# Then
            CustPenalty# = 0
            GoTo SkipEm
          End If
          CustPenalty# = Round#(PenBal# * PctAmt#)
        ElseIf UsePrevFlag And UseCurrFlag Then       'use curr and prev
          PenBal# = Tot
''                    If SunSetFlag Then
''                      GoSub CheckSunSet
''                      'This adjusts PenBal# for sunsets calc
''                    End If
''                    If TuckFlag Then
''                      GoSub CheckTucka
''                      'This adjusts PenBal# for TUCKASEIGEE calc
''                    End If
          If PenBal# <= MinBalance# Then
            CustPenalty# = 0
            GoTo SkipEm
          End If
          CustPenalty# = Round#(PenBal# * PctAmt#)
        End If
      Else            'Using a FIXED penalty amount
        If UsePrevFlag And Not UseCurrFlag Then
               '030398 Modified to consider a credit in cur or prev balances
          If Curr < 0 Then
            PenBal# = Tot
          Else
            PenBal# = Prev
          End If
        ElseIf UseCurrFlag And Not UsePrevFlag Then
               '030398 Modified to consider a credit in cur or prev balances
          If Not TennRdg Then
            If Prev < 0 Then
              PenBal# = Round(Curr + Prev)
            Else
              PenBal# = Curr
            End If
          Else
            PenBal# = Curr
          End If
        ElseIf UsePrevFlag And UseCurrFlag Then
               'do not need to check for prev >0 or curr>0 here!!
          PenBal# = Tot
        End If
        If PenBal# <= MinBalance# Then
          CustPenalty# = 0
          GoTo SkipEm
        End If
        CustPenalty# = FixAmt#
      End If
    'End If
   End If            'if balance >0
SkipEm:
  
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
'=======================================
Private Sub DOChurchTrans(UBBillRec() As UBTransRecType, UBCustRec() As NewUBCustRecType)
  'mod for cleveland***
  For MtrCnt = 1 To 7
    If UBBillRec(1).MtrTypes(MtrCnt) > 0 Then
      UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
      UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
      Exit For
    End If
  Next

  If UBBillRec(1).ReadDate <= 0 Then
    UBBillRec(1).ReadDate = BillDate - 30
  End If
  If UBBillRec(1).PrevDate <= 0 Then
    UBBillRec(1).PrevDate = UBBillRec(1).ReadDate - 30
  End If

  If UseEDateFlag Then
    UBBillRec(1).PrevDate = PRDate
    UBBillRec(1).ReadDate = CRDate
  End If

  UBBillRec(1).BillDate = BillDate
  'BillInfoRec(1).DueDate = DueDate2
  UBBillRec(1).PastDueDate = PastDate
  UBBillRec(1).DraftDate = DraftDate
  UBBillRec(1).BillMsg = Message$


  'UBBillRec(1)CustLocation = CustAcctNo&
  UBBillRec(1).CustAcctNo = CustAcctNo&

  Put UBBill, CustAcctNo&, UBBillRec(1)

End Sub

Private Sub DoBLMask()
  Dim UBRptA As Integer, Message As String
  Dim PCnt As Integer, AcctLen As Integer, MPCnt As Integer
  UBRptA = FreeFile
  MaskBill$ = UBPath$ + "UBBLA.RPT"
  Open MaskBill$ For Output As UBRptA
  Select Case BLType
  Case 8, 12, 14
    GoSub PrnStand21LineMask
  Case 3
    GoSub PrintNewStandBarMask
  Case 4, 15, 18
    GoSub PrintNewStandV1Mask
  Case 5
    GoSub PrintNewStandRmStampMask
  Case 6
    GoSub PrnStand24L2BxMask
  Case 7
    GoSub PrnStand24L3BxMask
  Case Else
    'NO MASK
  End Select
  Close UBRptA
Exit Sub
PrnStand21LineMask: '4
    Message$ = "XXXXXXXXXXXXXXXXXXXXXXXXX"
'
    Print #UBRptA, "~"; "#####"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(3); "XX"; Tab(8); "XX"; Tab(13); "XX";
    Print #UBRptA, Tab(17); "XX"; Tab(22); "XX"; Tab(27); "XX";
    If BLType = 12 Then
      Print #UBRptA, Tab(32); "Loc: XX-XXXXXX"
    Else
      Print #UBRptA, Tab(40); "XX/XX/XXXX"
    End If
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(2); "#########";
    Print #UBRptA, Tab(12); "#########";
    Print #UBRptA, Tab(22); "########";
    Print #UBRptA, Tab(35); "XXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXX  XX XXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); String$(24, "-")
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##";
    Print #UBRptA, Tab(34); Message$
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##"
    Print #UBRptA, Tab(3); "XXXXXXXX"; Tab(20); "#####.##"
    Print #UBRptA, " "
    Print #UBRptA, Tab(5); "XXXXX"; Tab(19); "######.##";
    If BLType = 14 Then
      Print #UBRptA, Tab(35); "XXXXX"; Tab(42); "######.##"; Tab(52); "######.##"
    Else
      Print #UBRptA, Tab(37); "XXXXX"; Tab(50); "######.##"
    End If
    Print #UBRptA, "~" 'Per Dale
Return

PrintNewStandV1Mask:      '4
    If InStr(TOWNNAME$, "RICH CREEK") Then
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      If OkiMode = 1 Then
        Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
      Else
        Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
      End If
    End If
    Print #UBRptA, "~"; Tab(50); "########"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    If BLType = 18 Then
      Print #UBRptA, " "
    End If
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "##########";
    Print #UBRptA, Tab(15); "XXXXXXXXXXXXXXXXXXX"; Tab(50); "########";
    Print #UBRptA, Tab(62); "XX/XX/XXXX"
    Print #UBRptA, " "

    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    Print #UBRptA, Tab(3); "XX/XX/XXXX";
    Print #UBRptA, Tab(15); "XX/XX/XXXX";
    Print #UBRptA, Tab(26); "XX/XX/XXXX";
    Print #UBRptA, Tab(40); "XXXX"
    If BLType = 18 Then
      Print #UBRptA, " "
    End If

    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    If Not BLType = 18 Then
      Print #UBRptA, " "
    Else
      Print #UBRptA,
    End If

    PCnt = 0
    For PCnt = 1 To 8
        Print #UBRptA, " "; "XXXXX";
        Print #UBRptA, Tab(36); "#####.##";
     
      Select Case PCnt
      Case 2
        If AcctBar = True Then
    '*************For Okidata to print Bar code
          Print #UBRptA, Tab(47); Chr$(27); Chr$(16); "A";
          Print #UBRptA, Chr$(8);
          Print #UBRptA, Chr$(2); "0";
          Print #UBRptA, "0"; Chr$(2);
          Print #UBRptA, Chr$(1); Chr$(1);
          Print #UBRptA, Chr$(1); Chr$(2);
          Print #UBRptA, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); "9999"
    '**************************
        Else
          Print #UBRptA, " "
        End If
      Case 5
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 6
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 7
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 8
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXX"; " "; "XX"; " "; "XXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    If Not BLType = 18 Then
     Print #UBRptA, "                     Other:  "; Tab(36); "#####.##"
     Print #UBRptA, "                       Tax:  "; Tab(34); "$###,###.##"
    End If
    Print #UBRptA, "                  Previous:  "; Tab(34); "$###,###.##"
    Print #UBRptA, "                   Current:  "; Tab(34); "$###,###.##"
    Print #UBRptA, Tab(32); "---------------"
    Print #UBRptA, "                     Total:  "; Tab(34); "$###,###.##"
    If Bucksport = False Then
      If Not BLType = 18 Then
        Print #UBRptA, " "
      End If
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XX-XXXXXXX"
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      If Not BLType = 18 Then
        Print #UBRptA, " "
      End If
    Else
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XX-XXXXXXX"
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Print #UBRptA, " "
      Print #UBRptA, " "
    End If
    Print #UBRptA, "~"
Return
PrintNewStandBarMask:  '3
    If InStr(TOWNNAME$, "PEACHLAND") Then
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    ElseIf OkiMode = 1 Then
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
    End If
    Print #UBRptA, "~"; Tab(50); "########"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(18); "XX/XX/XXXX"; Tab(33); "XX/XX/XXXX";
    Print #UBRptA, "      "; "####"
    Print #UBRptA, " "
    Print #UBRptA, " "

    MPCnt = 1
    PCnt = 0
    For PCnt = 1 To 6
      Print #UBRptA, "XXX";
      Print #UBRptA, Tab(4); "########";
      Print #UBRptA, Tab(13); "########";
      Print #UBRptA, Tab(23); "#######";
      Print #UBRptA, Tab(33); "#####.##";
      Select Case PCnt
      Case 1
        Print #UBRptA, Tab(45); "##########"
      Case 5
        Print #UBRptA, Tab(45); "XXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    Print #UBRptA, Tab(14); "     TAX:"; Tab(31); "$##,###.##"
    Print #UBRptA, Tab(14); "Previous:"; Tab(31); "$##,###.##"
    Print #UBRptA, Tab(14); " Current:"; Tab(31); "$##,###.##"
    Print #UBRptA, Tab(14); " Deposit:"; Tab(31); "$##,###.##";
    Print #UBRptA, Tab(45); "XX/XX/XXXX"; Tab(60); "XX/XX/XXXX"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(2); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(36); "Total: "; Tab(44); "$###,###.##"
    Print #UBRptA, " "
    If AcctBar = True Then
    '*************For Okidata to print Bar code
      Print #UBRptA, Tab(55); Chr$(27); Chr$(16); "A"; 'String$(50, " ")
      Print #UBRptA, Chr$(8);
      Print #UBRptA, Chr$(2); "0";
      Print #UBRptA, "0"; Chr$(2);
      Print #UBRptA, Chr$(1); Chr$(1);
      Print #UBRptA, Chr$(1); Chr$(2);
      Print #UBRptA, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); "0001"
'**************************
    Else
      Print #UBRptA, " "
    End If
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "##########"; Tab(22); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX";
    Print #UBRptA, Tab(55); "xx-xxxxxx"
    Print #UBRptA, Tab(22); "XXXXXXXXXXXXXX"; " XX  XXXXX"
    Print #UBRptA, " "
    'Print #UBRptA, " "
    Print #UBRptA, "#######.##"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "#######.##";
    If PostBar = True Then
      Print #UBRptA, Tab(22); Chr$(27); Chr$(16); "C"; Chr$(Len("11111")); "11111"
    Else
      Print #UBRptA, " "
    End If
    Print #UBRptA, " "
    Print #UBRptA, "~"
Return
PrintNewStandRmStampMask: '5
    If OkiMode = 1 Then
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(77);
    Else
      Print #UBRptA, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
    End If

    Print #UBRptA, "~"; Tab(50); "########"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, "   ######";
    Print #UBRptA, Tab(15); "XXXXXXXXXXXXXXXXXXX"; Tab(50); "   #####";
    Print #UBRptA, Tab(62); "XX/XX/XXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    Print #UBRptA, Tab(3); "XX/XX/XXXX"; ; Tab(15); "XX/XX/XXXX"; Tab(26); "XX/XX/XXXX";
    Print #UBRptA, Tab(40); "####"
    Print #UBRptA, Tab(50); "XX/XX/XXXX"; Tab(64); "#####.##"
    Print #UBRptA, " "
    'Print #UBRptA, " "
    'Print #UBRptA, " "
    'Print #UBRptA, " "
    PCnt = 0
    For PCnt = 1 To 6
      Print #UBRptA, " "; "XXXXX"; Tab(7); "#######"; Tab(17); "#######";
      Print #UBRptA, Tab(28); "######"; Tab(36); "#####.##";
      
      Select Case PCnt
      Case 2
        'print bar code acct
        ')))))))))))))))))))))))
        If AcctBar = True Then
    '*************For Okidata to print Bar code
          Print #UBRptA, Tab(47); Chr$(27); Chr$(16); "A";
          Print #UBRptA, Chr$(8);
          Print #UBRptA, Chr$(2); "0";
          Print #UBRptA, "0"; Chr$(2);
          Print #UBRptA, Chr$(1); Chr$(1);
          Print #UBRptA, Chr$(1); Chr$(2);
          Print #UBRptA, Chr$(27); Chr$(16); "B"; Chr$(AcctLen); "9999"
    '**************************
        Else
          Print #UBRptA, " "
        End If
    '))))))))))))))))))))))))))

      Case 5
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 6
        Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    Print #UBRptA, "                     Other:"; Tab(36); "#####.##";
    Print #UBRptA, Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "                       Tax:"; Tab(33); "$###,###.##";
    Print #UBRptA, Tab(47); "XXXXXXXXXXXXXX XX ZZZZZ"
    Print #UBRptA, "                  Previous:  "; Tab(33); "$###,###.##"
    Print #UBRptA, "                   Current:  "; Tab(33); "$###,###.##"
    Print #UBRptA, Tab(31); "--------------"
    Print #UBRptA, "                     Total:  "; Tab(33); "$###,###.##"
    Print #UBRptA, " "
    Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XX-XXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"; Tab(47); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, Tab(3); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "~"
Return
PrnStand24L2BxMask:   '6

    Print #UBRptA, "~"
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, " "
    Print #UBRptA, Tab(4); "XX"; Tab(9); "XX"; Tab(14); "XX";
    Print #UBRptA, Tab(18); "XX"; Tab(23); "XX"; Tab(28); "XX";
    Print #UBRptA, " "
    Print #UBRptA, Tab(40); "XX/XX/XXXX"
    Print #UBRptA, " "
    'Print #UBRptA, " "
    Print #UBRptA, Tab(2); "#########";
    Print #UBRptA, Tab(12); "#########";
    Print #UBRptA, Tab(22); "########"
    Print #UBRptA, Tab(2); "#########";
    Print #UBRptA, Tab(12); "#########";
    Print #UBRptA, Tab(22); "########";
    Print #UBRptA, Tab(35); "XXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "
    Print #UBRptA, Tab(45); "LOC:XX-XXXXXX"

    Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
    PCnt = 0
    For PCnt = 1 To 6
        Print #UBRptA, Tab(3); "XXXXX";
        Print #UBRptA, Tab(20); "#####.##";
      Select Case PCnt
      Case 1
        Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 2
        Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXXXXXXXXXXXX"
      Case 3
        Print #UBRptA, Tab(34); "XXXXXXXXXXXXXXX XX XXXXX"
      Case Else
        Print #UBRptA, " "
      End Select
    Next
    'Print #UBRptA, " "
    Print #UBRptA, Tab(3); "Previous:"; Tab(20); "######.##"
    Print #UBRptA, " "
    Print #UBRptA, Tab(8); "XXXXX"; Tab(20); "#####.##";
    If BLType = 7 Then
      Print #UBRptA, Tab(35); "XXXXX"; Tab(42); "#####.##"; Tab(52); "#####.##"
    Else
      Print #UBRptA, Tab(37); "XXXXX"; Tab(50); "#####.##"
    End If
    Print #UBRptA, " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"; " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"; " "; "XXXXXXXXXXXXXXXXXXXXXXXXX"
    Print #UBRptA, "~"

Return
PrnStand24L3BxMask:  '7
  GoSub PrnStand24L2BxMask
Return
End Sub
'    PCnt = 0
'    For WRevCnt = 1 To 6
'      PCnt = PCnt + 1
'      If UBBillRec(1).RevAmt(WRevCnt) <> 0 Then
'        Print #UBRpt, Left$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
'        Select Case PCnt
'        Case 1, 2  'water/sewer
'          If WFoundMtr Then
'            Print #UBRpt, Tab(4); Using; "#########"; WPrevRead&;
'            Print #UBRpt, Tab(14); Using; "#########"; WCurrRead&;
'            Print #UBRpt, Tab(25); Using; "#######"; WUsageAmt&;
'          End If
'        End Select
'        'IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
'        '  UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
'        '  IF UsageAmt& < 0 THEN
'        '    MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) -1)
'        '    UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
'        '  END IF
'        '  PRINT #UBRpt, TAB(4); USING "#########"; UBBillRec(1).PrevRead(WRevcnt);
'        '  PRINT #UBRpt, TAB(14); USING "#########"; UBBillRec(1).CurRead(WRevcnt);
'        '  PRINT #UBRpt, TAB(25); USING "#######"; UsageAmt&;
'        'END IF
'        Print #UBRpt, Tab(33); Using; "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
'      End If
'      Select Case PCnt
'      Case 1
'        Print #UBRpt, Tab(44); Using; "##########"; UBBillRec(1).CustAcctNo
'      Case 5
'        Print #UBRpt, Tab(49); Left$(UBCustRec(1).ServAddr, 26)
'      Case Else
'        Print #UBRpt,
'      End Select
'    Next


'***************************************************
