VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmLateNoticeLabels 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Late Notice Mailing Labels"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12216
   Icon            =   "frmLateNoticeLabels.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboLabelType 
      Height          =   348
      Left            =   4980
      TabIndex        =   0
      Top             =   5112
      Width           =   4548
      _Version        =   196608
      _ExtentX        =   8022
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
      ColDesigner     =   "frmLateNoticeLabels.frx":08CA
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
      TabIndex        =   2
      Top             =   6792
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
      TabIndex        =   1
      Top             =   6792
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
            TextSave        =   "10:56 AM"
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
   Begin EditLib.fpText fptxtNoticeCnt 
      Height          =   348
      Left            =   4980
      TabIndex        =   4
      Top             =   2844
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
      BackColor       =   -2147483648
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
      NoSpecialKeys   =   3
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
      ControlType     =   1
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
   Begin EditLib.fpCurrency fpMinBal 
      Height          =   348
      Left            =   4980
      TabIndex        =   9
      Top             =   4548
      Width           =   1476
      _Version        =   196608
      _ExtentX        =   2603
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
      BackColor       =   -2147483648
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   3
      AutoAdvance     =   -1  'True
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
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
   Begin EditLib.fpDateTime txtNoticeDate 
      Height          =   348
      Left            =   4980
      TabIndex        =   10
      Top             =   3408
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
      BackColor       =   -2147483648
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
      NoSpecialKeys   =   3
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
   Begin EditLib.fpText fptxtPrintOrder 
      Height          =   348
      Left            =   4992
      TabIndex        =   13
      Top             =   3984
      Width           =   3444
      _Version        =   196608
      _ExtentX        =   6075
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
      BackColor       =   -2147483648
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
      NoSpecialKeys   =   3
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   "0123456789"
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Balance:"
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
      Left            =   2700
      TabIndex        =   12
      Top             =   4596
      Width           =   2124
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notice Date:"
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
      Left            =   3324
      TabIndex        =   11
      Top             =   3456
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Late Notice Mailing Labels"
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
      TabIndex        =   8
      Top             =   1176
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   936
      Width           =   5772
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notice Count:"
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
      Left            =   3204
      TabIndex        =   7
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Label Type: "
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
      Left            =   2724
      TabIndex        =   6
      Top             =   5160
      Width           =   2172
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3372
      Left            =   2160
      Top             =   2496
      Width           =   7908
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
      Index           =   7
      Left            =   3108
      TabIndex        =   5
      Top             =   4008
      Width           =   1716
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3192
      Top             =   816
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmLateNoticeLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CycleFlag As Boolean, OKFlag As Boolean, BadDate As Boolean
Dim ErFlag As Boolean, LNType As Integer
Private Sub cmdExit_Click()
  Load frmUBLateNoticeMenu
  DoEvents
  frmUBLateNoticeMenu.Show
  Unload Me
  DoEvents
End Sub

Private Sub cmdPrint_Click()
  NoticeMailLabel
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
        UBLog "Closed via LateNoticeLabels by " + PWUser$
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

Private Sub Form_Load()
  Dim UBSetupreclen As Integer, NIfile As Integer, lenNI As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUp(1))
  LoadUBSetUpFile UBSetUp(), UBSetupreclen
  NIfile = FreeFile
  ReDim NoticeInfo(1) As NoticeInfoType
  lenNI = Len(NoticeInfo(1))
  Open UBPath$ + "UBLNINFO.DAT" For Random Shared As NIfile Len = lenNI
  Get NIfile, 1, NoticeInfo(1)
  Close
  fptxtNoticeCnt = NoticeInfo(1).PrnCnt
  txtNoticeDate = Num2Date(NoticeInfo(1).NoticeDate)
  fpMinBal = NoticeInfo(1).MinBalance
  Select Case NoticeInfo(1).PrnOrder
    Case 1:
      fptxtPrintOrder = "Customer Name Order"
    Case 2:
      fptxtPrintOrder = "Account Number Order"
    Case 3:
      fptxtPrintOrder = "Location Number Order"
    Case 4:
      fptxtPrintOrder = "ZipCode Order"
    Case 5:
      fptxtPrintOrder = "Zip/Location Order"
  End Select
  fpcboLabelType.InsertRow = "1) 1 X 3 1/2 (1-Label Wide)Text"
 ' fpcboLabelType.InsertRow = "2) 1 X 3 1/2 (3-Labels Wide)Text"
  fpcboLabelType.InsertRow = "2) 1 X 3 1/2 (4-Labels Wide)Text"
  fpcboLabelType.InsertRow = "3) 1 X 2 5/8 (Full Sheet 3 Wide)Graphics"
  fpcboLabelType.ListIndex = 0

End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub

Private Sub NoticeMailLabel()
  Dim lcnt As Integer, NIfile As Integer, lenNI As Integer
  Dim UBCustRecLen As Integer, IdxRecLen As Long, IdxFileSize As Long
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim IndexName As String, cnt As Long, UBCustM As Integer
  Dim RPTFile As String, UBRpt As Integer, AcctNumber As Long
  Dim LType As Integer, DidCnt As Integer, LabelCnt As Integer
  Dim PCnt As Integer, CustName As String, ADDR1 As String
  Dim Zip As String, CityStaZip As String, ADDR2 As String
  Dim TP1 As String, TP2 As String, TP3 As String, TP4 As String
  Dim xbar As Boolean, MaskLabel As String, Align As String, TP5 As String
  ReDim ToPrint(1 To 5) As String * 132
  TP1$ = ""
  TP2$ = ""
  TP3$ = ""
  TP4$ = ""
  TP5$ = ""
  For lcnt = 1 To 5
    LSet ToPrint(lcnt) = ""
  Next
  LType = fpcboLabelType.ListIndex + 1
  ReDim OSet(1 To 4) As Integer
  xbar = False
  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110

  If FileSize&(UBPath$ + "UBLNINFO.DAT") = 0 Then
    GoTo ExitMailLabListing:
  End If
  NIfile = FreeFile
  ReDim NoticeInfo(1) As NoticeInfoType
  lenNI = Len(NoticeInfo(1))
  Open UBPath$ + "UBLNINFO.DAT" For Random Shared As NIfile Len = lenNI
  Get NIfile, 1, NoticeInfo(1)
  Close

  'FGetAH "UBLNINFO.DAT", NoticeInfo, Len(NoticeInfo), 1
'***************

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  'Erase frm, Form$, Fld, Choice$
  IndexName$ = UBPath$ + "UBLNIDX.DAT"
  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize&(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  'FGetAH "UBLNIDX.DAT", IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs
  Handle = FreeFile
  Open IndexName$ For Random Shared As Handle Len = IdxRecLen
  For cnt& = 1 To IdxNumOfRecs
    Get #Handle, cnt&, IdxBuff(cnt&)
  Next
  Close Handle

  UBCustM = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustM Len = UBCustRecLen
  RPTFile$ = UBPath$ + "UBLNMAIL.RPT"
  UBRpt = FreeFile
  Open RPTFile$ For Output As UBRpt

'  BlockClear
'  ShowProcessingScrn "Mailing Labels"

  For cnt = 1 To NumOfRecs
    AcctNumber& = IdxBuff(cnt).RecNum
    Get UBCustM, AcctNumber&, UBCustRec(1)

    Select Case LType
    Case 1
      Print #UBRpt, " " '"Cust #" + Str$(AcctNumber&)
      Print #UBRpt, Left$(QPTrim$(UBCustRec(1).CustName), 23)
      Print #UBRpt, Left$(QPTrim$(UBCustRec(1).ADDR1), 23)
      If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
        Print #UBRpt, Left$(QPTrim$(UBCustRec(1).ADDR2), 23)
        Print #UBRpt, Left$(QPTrim$(UBCustRec(1).CITY), 13) + ", " + UBCustRec(1).STATE + " " + Left$(UBCustRec(1).ZIPCODE, 5)
      Else
        Print #UBRpt, Left$(QPTrim$(UBCustRec(1).CITY), 13) + ", " + UBCustRec(1).STATE + " " + Left$(UBCustRec(1).ZIPCODE, 5)
        Print #UBRpt, " "
      End If
      Print #UBRpt, " "
      DidCnt = DidCnt + 1
    Case 2
      LabelCnt = LabelCnt + 1
      Mid$(ToPrint(1), OSet(LabelCnt)) = " " '"Cust #" + Str$(AcctNumber&)
      Mid$(ToPrint(2), OSet(LabelCnt)) = Left$(QPTrim$(UBCustRec(1).CustName), 23)
      Mid$(ToPrint(3), OSet(LabelCnt)) = Left$(QPTrim$(UBCustRec(1).ADDR1), 23)
      If Len(QPTrim$(UBCustRec(1).ADDR2)) > 0 Then
        Mid$(ToPrint(4), OSet(LabelCnt)) = Left$(QPTrim$(UBCustRec(1).ADDR2), 23)
        Mid$(ToPrint(5), OSet(LabelCnt)) = Left$(QPTrim$(UBCustRec(1).CITY), 13) + ", " + UBCustRec(1).STATE + " " + Left$(UBCustRec(1).ZIPCODE, 5)
      Else
        Mid$(ToPrint(4), OSet(LabelCnt)) = Left$(QPTrim$(UBCustRec(1).CITY), 13) + ", " + UBCustRec(1).STATE + " " + Left$(UBCustRec(1).ZIPCODE, 5)
      End If
      If LabelCnt = 4 Then
        For PCnt = 1 To 5
          'LPRINT QPTrim$(ToPrint(PCnt))
          Print #UBRpt, ToPrint(PCnt)
          LSet ToPrint(PCnt) = " "
        Next
        Print #UBRpt, " "
        LabelCnt = 0
      End If
      Case 3:
        CustName$ = QPTrim$(UBCustRec(1).CustName)
        If Len(CustName$) < 1 Then GoTo NextLabel
        ADDR1$ = QPTrim$(UBCustRec(1).ADDR1)
        ADDR2$ = QPTrim$(UBCustRec(1).ADDR2)
        CityStaZip$ = Left$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + UBCustRec(1).ZIPCODE
        LabelCnt = LabelCnt + 1
        TP1$ = TP1$ + Left$(CustName$, 23) + "~"
        TP2$ = TP2$ + Left$(ADDR1$, 23) + "~"
        If Len(ADDR2$) > 0 Then
          TP3$ = TP3$ + Left$(ADDR2$, 23) + "~"
          TP4$ = TP4$ + CityStaZip$ + "~"
          TP5$ = TP5$ + " ~"
        Else
         TP3$ = TP3$ + CityStaZip$ + "~"
         TP4$ = TP4$ + " ~"
         TP5$ = TP5$ + " ~"
        End If
        If LabelCnt = 3 Then
         ' For cnt = 1 To 4
            'LPRINT QPTrim$(ToPrint(PCnt))
            Print #UBRpt, TP1$ + TP2$ + TP3$ + TP4$ + TP5$
            TP1$ = ""
            TP2$ = ""
            TP3$ = ""
            TP4$ = ""
            TP5$ = ""
  '        Next
  '        Print #UBRpt, Filler$
          LabelCnt = 0
        Else
          
        End If

    End Select

'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If

NextLabel:
  'ShowPctComp cnt, NumOfRecs
  'IF didcnt > 4 THEN EXIT FOR
  Next

  If LType = 2 Then
    If LabelCnt > 0 Then
      For PCnt = 1 To 5
        Print #UBRpt, QPTrim$(ToPrint(PCnt))
      Next
      Print #UBRpt, " "
    End If
  End If
  If LType = 3 Then
    If LabelCnt = 1 Then
      TP1$ = TP1$ + " ~ ~"
      TP2$ = TP2$ + " ~ ~"
      TP3$ = TP3$ + " ~ ~"
      TP4$ = TP4$ + " ~ ~"
      TP5$ = TP5$ + " ~ ~"
      Print #UBRpt, TP1$ + TP2$ + TP3$ + TP4$ + TP5$
    ElseIf LabelCnt = 2 Then
      TP1$ = TP1$ + " ~"
      TP2$ = TP2$ + " ~"
      TP3$ = TP3$ + " ~"
      TP4$ = TP4$ + " ~"
      TP5$ = TP5$ + " ~"
      Print #UBRpt, TP1$ + TP2$ + TP3$ + TP4$ + TP5$
    End If
  End If
  If LType <> 3 Then
    Print #UBRpt, Chr$(12);
  End If
  Close

  Erase IdxBuff, UBCustRec, ToPrint
 ' Erase frm, Form$, Fld, OSet

'  If Not AbortFlag Then
'    PrintRptFile "Mailing Labels", "UBLNMAIL.RPT", 1, RetCode, EntryPoint
'  End If
  'KillFile "UBLABEL.RPT"
  GoSub DoAlignLabelMask
  If LType <> 3 Then
    ViewPrint RPTFile$, "Mailing Labels", xbar, , True, MaskLabel$
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmLateNoticeLabels
    ARptMailLabels.GetName RPTFile$
    ARptMailLabels.startrpt
  End If
ExitMailLabListing:

  Exit Sub
DoAlignLabelMask:
  ReDim OSet(1 To 4) As Integer

  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  xbar = False
'  ReDim TempScrn(0)
'  SaveScrn TempScrn()
  Align$ = String$(34, "X")
  UBRpt = FreeFile
  MaskLabel$ = UBPath$ + "UBLblA.RPT"
  Open UBPath$ + "UBLblA.RPT" For Output As UBRpt
  Select Case LType
  Case 1
    For cnt = 1 To 5
      Print #UBRpt, Align$
    Next
    Print #UBRpt,
    xbar = False
  Case 2
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$; Tab(OSet(4)); Align$
    Next
    Print #UBRpt,
    xbar = True
  End Select
  Close UBRpt

Return


End Sub
