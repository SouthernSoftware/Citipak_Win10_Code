VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmUtilStripTrans 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Strip Utility"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmUTilStripTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboTransType 
      Height          =   375
      Left            =   5130
      TabIndex        =   3
      Top             =   4275
      Width           =   3900
      _Version        =   196608
      _ExtentX        =   6879
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
      ColDesigner     =   "frmUTilStripTrans.frx":08CA
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include all transactions prior to date entered."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2832
      TabIndex        =   14
      Top             =   5520
      Width           =   6468
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   8
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   582
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
            TextSave        =   "4:00 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "7/20/2018"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   8910
      TabIndex        =   6
      Top             =   7170
      Width           =   1320
      _Version        =   131072
      _ExtentX        =   2328
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmUTilStripTrans.frx":0BED
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOK 
      Height          =   480
      Left            =   7320
      TabIndex        =   5
      Top             =   7200
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmUTilStripTrans.frx":0DC9
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5112
      TabIndex        =   0
      Top             =   2664
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
   Begin EditLib.fpText fptxtRoute2 
      Height          =   348
      Left            =   5136
      TabIndex        =   2
      Top             =   3708
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
      Left            =   5136
      TabIndex        =   1
      Top             =   3192
      Width           =   636
      _Version        =   196608
      _ExtentX        =   1122
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
   Begin EditLib.fpText txtOperator 
      Height          =   348
      Left            =   5136
      TabIndex        =   4
      Top             =   4824
      Width           =   804
      _Version        =   196608
      _ExtentX        =   1418
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
      MaxLength       =   4
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
   Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
      Height          =   480
      Left            =   1830
      TabIndex        =   15
      Top             =   7170
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmUTilStripTrans.frx":0FA3
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   312
      Index           =   0
      Left            =   2712
      TabIndex        =   13
      Top             =   4836
      Width           =   2304
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
      Height          =   372
      Left            =   3576
      TabIndex        =   12
      Top             =   3252
      Width           =   1476
   End
   Begin VB.Label LabelB2 
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
      Height          =   372
      Left            =   3672
      TabIndex        =   11
      Top             =   3768
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date:"
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
      Index           =   1
      Left            =   2904
      TabIndex        =   10
      Top             =   2712
      Width           =   2088
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   4452
      Left            =   2280
      Top             =   2256
      Width           =   7332
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type:"
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
      Left            =   2928
      TabIndex        =   9
      Top             =   4320
      Width           =   2088
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Strip Transactions"
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
      Left            =   3288
      TabIndex        =   7
      Top             =   1608
      Width           =   5652
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3228
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3228
      Top             =   1248
      Width           =   5772
   End
End
Attribute VB_Name = "frmUtilStripTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim Oper As String
Dim BegRoute As String, EndRoute As String
Dim eastflag As Boolean
Private Sub cmdOk_Click()
  If Check1.Value = ValueTrue Then
    eastflag = True
  Else
    eastflag = False
  End If
  If eastflag = True Then   'for all trans prior to or equal to date
    StripemEast
  Else
    Stripem
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'CMLog "Closed via SelectPaySource by " + PWUser$ + " operator-" + Oper$
       ' CitiTerminate
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
      Call fpCmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call cmdOk_Click
    Case vbKeyF8:
      KeyCode = 0
      DoEvents
      Call fpBtn1_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  fpcboTransType.AddItem " 1) - Utility Bill"
  fpcboTransType.AddItem " 4) - Payment"
  fpcboTransType.AddItem " 6) - Penalty Charge"
  fpcboTransType.AddItem " 7) - Deposit Payment"
  fpcboTransType.AddItem "11) - Up Adjustment"
  fpcboTransType.AddItem "12) - Down Adjustment"
  fpcboTransType.AddItem "33) - Payment Adjustment"
  fpcboTransType.ListIndex = 0
  fptxtRoute1 = "00"
  fptxtRoute2 = "99"
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
'
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub

Private Sub fpBtn1_Click()
StripemNOBALADJ
 ' Stripem4Perq
End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtRoute1.SetFocus
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
    fpcboTransType.SetFocus
  End If
End Sub

Private Sub fptxtRoute2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpcboTransType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTransType.ListDown = True
  End If
  If fpcboTransType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtOperator.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtRoute2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub txtOperator_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdOK.SetFocus
  End If
End Sub

Private Sub fpCmdExit_Click()
  frmUBEditMenu.Show
  Unload Me
End Sub

Private Sub Stripem()
  Dim Date1 As Integer, UBTranRecLen As Integer, CustLen As Integer
  Dim UBFile1 As Integer, UBFile2 As Integer, UBFile3 As Integer
  Dim TNumOfRecs As Long, cnt As Long, RCnt As Integer
  Dim TrType As String, TrTyp As Integer, CustBook As Integer
  Dim Removed As Long, FromBook As Integer, ThruBook As Integer
  Dim operchk As Integer
  Date1 = Date2Num(txtDate1.Text)
  FromBook = Val(fptxtRoute1)
  ThruBook = Val(fptxtRoute2)
  If fpcboTransType.ListIndex <> -1 Then
    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
    TrTyp = Val(TrType$)
  Else
    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
    Exit Sub
  End If
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
'  If TrTyp = 0 Then
'    BegTrans = 1
'    EndTrans = 999
'  Else
'    BegTrans = TrTyp
'    EndTrans = TrTyp
'  End If
  operchk = Val(txtOperator)
  DeActivateControls frmUtilStripTrans
  FrmShowPctComp.Label1 = "Gathering Transactions to Remove"
  FrmShowPctComp.Show , Me
  UBLog "StripTrans - " & txtDate1.Text & ", book(" & fptxtRoute1 & "-" & fptxtRoute2 & "),Trans-" & TrType$
  ReDim UBCust(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  CustLen = Len(UBCust(1))
  Removed = 0

  UBFile1 = FreeFile
  Open "UBTRANS.dat" For Random Shared As UBFile1 Len = UBTranRecLen
  UBFile3 = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile3 Len = CustLen

  TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
  For cnt& = 1 To TNumOfRecs&
    FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
    Get UBFile1, cnt&, UBTranRec(1)
    
    If (UBTranRec(1).TransDate = Date1) Then 'AND (UBTranRec(1).OperatorNumber = 4) THEN
'      Select Case UBTranRec(1).TransType
'      Case TranUtilityBill
      If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
        If UBTranRec(1).CustAcctNo <= 0 Then
          UBLog "InvAcct1:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
          GoTo InvAcct1
        End If
        Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        CustBook = Val(UBCust(1).Book)
        If CustBook >= FromBook And CustBook <= ThruBook Then
         If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
          Select Case TrType
            Case 1  'Utility Bill
              Removed = Removed + 1
            Case 4   'payment
              Removed = Removed + 1
            Case 6    'Penalty charge
              Removed = Removed + 1
            Case 7    'deposit payment
              Removed = Removed + 1
            Case 11   '"Bill-Upward Adjustment"
              Removed = Removed + 1
            Case 12   '"Bill-Downward Adjustment"
              Removed = Removed + 1
            Case 33   '"Payment Adjustment"
              Removed = Removed + 1
            Case Else
            End Select
          End If
        End If
      End If
    End If
InvAcct1:
  Next
  Close
  If Removed > 0 Then
    If MsgBox("Num of Trans to be Removed: " & Removed & " Yes to Remove, No to Cancel", vbYesNo, "Num to Remove") = vbYes Then
      KillFile "UBTRANS.bak"
      Name "UBTRANS.DAT" As "UBTRANS.bak"
    
      UBFile1 = FreeFile
      Open "UBTRANS.bak" For Random Shared As UBFile1 Len = UBTranRecLen
    
      UBFile2 = FreeFile
      Open "UBTRANS.DAT" For Random Shared As UBFile2 Len = UBTranRecLen
    
      UBFile3 = FreeFile
      Open "UBCUST.DAT" For Random Shared As UBFile3 Len = CustLen
    
      TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
 
      
      Removed = 0
      FrmShowPctComp.Label1 = "Removing Transactions"
      FrmShowPctComp.Show , Me
      For cnt& = 1 To TNumOfRecs&
        FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          Exit Sub
        End If
        Get UBFile1, cnt&, UBTranRec(1)
        
        If (UBTranRec(1).TransDate = Date1) Then 'AND (UBTranRec(1).OperatorNumber = 4) THEN
    '      Select Case UBTranRec(1).TransType
    '      Case TranUtilityBill
          If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
            If UBTranRec(1).CustAcctNo <= 0 Then
              UBLog "InvAcct:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
              GoTo InvAcct
            End If
            Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
            CustBook = Val(UBCust(1).Book)
            If CustBook >= FromBook And CustBook <= ThruBook Then
             If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
              Select Case TrType
                Case 1  'Utility Bill
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 4   'payment
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) + UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance + UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 6    'Penalty charge
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 7    'deposit payment
                  UBCust(1).DepositAmt = uRound#(UBCust(1).DepositAmt - UBTranRec(1).Transamt)
                  If UBCust(1).DepositAmt < 0 Then UBCust(1).DepositAmt = 0
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 11   '"Bill-Upward Adjustment"
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound#(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 12   '"Bill-Downward Adjustment"
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) + UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance + UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 33   '"Payment Adjustment"
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound#(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case Else
                  Put UBFile2, , UBTranRec(1)
                End Select
              Else
               Put UBFile2, , UBTranRec(1)
              End If
            Else
             Put UBFile2, , UBTranRec(1)
            End If
          Else
            Put UBFile2, , UBTranRec(1)
          End If
        Else
          Put UBFile2, , UBTranRec(1)
        End If
InvAcct:
      Next
      Close
      'ActivateControls frmUtilStripTrans
      UBLog "Removed:" & Removed & " Using Strip Trans Util"
      MsgBox "Removed:" & Removed, vbOKOnly, "Removed Trans"
      UBRelinkTransactions
    Else
      Close
      ActivateControls frmUtilStripTrans
      UBLog "None Removed, Canceled by user, Using Strip Trans Util"
    End If
  Else
    Close
    ActivateControls frmUtilStripTrans
    UBLog "No Trans Removed 0 to remove in first pass Using Strip Trans Util"
    MsgBox "No Transactions met criteria to be Removed:" & Removed, vbOKOnly, "Removed 0 Trans"
  End If
End Sub


Private Sub UBRelinkTransactions()
  
  DoEvents
  UBLog " IN: Relink Utility Files"
  
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer, WorkOrderRecLen As Integer
  Dim UBFile As Integer, UBTran As Integer, UBWrkOrd As Integer
  Dim NumOfCRecs As Long, NumOfTRecs As Long, NumOfWORecs As Long
  Dim OddRecs As Integer, RecCnt As Long
  Dim TRRecs As Long, PutRec As Long
  Dim CCnt As Long, ChkCnt As Long
  Dim BlockSize As Long
  Dim NumChunks As Long
  Dim MaxBlockCnt As Integer
  Dim UBRTCustRec   As NewUBCustRecType
  Dim UBRTTransRec  As UBTransRecType
  Dim WorkOrderRec As WorkOrderRecType

  UBCustRecLen = Len(UBRTCustRec)              'Length of Cust Record Structure
  UBTranRecLen = Len(UBRTTransRec)             'Length of Tran Record Structure
  WorkOrderRecLen = Len(WorkOrderRec)
  'DeActivateControls frmUtilStripTrans
  
  DoEvents
  
  FrmShowPctComp.Label1 = "Checking Customers."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  
  DoEvents
  
  UBLog "BEGIN: Pass 1 of 3"
  
  UBTran = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  NumOfTRecs = LOF(UBTran) \ UBTranRecLen

  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfCRecs = LOF(UBFile) \ UBCustRecLen
    
  For CCnt = 1 To NumOfCRecs
    Get UBFile, CCnt, UBRTCustRec
    UBRTCustRec.LastTrans = 0
    UBRTCustRec.WOLastTrans = 0
    Put UBFile, CCnt, UBRTCustRec
    ChkCnt = ChkCnt + 1
    If ChkCnt >= 100 Then
      FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs
      ChkCnt = 0
    End If
  Next
  
  Unload FrmShowPctComp
  
  DoEvents
  FrmShowPctComp.Label1 = "Relinking Transactions."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  DoEvents
    
  UBLog "       Pass 2 of 3"
  MaxBlockCnt = 1024
  
  ReDim TransBuff(1 To MaxBlockCnt) As UBTransRecType
''************************************
  NumChunks& = NumOfTRecs \ MaxBlockCnt
''****DO NOT CHANGE THE DIVISION HERE!
  OddRecs = UBMod(NumOfTRecs, MaxBlockCnt)

  If NumChunks& = 0 Then        'if the actual cust count is less than
    MaxBlockCnt = OddRecs       'the work buffer
    NumChunks& = 1
    OddRecs = 0
  End If
  
  For CCnt& = 1 To NumChunks&
    For RecCnt = 1 To MaxBlockCnt
      TRRecs = TRRecs + 1
      Get UBTran, TRRecs, TransBuff(RecCnt)
      TransBuff(RecCnt).PenAtBill = TRRecs
    Next
    For RecCnt = 1 To MaxBlockCnt
      If (TransBuff(RecCnt).CustAcctNo > 0) And (TransBuff(RecCnt).CustAcctNo <= NumOfCRecs) Then
        Get UBFile, TransBuff(RecCnt).CustAcctNo, UBRTCustRec
        TransBuff(RecCnt).PrevTrans = UBRTCustRec.LastTrans
        PutRec = TransBuff(RecCnt).PenAtBill
        UBRTCustRec.LastTrans = PutRec
        Put UBFile, TransBuff(RecCnt).CustAcctNo, UBRTCustRec
        Put UBTran, PutRec, TransBuff(RecCnt)
      End If
    Next
    FrmShowPctComp.ShowPctComp TRRecs, NumOfTRecs
  Next
  
  If OddRecs Then
    For CCnt = TRRecs + 1 To NumOfTRecs
      Get UBTran, CCnt, TransBuff(1)
      TransBuff(1).PenAtBill = CCnt
      If (TransBuff(1).CustAcctNo > 0) And (TransBuff(1).CustAcctNo <= NumOfCRecs) Then
        Get UBFile, TransBuff(1).CustAcctNo, UBRTCustRec
        TransBuff(1).PrevTrans = UBRTCustRec.LastTrans
        PutRec = TransBuff(1).PenAtBill
        UBRTCustRec.LastTrans = PutRec
        Put UBFile, TransBuff(1).CustAcctNo, UBRTCustRec
        Put UBTran, PutRec, TransBuff(1)
      End If
      FrmShowPctComp.ShowPctComp CCnt, NumOfTRecs
    Next
  End If
  Close UBTran
  Unload FrmShowPctComp
  DoEvents
  Erase TransBuff
  
  UBLog "       Pass 3 of 3"
  
  FrmShowPctComp.Label1 = "Relinking WorkOrders."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show
  DoEvents
    
  UBWrkOrd = FreeFile
  Open UBPath + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
  NumOfWORecs& = LOF(UBWrkOrd) \ WorkOrderRecLen
 
  For CCnt = 1 To NumOfWORecs
    Get UBWrkOrd, CCnt, WorkOrderRec
    If (WorkOrderRec.CustRec > 0) And (WorkOrderRec.CustRec <= NumOfCRecs) Then
      Get UBFile, WorkOrderRec.CustRec, UBRTCustRec
      WorkOrderRec.PrevTransRec = UBRTCustRec.WOLastTrans
      UBRTCustRec.WOLastTrans = CCnt
      Put UBFile, WorkOrderRec.CustRec, UBRTCustRec
      Put UBWrkOrd, CCnt&, WorkOrderRec
    End If
    FrmShowPctComp.ShowPctComp CCnt, NumOfWORecs
  Next
  Unload FrmShowPctComp
  
  Close
  DoEvents
  UBLog "RELINK: Utility Files Completed."
  ReIndexSystem False
  ActivateControls frmUtilStripTrans
  MsgBox "Relink/Reindex Complete", vbOKOnly, "Completed"
  
  'Unload frmDataUpdated
  
  DoEvents

ExitRelink:
  UBLog "OUT: Relink Transaction History" + CrLf$
End Sub
Private Sub StripemEast()
  Dim Date1 As Integer, UBTranRecLen As Integer, CustLen As Integer
  Dim UBFile1 As Integer, UBFile2 As Integer, UBFile3 As Integer
  Dim TNumOfRecs As Long, cnt As Long, RCnt As Integer
  Dim TrType As String, TrTyp As Integer, CustBook As Integer
  Dim Removed As Long, FromBook As Integer, ThruBook As Integer
  Dim operchk As Integer
  Date1 = Date2Num(txtDate1.Text)
  FromBook = Val(fptxtRoute1)
  ThruBook = Val(fptxtRoute2)
  If fpcboTransType.ListIndex <> -1 Then
    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
    TrTyp = Val(TrType$)
  Else
    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
    Exit Sub
  End If
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
'  If TrTyp = 0 Then
'    BegTrans = 1
'    EndTrans = 999
'  Else
'    BegTrans = TrTyp
'    EndTrans = TrTyp
'  End If
  operchk = Val(txtOperator)
  DeActivateControls frmUtilStripTrans
  FrmShowPctComp.Label1 = "Gathering Transactions to Remove"
  FrmShowPctComp.Show , Me
  UBLog "StripTrans - " & txtDate1.Text & ", book(" & fptxtRoute1 & "-" & fptxtRoute2 & "),Trans-" & TrType$
  ReDim UBCust(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  CustLen = Len(UBCust(1))
  Removed = 0

  UBFile1 = FreeFile
  Open "UBTRANS.dat" For Random Shared As UBFile1 Len = UBTranRecLen
  UBFile3 = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile3 Len = CustLen

  TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
  For cnt& = 1 To TNumOfRecs&
    FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
    Get UBFile1, cnt&, UBTranRec(1)
    
    If (UBTranRec(1).TransDate <= Date1) Then 'AND (UBTranRec(1).OperatorNumber = 4) THEN
'      Select Case UBTranRec(1).TransType
'      Case TranUtilityBill
      If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
        If UBTranRec(1).CustAcctNo < 0 Then
          UBLog "InvAcct1:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
          GoTo InvAcct1
        End If
        Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        CustBook = Val(UBCust(1).Book)
        If CustBook >= FromBook And CustBook <= ThruBook Then
         If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
          Select Case TrType
            Case 1  'Utility Bill
              Removed = Removed + 1
            Case 4   'payment
              Removed = Removed + 1
            Case 6    'Penalty charge
              Removed = Removed + 1
            Case 7    'deposit payment
              Removed = Removed + 1
            Case 11   '"Bill-Upward Adjustment"
              Removed = Removed + 1
            Case 12   '"Bill-Downward Adjustment"
              Removed = Removed + 1
            Case 33   '"Payment Adjustment"
              Removed = Removed + 1
            Case Else
            End Select
          End If
        End If
      End If
    End If
InvAcct1:
  Next
  Close
  If Removed > 0 Then
    If MsgBox("Num of Trans to be Removed: " & Removed & " Yes to Remove, No to Cancel", vbYesNo, "Num to Remove") = vbYes Then
      KillFile "UBTRANS.bak"
      Name "UBTRANS.DAT" As "UBTRANS.bak"
    
      UBFile1 = FreeFile
      Open "UBTRANS.bak" For Random Shared As UBFile1 Len = UBTranRecLen
    
      UBFile2 = FreeFile
      Open "UBTRANS.DAT" For Random Shared As UBFile2 Len = UBTranRecLen
    
      UBFile3 = FreeFile
      Open "UBCUST.DAT" For Random Shared As UBFile3 Len = CustLen
    
      TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
 
      
      Removed = 0
      FrmShowPctComp.Label1 = "Removing Transactions"
      FrmShowPctComp.Show , Me
      For cnt& = 1 To TNumOfRecs&
        FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          Exit Sub
        End If
        Get UBFile1, cnt&, UBTranRec(1)
        
        If (UBTranRec(1).TransDate <= Date1) Then 'AND (UBTranRec(1).OperatorNumber = 4) THEN
    '      Select Case UBTranRec(1).TransType
    '      Case TranUtilityBill
          If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
            If UBTranRec(1).CustAcctNo < 0 Then
              UBLog "InvAcct:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
              GoTo InvAcct
            End If
            Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
            CustBook = Val(UBCust(1).Book)
            If CustBook >= FromBook And CustBook <= ThruBook Then
             If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
              Select Case TrType
                Case 1  'Utility Bill
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 4   'payment
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) + UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance + UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 6    'Penalty charge
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 7    'deposit payment
                  UBCust(1).DepositAmt = uRound#(UBCust(1).DepositAmt - UBTranRec(1).Transamt)
                  If UBCust(1).DepositAmt < 0 Then UBCust(1).DepositAmt = 0
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 11   '"Bill-Upward Adjustment"
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound#(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 12   '"Bill-Downward Adjustment"
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) + UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance + UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case 33   '"Payment Adjustment"
                  For RCnt = 1 To 15
                    UBCust(1).CurrRevAmts(RCnt) = uRound#(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
                  Next
                  UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).Transamt)
                  Put UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
                  Removed = Removed + 1
                Case Else
                  Put UBFile2, , UBTranRec(1)
                End Select
              Else
               Put UBFile2, , UBTranRec(1)
              End If
            Else
             Put UBFile2, , UBTranRec(1)
            End If
          Else
            Put UBFile2, , UBTranRec(1)
          End If
        Else
          Put UBFile2, , UBTranRec(1)
        End If
InvAcct:
      Next
      Close
      'ActivateControls frmUtilStripTrans
      UBLog "Removed:" & Removed & " Using Strip Trans Util"
      MsgBox "Removed:" & Removed, vbOKOnly, "Removed Trans"
      UBRelinkTransactions
    Else
      Close
      ActivateControls frmUtilStripTrans
      UBLog "None Removed, Canceled by user, Using Strip Trans Util"
    End If
  Else
    Close
    ActivateControls frmUtilStripTrans
    UBLog "No Trans Removed 0 to remove in first pass Using Strip Trans Util"
    MsgBox "No Transactions met criteria to be Removed:" & Removed, vbOKOnly, "Removed 0 Trans"
  End If
End Sub

Private Sub Stripem4Perq()
  Dim UBTranRecLen As Integer
  Dim UBFile1 As Integer, UBFile2 As Integer, UBFile3 As Integer
  Dim TNumOfRecs As Long, cnt As Long, RCnt As Integer
  
  Dim Removed As Long
  DeActivateControls frmUtilStripTrans
  FrmShowPctComp.Label1 = "Gathering Transactions to Remove"
  FrmShowPctComp.Show , Me
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  Removed = 0

  UBFile1 = FreeFile
  Open "UBTRANS.dat" For Random Shared As UBFile1 Len = UBTranRecLen

  TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
  For cnt& = 1 To TNumOfRecs&
    FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
    Get UBFile1, cnt&, UBTranRec(1)
    
    If UBTranRec(1).CustAcctNo >= 5450 Then
              Removed = Removed + 1
    End If
  Next
  Close
  If Removed > 0 Then
    If MsgBox("Num of Trans to be Removed: " & Removed & " Yes to Remove, No to Cancel", vbYesNo, "Num to Remove") = vbYes Then
      KillFile "UBTRANS.bak"
      Name "UBTRANS.DAT" As "UBTRANS.bak"
    
      UBFile1 = FreeFile
      Open "UBTRANS.bak" For Random Shared As UBFile1 Len = UBTranRecLen
    
      UBFile2 = FreeFile
      Open "UBTRANS.DAT" For Random Shared As UBFile2 Len = UBTranRecLen
    
    
      TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
 
      
      Removed = 0
      FrmShowPctComp.Label1 = "Removing Transactions"
      FrmShowPctComp.Show , Me
      For cnt& = 1 To TNumOfRecs&
        FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          Exit Sub
        End If
        Get UBFile1, cnt&, UBTranRec(1)
        
        If UBTranRec(1).CustAcctNo >= 5450 Then
                  Removed = Removed + 1
        Else
          Put UBFile2, , UBTranRec(1)
        End If
InvAcct:
      Next
      Close
      MsgBox "Removed:" & Removed, vbOKOnly, "Removed Trans"
      UBRelinkTransactions
    Else
      Close
      ActivateControls frmUtilStripTrans
      UBLog "None Removed, Canceled by user, Using Strip Trans Util"
    End If
  Else
    Close
    ActivateControls frmUtilStripTrans
    UBLog "No Trans Removed 0 to remove in first pass Using Strip Trans Util"
    MsgBox "No Transactions met criteria to be Removed:" & Removed, vbOKOnly, "Removed 0 Trans"
  End If
End Sub

Private Sub StripemNOBALADJ()
  Dim Date1 As Integer, UBTranRecLen As Integer, CustLen As Integer
  Dim UBFile1 As Integer, UBFile2 As Integer, UBFile3 As Integer
  Dim TNumOfRecs As Long, cnt As Long, RCnt As Integer, ToRemove As Long
  Dim TrType As String, TrTyp As Integer, CustBook As Integer
  Dim Removed As Long, FromBook As Integer, ThruBook As Integer
  Dim operchk As Integer
  Date1 = Date2Num(txtDate1.Text)
  FromBook = Val(fptxtRoute1)
  ThruBook = Val(fptxtRoute2)
  If fpcboTransType.ListIndex <> -1 Then
    TrType$ = QPTrim$(Left$(fpcboTransType.Text, 2))
    TrTyp = Val(TrType$)
  Else
    MsgBox "Invalid Transaction Type.", vbOKOnly, "Invalid Selection"
    Exit Sub
  End If
  ToRemove = 0
'this trtyp of 0 would only work if allowed all
'which we do not allow on transaction type - maybe in administrative section
'  If TrTyp = 0 Then
'    BegTrans = 1
'    EndTrans = 999
'  Else
'    BegTrans = TrTyp
'    EndTrans = TrTyp
'  End If
  operchk = Val(txtOperator)
  DeActivateControls frmUtilStripTrans
  FrmShowPctComp.Label1 = "Gathering Transactions to Remove"
  FrmShowPctComp.Show , Me
  UBLog "StripTrans - " & txtDate1.Text & ", book(" & fptxtRoute1 & "-" & fptxtRoute2 & "),Trans-" & TrType$
  ReDim UBCust(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  CustLen = Len(UBCust(1))
  ToRemove = 0

  UBFile1 = FreeFile
  Open "UBTRANS.dat" For Random Shared As UBFile1 Len = UBTranRecLen
  UBFile3 = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile3 Len = CustLen

  TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
  For cnt& = 1 To TNumOfRecs&
    FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
    Get UBFile1, cnt&, UBTranRec(1)
    If Check1.Value = ValueTrue Then
  
    If (UBTranRec(1).TransDate <= Date1) Then   'AND (UBTranRec(1).OperatorNumber = 4) THEN
'      Select Case UBTranRec(1).TransType
'      Case TranUtilityBill
      If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
        If UBTranRec(1).CustAcctNo < 0 Then
          UBLog "InvAcct1:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
          GoTo InvAcct1
        End If
        Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        CustBook = Val(UBCust(1).Book)
        If CustBook >= FromBook And CustBook <= ThruBook Then
         If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
          Select Case TrType
            Case 1  'Utility Bill
              ToRemove = ToRemove + 1
            Case 4   'payment
              ToRemove = ToRemove + 1
            Case 6    'Penalty charge
              ToRemove = ToRemove + 1
            Case 7    'deposit payment
              ToRemove = ToRemove + 1
            Case 11   '"Bill-Upward Adjustment"
              ToRemove = ToRemove + 1
            Case 12   '"Bill-Downward Adjustment"
              ToRemove = ToRemove + 1
            Case 33   '"Payment Adjustment"
              ToRemove = ToRemove + 1
            Case Else
            End Select
          End If
        End If
      End If
    End If
  Else
      If (UBTranRec(1).TransDate = Date1) Then   'AND (UBTranRec(1).OperatorNumber = 4) THEN
'      Select Case UBTranRec(1).TransType
'      Case TranUtilityBill
      If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
        If UBTranRec(1).CustAcctNo < 0 Then
          UBLog "InvAcct1:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
          GoTo InvAcct1
        End If
        Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        CustBook = Val(UBCust(1).Book)
        If CustBook >= FromBook And CustBook <= ThruBook Then
         If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
          Select Case TrType
            Case 1  'Utility Bill
              ToRemove = ToRemove + 1
            Case 4   'payment
              ToRemove = ToRemove + 1
            Case 6    'Penalty charge
              ToRemove = ToRemove + 1
            Case 7    'deposit payment
              ToRemove = ToRemove + 1
            Case 11   '"Bill-Upward Adjustment"
              ToRemove = ToRemove + 1
            Case 12   '"Bill-Downward Adjustment"
              ToRemove = ToRemove + 1
            Case 33   '"Payment Adjustment"
              ToRemove = ToRemove + 1
            Case Else
            End Select
          End If
        End If
      End If
    End If

  End If
InvAcct1:
  Next
  Close
  If ToRemove > 0 Then
    If MsgBox("Num of Trans to be Removed: " & ToRemove & " Yes to Remove, No to Cancel", vbYesNo, "Num to Remove") = vbYes Then
      KillFile "UBTRANS.bak"
      Name "UBTRANS.DAT" As "UBTRANS.bak"
    
      UBFile1 = FreeFile
      Open "UBTRANS.bak" For Random Shared As UBFile1 Len = UBTranRecLen
    
      UBFile2 = FreeFile
      Open "UBTRANS.DAT" For Random Shared As UBFile2 Len = UBTranRecLen
    
      UBFile3 = FreeFile
      Open "UBCUST.DAT" For Random Shared As UBFile3 Len = CustLen
    
      TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
 
      
      Removed = 0
      FrmShowPctComp.Label1 = "Removing Transactions"
      FrmShowPctComp.Show , Me
      For cnt& = 1 To TNumOfRecs&
        FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
        If FrmShowPctComp.Out = True Then
          Close
          FrmShowPctComp.Out = False
          Exit Sub
        End If
        Get UBFile1, cnt&, UBTranRec(1)
        If Check1.Value = ValueTrue Then
        If (UBTranRec(1).TransDate <= Date1) Then   'AND (UBTranRec(1).OperatorNumber = 4) THEN
    '      Select Case UBTranRec(1).TransType
    '      Case TranUtilityBill
          If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
            If UBTranRec(1).CustAcctNo < 0 Then
              UBLog "InvAcct:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
              GoTo InvAcct
            End If
            Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
            CustBook = Val(UBCust(1).Book)
            If CustBook >= FromBook And CustBook <= ThruBook Then
             If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
              Select Case TrType
                Case 1  'Utility Bill
                  Removed = Removed + 1
                Case 4   'payment
                  Removed = Removed + 1
                Case 6    'Penalty charge
                  Removed = Removed + 1
                Case 7    'deposit payment
                  Removed = Removed + 1
                Case 11   '"Bill-Upward Adjustment"
                  Removed = Removed + 1
                Case 12   '"Bill-Downward Adjustment"
                  Removed = Removed + 1
                Case 33   '"Payment Adjustment"
                  Removed = Removed + 1
                Case Else
                  Put UBFile2, , UBTranRec(1)
                End Select
              Else
               Put UBFile2, , UBTranRec(1)
              End If
            Else
             Put UBFile2, , UBTranRec(1)
            End If
          Else
            Put UBFile2, , UBTranRec(1)
          End If
        Else
          Put UBFile2, , UBTranRec(1)
        End If
      Else
          If (UBTranRec(1).TransDate = Date1) Then   'AND (UBTranRec(1).OperatorNumber = 4) THEN
          If (UBTranRec(1).TransType = TrType) Or (UBTranRec(1).TransType = TrType + 100) Then
            If UBTranRec(1).CustAcctNo < 0 Then
              UBLog "InvAcct:" & Str(UBTranRec(1).CustAcctNo) & "," & Str(cnt&) & "," & Str(UBTranRec(1).Transamt)
              GoTo InvAcct
            End If
            Get UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
            CustBook = Val(UBCust(1).Book)
            If CustBook >= FromBook And CustBook <= ThruBook Then
             If operchk = 0 Or operchk = UBTranRec(1).OperatorNumber Then
              Select Case TrType
                Case 1  'Utility Bill
                  Removed = Removed + 1
                Case 4   'payment
                  Removed = Removed + 1
                Case 6    'Penalty charge
                  Removed = Removed + 1
                Case 7    'deposit payment
                  Removed = Removed + 1
                Case 11   '"Bill-Upward Adjustment"
                  Removed = Removed + 1
                Case 12   '"Bill-Downward Adjustment"
                  Removed = Removed + 1
                Case 33   '"Payment Adjustment"
                  Removed = Removed + 1
                Case Else
                  Put UBFile2, , UBTranRec(1)
                End Select
              Else
               Put UBFile2, , UBTranRec(1)
              End If
            Else
             Put UBFile2, , UBTranRec(1)
            End If
          Else
            Put UBFile2, , UBTranRec(1)
          End If
        Else
          Put UBFile2, , UBTranRec(1)
        End If

      End If
InvAcct:
      Next
      Close
      'ActivateControls frmUtilStripTrans
      UBLog "Removed:" & Removed & " Using Strip Trans Util"
      MsgBox "Removed:" & Removed, vbOKOnly, "Removed Trans"
      UBRelinkTransactions
    Else
      Close
      ActivateControls frmUtilStripTrans
      UBLog "None Removed, Canceled by user, Using Strip Trans Util"
    End If
  Else
    Close
    ActivateControls frmUtilStripTrans
    UBLog "No Trans Removed 0 to remove in first pass Using Strip Trans Util"
    MsgBox "No Transactions met criteria to be Removed:" & Removed, vbOKOnly, "Removed 0 Trans"
  End If
End Sub
