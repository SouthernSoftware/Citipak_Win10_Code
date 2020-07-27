VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptSCnsmpTopTen 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Top Ten Consumption Users"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmRptSCnsmpTopTen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   3885
      TabIndex        =   2
      Top             =   1920
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColDesigner     =   "frmRptSCnsmpTopTen.frx":08CA
   End
   Begin LpLib.fpList fplstRates 
      Height          =   1725
      Left            =   6435
      TabIndex        =   3
      Top             =   1485
      Width           =   4935
      _Version        =   196608
      _ExtentX        =   8705
      _ExtentY        =   3043
      TextAlias       =   ""
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmRptSCnsmpTopTen.frx":0C30
   End
   Begin LpLib.fpList fplstGCodes 
      Height          =   915
      Left            =   6990
      TabIndex        =   15
      Top             =   6435
      Width           =   4350
      _Version        =   196608
      _ExtentX        =   7673
      _ExtentY        =   1614
      TextAlias       =   ""
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
      Columns         =   3
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmRptSCnsmpTopTen.frx":0F4C
   End
   Begin VB.CheckBox Detail 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Detail :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2928
      TabIndex        =   4
      Top             =   2355
      Width           =   1164
   End
   Begin VB.OptionButton optAllCust 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All Customers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2325
      TabIndex        =   6
      Top             =   3600
      Width           =   2100
   End
   Begin EditLib.fpText fptxtBook 
      Height          =   780
      Left            =   6990
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4515
      Width           =   4335
      _Version        =   196608
      _ExtentX        =   7641
      _ExtentY        =   1376
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
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
      ButtonStyle     =   0
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   -1  'True
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.OptionButton OptGroup 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1098
      TabIndex        =   11
      Top             =   6435
      Width           =   1260
   End
   Begin VB.OptionButton OptCycle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cycle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1104
      TabIndex        =   9
      Top             =   5490
      Width           =   1260
   End
   Begin VB.OptionButton OptBook 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   1098
      TabIndex        =   7
      Top             =   4515
      Width           =   1260
   End
   Begin EditLib.fpLongInteger fptxtCycleSel 
      Height          =   345
      Left            =   4110
      TabIndex        =   10
      Top             =   5565
      Width           =   630
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
      ButtonMin       =   1
      ButtonMax       =   100
      ButtonStyle     =   0
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CheckBox ActiveOnly 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Active Only:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2448
      TabIndex        =   5
      Top             =   2850
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
      Left            =   8568
      TabIndex        =   12
      Top             =   7560
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
      Left            =   10176
      TabIndex        =   13
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   8280
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
            TextSave        =   "3:39 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10/12/2018"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   345
      Left            =   3885
      TabIndex        =   1
      Top             =   1515
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   345
      Left            =   3885
      TabIndex        =   0
      Top             =   1110
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpLongInteger fptxtBookSel 
      Height          =   345
      Left            =   4110
      TabIndex        =   8
      Top             =   4530
      Width           =   630
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
      ButtonMin       =   1
      ButtonMax       =   100
      ButtonStyle     =   0
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtcycle 
      Height          =   705
      Left            =   6990
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5490
      Width           =   4335
      _Version        =   196608
      _ExtentX        =   7641
      _ExtentY        =   1249
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
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
      ButtonStyle     =   0
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
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      OnFocusPosition =   1
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   -1  'True
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Press 0 and Enter to Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   2805
      TabIndex        =   32
      Top             =   5925
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Press 0 and Enter to Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   2790
      TabIndex        =   31
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
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
      Index           =   4
      Left            =   765
      TabIndex        =   30
      Top             =   3630
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1035
      X2              =   11511
      Y1              =   4395
      Y2              =   4407
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1035
      X2              =   11523
      Y1              =   6330
      Y2              =   6330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1005
      X2              =   11493
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3990
      Left            =   555
      Top             =   3510
      Width           =   11010
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Cycles:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   345
      Index           =   2
      Left            =   4515
      TabIndex        =   29
      Top             =   5550
      Width           =   2340
   End
   Begin VB.Label lblGrp1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Group Codes From List:"
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
      Height          =   330
      Index           =   8
      Left            =   3405
      TabIndex        =   27
      Top             =   6525
      Width           =   3420
   End
   Begin VB.Label lblGrp2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Press SpaceBar or Mouse to Toggle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   3630
      TabIndex        =   26
      Top             =   6870
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or Select Search Option Below: "
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
      Height          =   390
      Left            =   690
      TabIndex        =   25
      Top             =   4050
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cycle:"
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
      Index           =   3
      Left            =   2565
      TabIndex        =   24
      Top             =   5610
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Book:"
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
      Index           =   1
      Left            =   2565
      TabIndex        =   23
      Top             =   4590
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Books:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   345
      Index           =   0
      Left            =   4890
      TabIndex        =   22
      Top             =   4560
      Width           =   1950
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Rate Codes From List:"
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
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   20
      Top             =   1080
      Width           =   3330
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Height          =   330
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
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
      Height          =   420
      Left            =   2070
      TabIndex        =   18
      Top             =   1155
      Width           =   1665
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2505
      Left            =   555
      Top             =   885
      Width           =   11010
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   1470
      TabIndex        =   17
      Top             =   1950
      Width           =   2340
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   585
      Left            =   3210
      Top             =   195
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Top Ten Consumption Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3210
      TabIndex        =   16
      Top             =   330
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   690
      Left            =   3210
      Top             =   90
      Width           =   5775
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1032
      X2              =   1020
      Y1              =   4395
      Y2              =   7407
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
Attribute VB_Name = "frmRptSCnsmpTopTen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim RateRec As Integer, RCnt As Integer
Dim Doall As Boolean, Allcust As Boolean
Dim CycleCnt As Integer, BookCnt As Integer
Dim Cycle(1 To 30) As Integer
Dim Book(1 To 30) As Integer
Dim CustAcT(1 To 10) As UBConsumpRptType
'Dim CustBCG(1 To 10) As UBDraftPayRecType
Dim CustAcR(1 To 10) As UBConsumpRptType
Dim NumOfcdsrpt As Integer, Codefile As Integer, GCnt As Integer
Dim CodeName As String, Grp As String, RateD As String, RCName As String
Dim NumOfratesrpt As Integer, RCfile As Integer, MoFlag As Boolean
Dim CodestoRpt As GroupCodeRptType
Dim RCstoRpt As RateCodeRptType
Private Sub cmdExit_Click()
  frmUBStatReportsMenu.Show
  Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptSCnsmpExtra by " + PWUser$
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
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
  Dim RCode As String
  Dim Handle As Integer, cnt As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  Me.HelpContextID = hlpTopTen
  ReDim UBRateTblRec(1) As UBRateTblRecType
  RCode$ = Space$(10)
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  Handle = FreeFile
  fplstRates.AddItem "ALL" & Chr$(9) & "-Print All Rates"
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get Handle, cnt, UBRateTblRec(1)
    LSet RCode$ = QPTrim$(UBRateTblRec(1).Ratecode)
    fplstRates.AddItem RCode$ & Chr$(9) & QPTrim$(UBRateTblRec(1).RATEDESC)
  Next
  Close
  fplstRates.ListIndex = 1
  fplstRates.Selected(1) = True
  GCodesList fplstGCodes
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  optAllCust.Value = True
  Erase Cycle()
  Erase Book()
  BookCnt = 0
  CycleCnt = 0
  If InStr(TOWNNAME$, "MOORE") > 0 Or InStr(TOWNNAME$, "JOHNSTON") > 0 Then
    MoFlag = True
  Else
    MoFlag = False
  End If
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

'Private Sub OptBook_Click()
'  If OptBook.Value = True Then
'
'End Sub

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fplstRates.SetFocus
  End If
End Sub


Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

Private Sub cmdPrint_Click()
  Dim RptType As Integer
'RptOpt 1 is Consumption by rate code
'2 is Irrigation Consumption Report
  If fplstRates.SelCount > 0 Then
    RateRec = fplstRates.ListIndex
    If RateRec = 0 Then
      Doall = True
    Else
      Doall = False
      GetRCodestoReport fplstRates
    End If
  Else
    MsgBox "Error with Rate selection, please select a Rate option from the list.", vbOKOnly, "Invalid Selection"
    Exit Sub
  End If
  If optAllCust.Value = True Then
    Allcust = True
  Else
    Allcust = False
  End If
  If OptBook.Value = True Then
    If Not CheckBooks% Then
      MsgBox "Error with Book selection, please enter books again.", vbOKOnly, "Invalid Selection"
      Exit Sub
    End If
  ElseIf OptCycle.Value = True Then
    If Not CheckCycles% Then
      MsgBox "Error with Cycle selection, please enter cycles again.", vbOKOnly, "Invalid Selection"
      Exit Sub
    End If
  ElseIf OptGroup.Value = True Then
    If Not fplstGCodes.SelCount > 0 Then
      MsgBox "Error with Group selection, please select groups again.", vbOKOnly, "Invalid Selection"
      Exit Sub
    Else
      GetGCodestoReport fplstGCodes
    End If
  End If
  If ValidDate Then
'    DeActivateControls Me, True
    RptType = fpcboRptType.ListIndex
        ConsumpUnitStep RptType
 '     ActivateControls Me, True
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
        fplstRates.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fptxtBookSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtBookSel.Text) <> 0 Then
      getbooklist
    Else
     'cmdOk.SetFocus
    End If
  End If
End Sub
Private Sub getbooklist()
  Dim TBok As String
  Dim ThisBook As Integer
  Dim cnt As Integer
 
  TBok$ = QPTrim$(fptxtBookSel.Text)
  If TBok$ = "0" Then
    fptxtBook.Text = ""
    BookCnt = 0
    Erase Book
    'cmdOk.SetFocus
  Else
    If Len(TBok$) > 0 Then
      ThisBook = Val(fptxtBookSel.Text)
      For cnt = 1 To 30
        If ThisBook = Book(cnt) Then
          GoTo DupeExit
        End If
      Next
      BookCnt = BookCnt + 1
      If BookCnt > 30 Then
        BookCnt = 30
        GoTo DupeExit
      End If
      Book(BookCnt) = ThisBook
      fptxtBook.Text = ""
      For cnt = 1 To BookCnt
        If cnt = BookCnt Then
          fptxtBook.Text = fptxtBook.Text & Book(cnt)
        Else
          fptxtBook.Text = fptxtBook.Text & Book(cnt) & ","
        End If
      Next
    End If
  End If
DupeExit:
  fptxtBookSel.Text = ""
End Sub
Private Function CheckBooks%()
  
  Dim BooksOK As Boolean
  BooksOK = False
  For RCnt = 1 To 30
    If Book(RCnt) > 0 Then
      BooksOK = True
      Exit For
    End If
  Next
  
  If Not BooksOK Then 'duh nothing to export
'    frmMsgDialog.RetLabel = "-2"
'    frmMsgDialog.Caption = "ERROR:"
'    For RCnt = 0 To 4
'      frmMsgDialog.Label(RCnt).Caption = ""
'      frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
'    Next
'    frmMsgDialog.Label(1).Caption = "NO CYCLES ENTERED TO EXPORT."
'    frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
'    frmMsgDialog.Label(3).Caption = "additional Information."
'    frmMsgDialog.Show vbModal
'    Unload frmMsgDialog
    GoTo CheckBooksExit
  End If

CheckBooksExit:
  
  CheckBooks% = BooksOK
End Function

Private Sub fptxtCycleSel_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cnt As Integer
  If KeyCode = vbKeyReturn Then
    If Len(fptxtCycleSel.Text) <> 0 Then
      getcyclelist
    Else
     'cmdOk.SetFocus
    End If
  End If
End Sub

Private Sub getcyclelist()
  Dim TCyc As String
  Dim ThisCycle As Integer
  Dim cnt As Integer
 
  TCyc$ = QPTrim$(fptxtCycleSel.Text)
  If TCyc$ = "0" Then
    fptxtcycle.Text = ""
    CycleCnt = 0
    Erase Cycle
    'cmdOk.SetFocus
  Else
    If Len(TCyc$) > 0 Then
      ThisCycle = Val(fptxtCycleSel.Text)
      For cnt = 1 To 30
        If ThisCycle = Cycle(cnt) Then
          GoTo DupeExit
        End If
      Next
      CycleCnt = CycleCnt + 1
      If CycleCnt > 30 Then
        CycleCnt = 30
        GoTo DupeExit
      End If
      Cycle(CycleCnt) = ThisCycle
      fptxtcycle.Text = ""
      For cnt = 1 To CycleCnt
        If cnt = CycleCnt Then
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt)
        Else
          fptxtcycle.Text = fptxtcycle.Text & Cycle(cnt) & ","
        End If
      Next
    End If
  End If
DupeExit:
  fptxtCycleSel.Text = ""
End Sub
Private Function CheckCycles%()
  
  Dim CyclesOK As Boolean
  CyclesOK = False
  For RCnt = 1 To 30
    If Cycle(RCnt) > 0 Then
      CyclesOK = True
      Exit For
    End If
  Next
  
  If Not CyclesOK Then 'duh nothing to export
'    frmMsgDialog.RetLabel = "-2"
'    frmMsgDialog.Caption = "ERROR:"
'    For RCnt = 0 To 4
'      frmMsgDialog.Label(RCnt).Caption = ""
'      frmMsgDialog.Label(RCnt).FontSize = frmMsgDialog.Label(RCnt).FontSize + 2
'    Next
'    frmMsgDialog.Label(1).Caption = "NO CYCLES ENTERED TO EXPORT."
'    frmMsgDialog.Label(2).Caption = "Please call Southern Software for"
'    frmMsgDialog.Label(3).Caption = "additional Information."
'    frmMsgDialog.Show vbModal
'    Unload frmMsgDialog
    GoTo CheckCyclesExit
  End If

CheckCyclesExit:
  
  CheckCycles% = CyclesOK
End Function

Private Sub ConsumpUnitStep(RptType)
  Dim Dash80 As String, IdxName As String, IdxRecLen As Integer
  'Dim MINAMT(1 To 1) As Double
  'Dim RATECODE(1 To 1) As String
  Dim UBCustRecLen As Integer, UBCust As Integer, RCnt As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim Handle As Integer, UBTrans As Integer, NumOfRecs As Long
  Dim UBRateTblRecLen As Integer, NumOfRates As Integer
  Dim UBRpt As Integer, UBSetUp As Integer, ValidCustomer As Integer
  Dim BegDate As Integer, EndDate As Integer, Snt As Integer
  Dim TownLen As Integer, TabStop As Integer, MeterConsp As Long
  Dim UBSetupLen As Integer, RateFile As Integer, MT As Integer
  Dim Greater As Boolean, MaxMeterAmt As Long, RCode As String
  Dim Tnt As Integer, NMinAMT As Double, ToPrint As String
  Dim CustomerRecord As Long, MCnt As Integer, GTMeterConsp As Double
  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
  Dim MTRType As String, MType As String, TMeterConsp As Double
  Dim NonUpdated As Integer, LL As Integer, BigUTotal As Double
  Dim ReportFile As String, MinGT As Double, GBBigUTotal As Double
  Dim GBMinGT As Double, GBGTMeterConsp As Double, BigTotCust As Long
  Dim GBCustTot As Long, Bcnt As Integer, CCnt As Integer, GCnt As Integer
  Dim RCstoRpt As RateCodeRptType
  Dim RCfile As Integer, RptInfo As String, RptInfo2 As String
  Dim RCName As String, NumOfCust As Long, ThisTrans As Long
  Dim PCnt As Integer, AvgR As Double, trnum As Long, Avg4R As Double
  Dim BadCount As Integer, TotCust4Rate As Long
  BadCount = 0
  PageNo = 0
  TotCust4Rate = 0
  MaxLines = 56
  Dash80$ = String$(85, "-")
  NumOfRates = GetNumRateRecs%
  If Not Doall Then
    RCName$ = "Ratecds.LST"
    If Not Exist(RCName$) Then GoTo ExitConsStep
    RCfile = FreeFile
    Open RCName$ For Random As RCfile Len = Len(RCstoRpt)
    NumOfratesrpt = FileSize(RCName$) \ Len(RCstoRpt)
    Close RCfile
    ReDim UBRateTbls(1 To NumOfratesrpt) As UBRateTblRecType
    ReDim Ratecode(1 To NumOfratesrpt) As String
  Else
    NumOfratesrpt = NumOfRates
    ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
    ReDim Ratecode(1 To NumOfRates) As String
  End If
  ReDim CustAc(1 To 10) As UBDraftPayRecType
  ReDim UBSetUpRec(1) As UBSetupRecType

  CodeName$ = "grpcds.LST"
 ' RCName$ = "Ratecds.lst"
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  BegDate = Date2Num%(txtDate1)
  EndDate = Date2Num%(txtDate2)
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize(BookIndexFile)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  Handle = FreeFile
  Open BookIndexFile For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close #Handle
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCust& = LOF(UBCust) / UBCustRecLen
  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfRecs& = LOF(UBTrans) / UBTransRecLen
  ReportFile$ = UBPath$ + "UBBKCNSP.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Rem Report Goes Here
  UBSetUp = FreeFile
  Open UBPath$ + "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp Len = UBSetupreclen
  If LOF(UBSetUp) / UBSetupreclen = 0 Then
    TOWNNAME$ = "Undefined"
  Else
    Get UBSetUp, 1, UBSetUpRec(1)
    TOWNNAME$ = UBSetUpRec(1).UTILNAME
    TownLen = Len(RTrim$(TOWNNAME$))
    TabStop = 40 - (TownLen / 2)
    If TabStop < 1 Then TabStop = 1
  End If
  Close UBSetUp
  If Not Doall Then
    RCfile = FreeFile
    Open RCName$ For Random As RCfile Len = Len(RCstoRpt)
  
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
    For PCnt = 1 To NumOfratesrpt
       Get RCfile, PCnt, RCstoRpt
       RateRec = RCstoRpt.RecordNum
       Get RateFile, RateRec, UBRateTbls(PCnt)
       'Rate2Rpt(PCnt) = RateRec
       RptInfo$ = RptInfo$ + QPTrim$(UBRateTbls(PCnt).Ratecode) + " "
       Ratecode(PCnt) = QPTrim$(UBRateTbls(PCnt).Ratecode)
      
     Next PCnt
  Close RateFile
  Else
    RptInfo$ = "All Rates"
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
       For RateRec = 1 To NumOfratesrpt
        Greater = False
        Get RateFile, RateRec, UBRateTbls(RateRec)
        Ratecode(RateRec) = QPTrim$(UBRateTbls(RateRec).Ratecode)
      Next RateRec
    Close RateFile
  End If
  If OptBook.Value = True Then
    For Bcnt = 1 To BookCnt
      RptInfo2$ = RptInfo2$ + Str(Book(Bcnt)) + " "
    Next Bcnt
    RptInfo2$ = "By Book - " + RptInfo2$
  ElseIf OptCycle.Value = True Then
    For CCnt = 1 To CycleCnt
      RptInfo2$ = RptInfo2$ + Str(Cycle(CCnt)) + " "
   Next
   RptInfo2$ = "By Cycle - " + RptInfo2$
  ElseIf OptGroup.Value = True Then
    Codefile = FreeFile
    Open CodeName$ For Random As Codefile Len = Len(CodestoRpt)
    NumOfcdsrpt = FileSize(CodeName$) \ Len(CodestoRpt)
    For GCnt = 1 To NumOfcdsrpt
      Get Codefile, GCnt, CodestoRpt
      RptInfo2$ = RptInfo2$ + QPTrim$(CodestoRpt.GroupCode) + " "
    Next
    Close Codefile
    RptInfo2$ = "By Group - " + RptInfo2$
  ElseIf optAllCust.Value = True Then
   RptInfo2$ = "All Customers"
  End If
  
  GoSub DoRptHeader
'  If Doall Then
'    For RCnt = 1 To NumOfRates
'     'NMinAMT# = MINAMT(RCnt)
'      RCode$ = RATECODE(RCnt)
'      RateRec = RCnt
'      GoSub DoRateHeader
'      GoSub DoEachRate
'      GoSub DoUnitStepFooter
'    Next
'  Else
    For RCnt = 1 To NumOfratesrpt

     ' NMinAMT# = MINAMT(RCnt)
      RCode$ = Ratecode(RCnt)
      RateRec = RCnt
      GoSub DoRateHeader
      Avg4R = 0
      GTMeterConsp# = 0
      GoSub DoEachRate
      GoSub DoUnitStepFooter
      GoSub DotheDetailTot
    Next
'  End If
  GoSub DoGrandFooter
  Close

  'Erase TblBreak&, TotalConsp#, TotalCust
  Erase CustAcT()
  TotCust4Rate = 0
  Doall = False
  If RptType = 0 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptSCnsmpTopTen
    ARptSConsTopTen.txtDate = Now
    ARptSConsTopTen.txtTown = TOWNNAME$
    ARptSConsTopTen.Title = "TOP TEN CONSUMPTION LIST"
    ARptSConsTopTen.txtDate1 = txtDate1.Text
    ARptSConsTopTen.txtDate2 = txtDate2.Text
    ARptSConsTopTen.FldRptInfo = "(Rates - " + RptInfo$ + ") (Customers - " + RptInfo2$ + ")"
    'ARptSCnsmpRate.totCust = Using("###,###,###,###", GBCustTot)
'    ARptSCnsmpRate.totConsump = Using("###,###,###,###", GBGTMeterConsp#)
'    ARptSCnsmpRate.totUsage = Using(" $ ##,###,###.##", GBBigUTotal#)
'    ARptSCnsmpRate.totMin = Using(" $ ##,###,###.##", GBMinGT#)
'    If PageBrk.Value = 1 Then
'      ARptSCnsmpRate.GetName ReportFile$, True
'    Else
      ARptSConsTopTen.GetName ReportFile$
'    End If
    ARptSConsTopTen.startrpt
  Else
    ViewPrint ReportFile$, "Consumption by RateCode"
  'KillFile "UBBKCNSP.RPT"
  End If
  Exit Sub
DoEachRate:
  FrmShowPctComp.Label1 = "Processing Rate " + RCode$
  FrmShowPctComp.Show
If OptBook.Value = True Then
'for book

  For cnt& = 1 To NumOfCust&
    FrmShowPctComp.ShowPctComp cnt, NumOfCust&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If

    Get UBCust, cnt&, UBCustRec(1)
    trnum = 0
    If ActiveOnly.Value = 1 Then
      If UBCustRec(1).Status <> "A" Then
        GoTo skipemB
      End If
    End If
    For Bcnt = 1 To BookCnt
      If Val(UBCustRec(1).Book) <> Book(Bcnt) Then ' And (UBCustRec(1).Status <> "F") Then
        If Bcnt >= BookCnt Then GoTo skipemB
      Else
        Exit For
      End If
    Next Bcnt
    ThisTrans& = UBCustRec(1).LastTrans
    Do While ThisTrans& > 0
      Get #UBTrans, ThisTrans&, UBTransRec(1)
      If MoFlag Then
        If UBTransRec(1).TransDate < BegDate Then
          BadCount = BadCount + 1
          If BadCount > 3 Then
            Exit Do
          End If
        End If
      End If

      If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
        If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
          'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
          ValidCustomer = 0
          If RptType = 1 Then
            If LineCnt > MaxLines Then
              Print #UBRpt, Chr$(12)
              GoSub DoRptHeader
              GoSub DoRateHeader
            End If
          End If
          'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
          For Snt = 1 To 15
            If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
              MTRType$ = UBCustRec(1).serv(Snt).RMtrType
              Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
              ValidCustomer = 1
              Exit For
            End If
          Next Snt
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrB
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF

            MeterConsp& = 0
            'END IF
SkipThisMtrB:
          Next MCnt
          trnum = trnum + 1
        End If
        End If
       End If
      ThisTrans& = UBTransRec(1).PrevTrans
      Loop
       
      If (TMeterConsp# > 0) And (ValidCustomer = 1) Then
        AvgR = Round(TMeterConsp# / trnum)
        Avg4R = Avg4R + AvgR
        CustomerRecord = cnt&
        TotCust4Rate = TotCust4Rate + 1
        If Detail.Value = 1 Then
          If RptType = 1 Then
            Print #UBRpt, Tab(2); Str(CustomerRecord); Tab(9); Mid$(QPTrim$(UBCustRec(1).CustName), 1, 20); Tab(34); Mid$(QPTrim$(UBCustRec(1).ServAddr), 1, 25); Tab(59); Using("###,###,###,###", TMeterConsp#); Tab(75); Using("###,###,###", AvgR)
            LineCnt = LineCnt + 1
          Else
            Print #UBRpt, RCode$ + "~ ~" + Str(CustomerRecord) + "~" + QPTrim$(UBCustRec(1).CustName) + "~" + QPTrim$(UBCustRec(1).ServAddr) + "~" + Using("###,###,###,###", TMeterConsp#) + "~" + Using("###,###,###", AvgR)
          End If
        End If
        GoSub DotheSums
      End If
      TMeterConsp# = 0
  
skipemB:
  Next
ElseIf OptCycle.Value = True Then
'for cycle
  FrmShowPctComp.Label1 = "Processing Rate " + RCode$
  FrmShowPctComp.Show

  For cnt& = 1 To NumOfCust&
    FrmShowPctComp.ShowPctComp cnt, NumOfCust&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If

  Get UBCust, cnt&, UBCustRec(1)
  trnum = 0
  If ActiveOnly.Value = 1 Then
    If UBCustRec(1).Status <> "A" Then
      GoTo skipemC
    End If
  End If
  For CCnt = 1 To CycleCnt
    If Val(UBCustRec(1).BILLCYCL) <> Cycle(CCnt) Then 'And (UBCustRec(1).Status <> "F") Then
      If CCnt >= CycleCnt Then GoTo skipemC
    Else
      Exit For
    End If
  Next

  ThisTrans& = UBCustRec(1).LastTrans
  Do While ThisTrans& > 0
    Get #UBTrans, ThisTrans&, UBTransRec(1)
    If MoFlag Then
      If UBTransRec(1).TransDate < BegDate Then
        BadCount = BadCount + 1
        If BadCount > 3 Then
          Exit Do
        End If
      End If
    End If
    If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If RptType = 1 Then
        If LineCnt > MaxLines Then
          Print #UBRpt, Chr$(12)
          GoSub DoRptHeader
          GoSub DoRateHeader
        End If
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl

            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MTRType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrC
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF

            MeterConsp& = 0
            'END IF
SkipThisMtrC:
          Next MCnt
          trnum = trnum + 1
        End If
        End If
       End If
      ThisTrans& = UBTransRec(1).PrevTrans
      Loop
       
        If (TMeterConsp# > 0) And (ValidCustomer = 1) Then
         AvgR = Round(TMeterConsp# / trnum)
         Avg4R = Avg4R + AvgR
            CustomerRecord = cnt&
            TotCust4Rate = TotCust4Rate + 1
            If Detail.Value = 1 Then
              If RptType = 1 Then
                Print #UBRpt, Tab(2); Str(CustomerRecord); Tab(9); Mid$(QPTrim$(UBCustRec(1).CustName), 1, 20); Tab(34); Mid$(QPTrim$(UBCustRec(1).ServAddr), 1, 25); Tab(59); Using("###,###,###,###", TMeterConsp#); Tab(75); Using("###,###,###", AvgR)
                LineCnt = LineCnt + 1
              Else
            Print #UBRpt, RCode$ + "~ ~" + Str(CustomerRecord) + "~" + QPTrim$(UBCustRec(1).CustName) + "~" + QPTrim$(UBCustRec(1).ServAddr) + "~" + Using("###,###,###,###", TMeterConsp#) + "~" + Using("###,###,###", AvgR)
              End If
            End If
            GoSub DotheSums
        End If
        TMeterConsp# = 0
skipemC:
  Next
ElseIf OptGroup.Value = True Then
  FrmShowPctComp.Label1 = "Processing Rate " + RCode$
  FrmShowPctComp.Show

'for group
  Codefile = FreeFile
  Open CodeName$ For Random As Codefile Len = Len(CodestoRpt)
  NumOfcdsrpt = FileSize(CodeName$) \ Len(CodestoRpt)

  For cnt& = 1 To NumOfCust&
    FrmShowPctComp.ShowPctComp cnt, NumOfCust&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If

  Get UBCust, cnt&, UBCustRec(1)
  trnum = 0
  If ActiveOnly.Value = 1 Then
    If UBCustRec(1).Status <> "A" Then
      GoTo bskipem
    End If
  End If
  For GCnt = 1 To NumOfcdsrpt
    Get Codefile, GCnt, CodestoRpt
    If UBCustRec(1).GroupCodeRec <> CodestoRpt.RecordNum Then
      If GCnt >= NumOfcdsrpt Then GoTo bskipem
    Else
      Exit For
    End If
  Next
  Grp$ = QPTrim$(CodestoRpt.GroupCode)
  ThisTrans& = UBCustRec(1).LastTrans
  Do While ThisTrans& > 0
    Get #UBTrans, ThisTrans&, UBTransRec(1)
    If MoFlag Then
      If UBTransRec(1).TransDate < BegDate Then
        BadCount = BadCount + 1
        If BadCount > 3 Then
          Exit Do
        End If
      End If
    End If

     If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If RptType = 1 Then
        If LineCnt > MaxLines Then
          Print #UBRpt, Chr$(12)
          GoSub DoRptHeader
          GoSub DoRateHeader
        End If
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
            
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MTRType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtrG
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            ''If MeterConsp& > 0 Then Stop
            'LPRINT CustomerRecord
            'STOP
            'END IF

            MeterConsp& = 0
            'END IF
SkipThisMtrG:

          Next MCnt
          trnum = trnum + 1
        End If
        End If
       End If
      ThisTrans& = UBTransRec(1).PrevTrans
      Loop
       
        If (TMeterConsp# > 0) And (ValidCustomer = 1) Then
         AvgR = Round(TMeterConsp# / trnum)
         Avg4R = Avg4R + AvgR
            CustomerRecord = cnt&
            TotCust4Rate = TotCust4Rate + 1
            If Detail.Value = 1 Then
              If RptType = 1 Then
                Print #UBRpt, Tab(2); Str(CustomerRecord); Tab(9); Mid$(QPTrim$(UBCustRec(1).CustName), 1, 20); Tab(34); Mid$(QPTrim$(UBCustRec(1).ServAddr), 1, 25); Tab(59); Using("###,###,###,###", TMeterConsp#); Tab(75); Using("###,###,###", AvgR)
                LineCnt = LineCnt + 1
              Else
            Print #UBRpt, RCode$ + "~ ~" + Str(CustomerRecord) + "~" + QPTrim$(UBCustRec(1).CustName) + "~" + QPTrim$(UBCustRec(1).ServAddr) + "~" + Using("###,###,###,###", TMeterConsp#) + "~" + Using("###,###,###", AvgR)
              End If
            End If
            GoSub DotheSums
        End If
        TMeterConsp# = 0
  
bskipem:
  Next
ElseIf optAllCust.Value = True Then
  FrmShowPctComp.Label1 = "Processing Rate " + RCode$
  FrmShowPctComp.Show

  For cnt& = 1 To NumOfCust&
    FrmShowPctComp.ShowPctComp cnt, NumOfCust&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If

    Get UBCust, cnt&, UBCustRec(1)
  '  If cnt& = 154 Then Stop
    trnum = 0
    If ActiveOnly.Value = 1 Then
      If UBCustRec(1).Status <> "A" Then
        GoTo skiptonextcust
      End If
    End If
    ThisTrans& = UBCustRec(1).LastTrans
    Do While ThisTrans& > 0
      Get #UBTrans, ThisTrans&, UBTransRec(1)
      If MoFlag Then
        If UBTransRec(1).TransDate < BegDate Then
          BadCount = BadCount + 1
          If BadCount > 3 Then
            Exit Do
          End If
        End If
      End If

      If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
        If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
          ValidCustomer = 0
          If RptType = 1 Then
            If LineCnt > MaxLines Then
              Print #UBRpt, Chr$(12)
              GoSub DoRptHeader
              GoSub DoRateHeader
            End If
          End If
          For Snt = 1 To 15
            If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
              MTRType$ = UBCustRec(1).serv(Snt).RMtrType
              Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
                Case Else
                  MT = 4
                End Select
              ValidCustomer = 1
              Exit For
            End If
          Next Snt
          'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
          If ValidCustomer = 1 Then
            Multi& = 0
            Cubic = False
            For MCnt = 1 To 7
              MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
              If MTRType$ = MType$ Then
              'IF LEN(MType$) > 0 THEN
                Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
                If Multi& = 0 Then
                  Multi& = 1
                End If
                If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
                  Cubic = True
                End If
                Exit For
              End If
            Next
            If Multi& <= 0 Then Multi& = 1
            'IF WhatRev > 0 THEN
              ChkMtr = True
            'ELSE
            '  ChkMtr = False
            'END IF
            For MCnt = 1 To 7
              If ChkMtr = True Then
                If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                  GoTo SkipThisMtrA
                End If
              End If
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
              If MeterConsp& < 0 Then
                MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
                MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
              End If
              If Cubic Then
                MeterConsp& = MeterConsp& * 7.481
              End If
              MeterConsp& = MeterConsp& * Multi&
              'IF MeterConsp& = 1 THEN STOP
              TMeterConsp# = TMeterConsp# + MeterConsp&
              GTMeterConsp# = GTMeterConsp# + MeterConsp&
              ''If MeterConsp& > 0 Then Stop
              'LPRINT CustomerRecord
              'STOP
              'END IF
  
              MeterConsp& = 0
              'END IF
SkipThisMtrA:
            Next MCnt
            trnum = trnum + 1
          End If
        End If
      End If
     ThisTrans& = UBTransRec(1).PrevTrans
     Loop
    
     If (TMeterConsp# > 0) And (ValidCustomer = 1) Then
      AvgR = Round(TMeterConsp# / trnum)
      Avg4R = Avg4R + AvgR
        TotCust4Rate = TotCust4Rate + 1
         CustomerRecord = cnt&
         If Detail.Value = 1 Then
          If RptType = 1 Then
            Print #UBRpt, Tab(2); Str(CustomerRecord); Tab(9); Mid$(QPTrim$(UBCustRec(1).CustName), 1, 20); Tab(34); Mid$(QPTrim$(UBCustRec(1).ServAddr), 1, 25); Tab(59); Using("###,###,###,###", TMeterConsp#); Tab(75); Using("###,###,###", AvgR)
            LineCnt = LineCnt + 1
          Else
            Print #UBRpt, RCode$ + "~ ~" + Str(CustomerRecord) + "~" + QPTrim$(UBCustRec(1).CustName) + "~" + QPTrim$(UBCustRec(1).ServAddr) + "~" + Using("###,###,###,###", TMeterConsp#) + "~" + Using("###,###,###", AvgR)
          End If
         End If
         GoSub DotheSums
     End If
     TMeterConsp# = 0
skiptonextcust:
  Next
 End If
Return
DoRptHeader:
 If RptType = 1 Then
  PageNo = PageNo + 1
  Print #UBRpt, Tab(29); "Top Ten Consumption by RateCode"; Tab(70); "Page #"; PageNo
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, "Report Date: "; Now
  Print #UBRpt, "Period Beginning: "; txtDate1; Tab(35); "    Period Ending: "; txtDate2
  Print #UBRpt, "Report Opt:  "; "(Rates - "; RptInfo$; ") (Customers - "; RptInfo2$; ")"
'  Print #UBRpt, "    For Rate Code: "; fpComboRates.Text
'  Print #UBRpt, " Period Beginning: "; txtDate1
'  Print #UBRpt, "    Period Ending: "; txtDate2
  Print #UBRpt, " "
  Print #UBRpt, Dash80$
  LineCnt = 5
 End If
Return
DoRateHeader:
  If RptType = 1 Then
  'PageNo = PageNo + 1
'  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, TOWNNAME$
'  Print #UBRpt, "Report Date: "; Now
  Print #UBRpt, " "
  Print #UBRpt, "    For Rate Code: "; RCode$
  If Detail.Value = 1 Then
    Print #UBRpt, ; Tab(2); "Account"; Tab(11); "Name"; Tab(35); "Service Address"; Tab(67); "Consmpt"; Tab(78); "Avg Cnsp"
  End If
  Print #UBRpt, Dash80$
  LineCnt = LineCnt + 6
 End If
Return
DotheDetailTot:
  If RptType = 1 Then
    Print #UBRpt, " Rate Totals : "; Tab(20); Using("###,###,###,###", TotCust4Rate); Tab(59); Using("###,###,###,###", GTMeterConsp#); Tab(75); Using("###,###,###", Avg4R)
    Print #UBRpt, Dash80$
    LineCnt = LineCnt + 2
  Else
    Print #UBRpt, RCode$ + "~ ~ ~ ~ ~ ~ "
    Print #UBRpt, RCode$ + "~ ~** ~" + "RATE TOTALS:" + "~" + "Customers - " + Using("###,###,###,###", TotCust4Rate) + "~" + Using("###,###,###,###", GTMeterConsp#) + "~" + Using("###,###,###", Avg4R)
    Print #UBRpt, RCode$ + "~ ~ ~ ~ ~ ~ "
  End If
  TotCust4Rate = 0
  GTMeterConsp# = 0
Return
DoUnitStepFooter:
If GTMeterConsp# > 0 Then
 ' TblBreak&(RateRec, MaxStep(RateRec)) = 99999999
  If RptType = 1 Then
    Print #UBRpt, "Top Ten List"
    Print #UBRpt, ; Tab(7); "Acct"; Tab(12); "Name"; Tab(35); "Service Address"; Tab(67); "Consmpt"; Tab(78); "Avg Cnsp"
    For LL = 1 To 10
      Print #UBRpt, LL;
      Print #UBRpt, Tab(5); Str(CustAcR(LL).CustAcct); Tab(12); Mid$(QPTrim$(CustAcR(LL).CustName), 1, 20); Tab(33); Mid$(QPTrim$(CustAcR(LL).SvcAddr), 1, 25); Tab(59); Using("###,###,###,###", CustAcR(LL).ConsumpAmt); Tab(75); Using("###,###,###", CustAcR(LL).AvgAmt)
      'Print #UBRpt, ""
  '    Print #UBRpt, Dash80$
      LineCnt = LineCnt + 1
    Next LL
    LineCnt = LineCnt + 1
  '  If PageBrk = 1 Then
      Print #UBRpt, Chr$(12);
  '  Else
      Print #UBRpt,
      LineCnt = LineCnt + 1
  '  End If
    BigUTotal# = 0
    MinGT# = 0
    BigTotCust = 0
    Erase CustAcR()
    
  ElseIf RptType = 0 Then
    'Print #UBRpt, RCode$ + "~ ~ ~ ~ ~ ~ "
    'Print #UBRpt, RCode$ + "~ ~ ~ ~ ~ ~ "
    Print #UBRpt, RCode$ + "~ ~ ~" + "TOP TEN LIST**" + "~ ~ ~"
    ''Print #UBRpt, RCode$ + "~ ~Acct~Name~Svc Addr~Consump~Average"
    For LL = 1 To 10
      Print #UBRpt, RCode$ + "~" + Str(LL) + "~" + Str(CustAcR(LL).CustAcct) + "~" + QPTrim$(CustAcR(LL).CustName) + "~" + QPTrim$(CustAcR(LL).SvcAddr) + "~" + Using("###,###,###,###", CustAcR(LL).ConsumpAmt) + "~" + Using("###,###,###", CustAcR(LL).AvgAmt)
    Next LL
    BigUTotal# = 0
    MinGT# = 0
    BigTotCust = 0
    Erase CustAcR()
    
  End If
End If
Return
DoGrandFooter:
GoSub DoRptHeader
If RptType = 1 Then
  Print #UBRpt, ; Tab(7); "Acct"; Tab(12); "Name"; Tab(35); "Service Address"; Tab(67); "Consmpt"; Tab(78); "Avg Cnsp"
  For LL = 1 To 10
    Print #UBRpt, LL;
    Print #UBRpt, Tab(5); Str(CustAcT(LL).CustAcct); Tab(12); Mid$(QPTrim$(CustAcT(LL).CustName), 1, 20); Tab(33); Mid$(QPTrim$(CustAcT(LL).SvcAddr), 1, 25); Tab(59); Using("###,###,###,###", CustAcT(LL).ConsumpAmt); Tab(75); Using("###,###,###", CustAcT(LL).AvgAmt)
    'Print #UBRpt, ""
'    Print #UBRpt, Dash80$
    LineCnt = LineCnt + 1
  Next LL
  LineCnt = LineCnt + 1
'  If PageBrk = 1 Then
    Print #UBRpt, Chr$(12);
'  Else
    Print #UBRpt,
    LineCnt = LineCnt + 1
'  End If
  BigUTotal# = 0
  MinGT# = 0
  BigTotCust = 0
  Erase CustAcT()
  GTMeterConsp# = 0

ElseIf RptType = 0 Then
  For LL = 1 To 10
    Print #UBRpt, "Total" + "~" + Str(LL) + "~" + Str(CustAcT(LL).CustAcct) + "~" + QPTrim$(CustAcT(LL).CustName) + "~" + QPTrim$(CustAcT(LL).SvcAddr) + "~" + Using("###,###,###,###", CustAcT(LL).ConsumpAmt) + "~" + Using("###,###,###", CustAcT(LL).AvgAmt)
  Next LL
  BigUTotal# = 0
  MinGT# = 0
  BigTotCust = 0
  Erase CustAcT()
  GTMeterConsp# = 0
End If
Return
DotheSums:
            If TMeterConsp# > CustAcR(1).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = CustAcR(6).ConsumpAmt
              CustAcR(7).CustAcct = CustAcR(6).CustAcct
              CustAcR(7).CustName = CustAcR(6).CustName
              CustAcR(7).SvcAddr = CustAcR(6).SvcAddr
              CustAcR(7).AvgAmt = CustAcR(6).AvgAmt
              CustAcR(6).ConsumpAmt = CustAcR(5).ConsumpAmt
              CustAcR(6).CustAcct = CustAcR(5).CustAcct
              CustAcR(6).CustName = CustAcR(5).CustName
              CustAcR(6).SvcAddr = CustAcR(5).SvcAddr
              CustAcR(6).AvgAmt = CustAcR(5).AvgAmt
              CustAcR(5).ConsumpAmt = CustAcR(4).ConsumpAmt
              CustAcR(5).CustAcct = CustAcR(4).CustAcct
              CustAcR(5).CustName = CustAcR(4).CustName
              CustAcR(5).SvcAddr = CustAcR(4).SvcAddr
              CustAcR(5).AvgAmt = CustAcR(4).AvgAmt
              CustAcR(4).ConsumpAmt = CustAcR(3).ConsumpAmt
              CustAcR(4).CustAcct = CustAcR(3).CustAcct
              CustAcR(4).CustName = CustAcR(3).CustName
              CustAcR(4).SvcAddr = CustAcR(3).SvcAddr
              CustAcR(4).AvgAmt = CustAcR(3).AvgAmt
              CustAcR(3).ConsumpAmt = CustAcR(2).ConsumpAmt
              CustAcR(3).CustAcct = CustAcR(2).CustAcct
              CustAcR(3).CustName = CustAcR(2).CustName
              CustAcR(3).SvcAddr = CustAcR(2).SvcAddr
              CustAcR(3).AvgAmt = CustAcR(2).AvgAmt
              CustAcR(2).ConsumpAmt = CustAcR(1).ConsumpAmt
              CustAcR(2).CustAcct = CustAcR(1).CustAcct
              CustAcR(2).CustName = CustAcR(1).CustName
              CustAcR(2).SvcAddr = CustAcR(1).SvcAddr
              CustAcR(2).AvgAmt = CustAcR(1).AvgAmt
              CustAcR(1).ConsumpAmt = TMeterConsp#
              CustAcR(1).CustAcct = CustomerRecord
              CustAcR(1).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(1).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(1).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(1).ConsumpAmt Or TMeterConsp# > CustAcR(2).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = CustAcR(6).ConsumpAmt
              CustAcR(7).CustAcct = CustAcR(6).CustAcct
              CustAcR(7).CustName = CustAcR(6).CustName
              CustAcR(7).SvcAddr = CustAcR(6).SvcAddr
              CustAcR(7).AvgAmt = CustAcR(6).AvgAmt
              CustAcR(6).ConsumpAmt = CustAcR(5).ConsumpAmt
              CustAcR(6).CustAcct = CustAcR(5).CustAcct
              CustAcR(6).CustName = CustAcR(5).CustName
              CustAcR(6).SvcAddr = CustAcR(5).SvcAddr
              CustAcR(6).AvgAmt = CustAcR(5).AvgAmt
              CustAcR(5).ConsumpAmt = CustAcR(4).ConsumpAmt
              CustAcR(5).CustAcct = CustAcR(4).CustAcct
              CustAcR(5).CustName = CustAcR(4).CustName
              CustAcR(5).SvcAddr = CustAcR(4).SvcAddr
              CustAcR(5).AvgAmt = CustAcR(4).AvgAmt
              CustAcR(4).ConsumpAmt = CustAcR(3).ConsumpAmt
              CustAcR(4).CustAcct = CustAcR(3).CustAcct
              CustAcR(4).CustName = CustAcR(3).CustName
              CustAcR(4).SvcAddr = CustAcR(3).SvcAddr
              CustAcR(4).AvgAmt = CustAcR(3).AvgAmt
              CustAcR(3).ConsumpAmt = CustAcR(2).ConsumpAmt
              CustAcR(3).CustAcct = CustAcR(2).CustAcct
              CustAcR(3).CustName = CustAcR(2).CustName
              CustAcR(3).SvcAddr = CustAcR(2).SvcAddr
              CustAcR(3).AvgAmt = CustAcR(2).AvgAmt
              CustAcR(2).ConsumpAmt = TMeterConsp#
              CustAcR(2).CustAcct = CustomerRecord
              CustAcR(2).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(2).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(2).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(2).ConsumpAmt Or TMeterConsp# > CustAcR(3).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = CustAcR(6).ConsumpAmt
              CustAcR(7).CustAcct = CustAcR(6).CustAcct
              CustAcR(7).CustName = CustAcR(6).CustName
              CustAcR(7).SvcAddr = CustAcR(6).SvcAddr
              CustAcR(7).AvgAmt = CustAcR(6).AvgAmt
              CustAcR(6).ConsumpAmt = CustAcR(5).ConsumpAmt
              CustAcR(6).CustAcct = CustAcR(5).CustAcct
              CustAcR(6).CustName = CustAcR(5).CustName
              CustAcR(6).SvcAddr = CustAcR(5).SvcAddr
              CustAcR(6).AvgAmt = CustAcR(5).AvgAmt
              CustAcR(5).ConsumpAmt = CustAcR(4).ConsumpAmt
              CustAcR(5).CustAcct = CustAcR(4).CustAcct
              CustAcR(5).CustName = CustAcR(4).CustName
              CustAcR(5).SvcAddr = CustAcR(4).SvcAddr
              CustAcR(5).AvgAmt = CustAcR(4).AvgAmt
              CustAcR(4).ConsumpAmt = CustAcR(3).ConsumpAmt
              CustAcR(4).CustAcct = CustAcR(3).CustAcct
              CustAcR(4).CustName = CustAcR(3).CustName
              CustAcR(4).SvcAddr = CustAcR(3).SvcAddr
              CustAcR(4).AvgAmt = CustAcR(3).AvgAmt
              CustAcR(3).ConsumpAmt = TMeterConsp#
              CustAcR(3).CustAcct = CustomerRecord
              CustAcR(3).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(3).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(3).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(3).ConsumpAmt Or TMeterConsp# > CustAcR(4).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = CustAcR(6).ConsumpAmt
              CustAcR(7).CustAcct = CustAcR(6).CustAcct
              CustAcR(7).CustName = CustAcR(6).CustName
              CustAcR(7).SvcAddr = CustAcR(6).SvcAddr
              CustAcR(7).AvgAmt = CustAcR(6).AvgAmt
              CustAcR(6).ConsumpAmt = CustAcR(5).ConsumpAmt
              CustAcR(6).CustAcct = CustAcR(5).CustAcct
              CustAcR(6).CustName = CustAcR(5).CustName
              CustAcR(6).SvcAddr = CustAcR(5).SvcAddr
              CustAcR(6).AvgAmt = CustAcR(5).AvgAmt
              CustAcR(5).ConsumpAmt = CustAcR(4).ConsumpAmt
              CustAcR(5).CustAcct = CustAcR(4).CustAcct
              CustAcR(5).CustName = CustAcR(4).CustName
              CustAcR(5).SvcAddr = CustAcR(4).SvcAddr
              CustAcR(5).AvgAmt = CustAcR(4).AvgAmt
              CustAcR(4).ConsumpAmt = TMeterConsp#
              CustAcR(4).CustAcct = CustomerRecord
              CustAcR(4).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(4).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(4).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(4).ConsumpAmt Or TMeterConsp# > CustAcR(5).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = CustAcR(6).ConsumpAmt
              CustAcR(7).CustAcct = CustAcR(6).CustAcct
              CustAcR(7).CustName = CustAcR(6).CustName
              CustAcR(7).SvcAddr = CustAcR(6).SvcAddr
              CustAcR(7).AvgAmt = CustAcR(6).AvgAmt
              CustAcR(6).ConsumpAmt = CustAcR(5).ConsumpAmt
              CustAcR(6).CustAcct = CustAcR(5).CustAcct
              CustAcR(6).CustName = CustAcR(5).CustName
              CustAcR(6).SvcAddr = CustAcR(5).SvcAddr
              CustAcR(6).AvgAmt = CustAcR(5).AvgAmt
              CustAcR(5).ConsumpAmt = TMeterConsp#
              CustAcR(5).CustAcct = CustomerRecord
              CustAcR(5).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(5).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(5).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(5).ConsumpAmt Or TMeterConsp# > CustAcR(6).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = CustAcR(6).ConsumpAmt
              CustAcR(7).CustAcct = CustAcR(6).CustAcct
              CustAcR(7).CustName = CustAcR(6).CustName
              CustAcR(7).SvcAddr = CustAcR(6).SvcAddr
              CustAcR(7).AvgAmt = CustAcR(6).AvgAmt
              CustAcR(6).ConsumpAmt = TMeterConsp#
              CustAcR(6).CustAcct = CustomerRecord
              CustAcR(6).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(6).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(6).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(6).ConsumpAmt Or TMeterConsp# > CustAcR(7).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = CustAcR(7).ConsumpAmt
              CustAcR(8).CustAcct = CustAcR(7).CustAcct
              CustAcR(8).CustName = CustAcR(7).CustName
              CustAcR(8).SvcAddr = CustAcR(7).SvcAddr
              CustAcR(8).AvgAmt = CustAcR(7).AvgAmt
              CustAcR(7).ConsumpAmt = TMeterConsp#
              CustAcR(7).CustAcct = CustomerRecord
              CustAcR(7).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(7).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(7).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(7).ConsumpAmt Or TMeterConsp# > CustAcR(8).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = CustAcR(8).ConsumpAmt
              CustAcR(9).CustAcct = CustAcR(8).CustAcct
              CustAcR(9).CustName = CustAcR(8).CustName
              CustAcR(9).SvcAddr = CustAcR(8).SvcAddr
              CustAcR(9).AvgAmt = CustAcR(8).AvgAmt
              CustAcR(8).ConsumpAmt = TMeterConsp#
              CustAcR(8).CustAcct = CustomerRecord
              CustAcR(8).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(8).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(8).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(8).ConsumpAmt Or TMeterConsp# > CustAcR(9).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = CustAcR(9).ConsumpAmt
              CustAcR(10).CustAcct = CustAcR(9).CustAcct
              CustAcR(10).CustName = CustAcR(9).CustName
              CustAcR(10).SvcAddr = CustAcR(9).SvcAddr
              CustAcR(10).AvgAmt = CustAcR(9).AvgAmt
              CustAcR(9).ConsumpAmt = TMeterConsp#
              CustAcR(9).CustAcct = CustomerRecord
              CustAcR(9).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(9).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(9).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcR(9).ConsumpAmt Or TMeterConsp# > CustAcR(10).ConsumpAmt Then
              CustAcR(10).ConsumpAmt = TMeterConsp#
              CustAcR(10).CustAcct = CustomerRecord
              CustAcR(10).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcR(10).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcR(10).AvgAmt = AvgR
            End If
            
            If TMeterConsp# > CustAcT(1).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = CustAcT(6).ConsumpAmt
              CustAcT(7).CustAcct = CustAcT(6).CustAcct
              CustAcT(7).CustName = CustAcT(6).CustName
              CustAcT(7).SvcAddr = CustAcT(6).SvcAddr
              CustAcT(7).AvgAmt = CustAcT(6).AvgAmt
              CustAcT(6).ConsumpAmt = CustAcT(5).ConsumpAmt
              CustAcT(6).CustAcct = CustAcT(5).CustAcct
              CustAcT(6).CustName = CustAcT(5).CustName
              CustAcT(6).SvcAddr = CustAcT(5).SvcAddr
              CustAcT(6).AvgAmt = CustAcT(5).AvgAmt
              CustAcT(5).ConsumpAmt = CustAcT(4).ConsumpAmt
              CustAcT(5).CustAcct = CustAcT(4).CustAcct
              CustAcT(5).CustName = CustAcT(4).CustName
              CustAcT(5).SvcAddr = CustAcT(4).SvcAddr
              CustAcT(5).AvgAmt = CustAcT(4).AvgAmt
              CustAcT(4).ConsumpAmt = CustAcT(3).ConsumpAmt
              CustAcT(4).CustAcct = CustAcT(3).CustAcct
              CustAcT(4).CustName = CustAcT(3).CustName
              CustAcT(4).SvcAddr = CustAcT(3).SvcAddr
              CustAcT(4).AvgAmt = CustAcT(3).AvgAmt
              CustAcT(3).ConsumpAmt = CustAcT(2).ConsumpAmt
              CustAcT(3).CustAcct = CustAcT(2).CustAcct
              CustAcT(3).CustName = CustAcT(2).CustName
              CustAcT(3).SvcAddr = CustAcT(2).SvcAddr
              CustAcT(3).AvgAmt = CustAcT(2).AvgAmt
              CustAcT(2).ConsumpAmt = CustAcT(1).ConsumpAmt
              CustAcT(2).CustAcct = CustAcT(1).CustAcct
              CustAcT(2).CustName = CustAcT(1).CustName
              CustAcT(2).SvcAddr = CustAcT(1).SvcAddr
              CustAcT(2).AvgAmt = CustAcT(1).AvgAmt
              CustAcT(1).ConsumpAmt = TMeterConsp#
              CustAcT(1).CustAcct = CustomerRecord
              CustAcT(1).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(1).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(1).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(1).ConsumpAmt Or TMeterConsp# > CustAcT(2).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = CustAcT(6).ConsumpAmt
              CustAcT(7).CustAcct = CustAcT(6).CustAcct
              CustAcT(7).CustName = CustAcT(6).CustName
              CustAcT(7).SvcAddr = CustAcT(6).SvcAddr
              CustAcT(7).AvgAmt = CustAcT(6).AvgAmt
              CustAcT(6).ConsumpAmt = CustAcT(5).ConsumpAmt
              CustAcT(6).CustAcct = CustAcT(5).CustAcct
              CustAcT(6).CustName = CustAcT(5).CustName
              CustAcT(6).SvcAddr = CustAcT(5).SvcAddr
              CustAcT(6).AvgAmt = CustAcT(5).AvgAmt
              CustAcT(5).ConsumpAmt = CustAcT(4).ConsumpAmt
              CustAcT(5).CustAcct = CustAcT(4).CustAcct
              CustAcT(5).CustName = CustAcT(4).CustName
              CustAcT(5).SvcAddr = CustAcT(4).SvcAddr
              CustAcT(5).AvgAmt = CustAcT(4).AvgAmt
              CustAcT(4).ConsumpAmt = CustAcT(3).ConsumpAmt
              CustAcT(4).CustAcct = CustAcT(3).CustAcct
              CustAcT(4).CustName = CustAcT(3).CustName
              CustAcT(4).SvcAddr = CustAcT(3).SvcAddr
              CustAcT(4).AvgAmt = CustAcT(3).AvgAmt
              CustAcT(3).ConsumpAmt = CustAcT(2).ConsumpAmt
              CustAcT(3).CustAcct = CustAcT(2).CustAcct
              CustAcT(3).CustName = CustAcT(2).CustName
              CustAcT(3).SvcAddr = CustAcT(2).SvcAddr
              CustAcT(3).AvgAmt = CustAcT(2).AvgAmt
              CustAcT(2).ConsumpAmt = TMeterConsp#
              CustAcT(2).CustAcct = CustomerRecord
              CustAcT(2).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(2).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(2).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(2).ConsumpAmt Or TMeterConsp# > CustAcT(3).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = CustAcT(6).ConsumpAmt
              CustAcT(7).CustAcct = CustAcT(6).CustAcct
              CustAcT(7).CustName = CustAcT(6).CustName
              CustAcT(7).SvcAddr = CustAcT(6).SvcAddr
              CustAcT(7).AvgAmt = CustAcT(6).AvgAmt
              CustAcT(6).ConsumpAmt = CustAcT(5).ConsumpAmt
              CustAcT(6).CustAcct = CustAcT(5).CustAcct
              CustAcT(6).CustName = CustAcT(5).CustName
              CustAcT(6).SvcAddr = CustAcT(5).SvcAddr
              CustAcT(6).AvgAmt = CustAcT(5).AvgAmt
              CustAcT(5).ConsumpAmt = CustAcT(4).ConsumpAmt
              CustAcT(5).CustAcct = CustAcT(4).CustAcct
              CustAcT(5).CustName = CustAcT(4).CustName
              CustAcT(5).SvcAddr = CustAcT(4).SvcAddr
              CustAcT(5).AvgAmt = CustAcT(4).AvgAmt
              CustAcT(4).ConsumpAmt = CustAcT(3).ConsumpAmt
              CustAcT(4).CustAcct = CustAcT(3).CustAcct
              CustAcT(4).CustName = CustAcT(3).CustName
              CustAcT(4).SvcAddr = CustAcT(3).SvcAddr
              CustAcT(4).AvgAmt = CustAcT(3).AvgAmt
              CustAcT(3).ConsumpAmt = TMeterConsp#
              CustAcT(3).CustAcct = CustomerRecord
              CustAcT(3).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(3).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(3).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(3).ConsumpAmt Or TMeterConsp# > CustAcT(4).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = CustAcT(6).ConsumpAmt
              CustAcT(7).CustAcct = CustAcT(6).CustAcct
              CustAcT(7).CustName = CustAcT(6).CustName
              CustAcT(7).SvcAddr = CustAcT(6).SvcAddr
              CustAcT(7).AvgAmt = CustAcT(6).AvgAmt
              CustAcT(6).ConsumpAmt = CustAcT(5).ConsumpAmt
              CustAcT(6).CustAcct = CustAcT(5).CustAcct
              CustAcT(6).CustName = CustAcT(5).CustName
              CustAcT(6).SvcAddr = CustAcT(5).SvcAddr
              CustAcT(6).AvgAmt = CustAcT(5).AvgAmt
              CustAcT(5).ConsumpAmt = CustAcT(4).ConsumpAmt
              CustAcT(5).CustAcct = CustAcT(4).CustAcct
              CustAcT(5).CustName = CustAcT(4).CustName
              CustAcT(5).SvcAddr = CustAcT(4).SvcAddr
              CustAcT(5).AvgAmt = CustAcT(4).AvgAmt
              CustAcT(4).ConsumpAmt = TMeterConsp#
              CustAcT(4).CustAcct = CustomerRecord
              CustAcT(4).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(4).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(4).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(4).ConsumpAmt Or TMeterConsp# > CustAcT(5).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = CustAcT(6).ConsumpAmt
              CustAcT(7).CustAcct = CustAcT(6).CustAcct
              CustAcT(7).CustName = CustAcT(6).CustName
              CustAcT(7).SvcAddr = CustAcT(6).SvcAddr
              CustAcT(7).AvgAmt = CustAcT(6).AvgAmt
              CustAcT(6).ConsumpAmt = CustAcT(5).ConsumpAmt
              CustAcT(6).CustAcct = CustAcT(5).CustAcct
              CustAcT(6).CustName = CustAcT(5).CustName
              CustAcT(6).SvcAddr = CustAcT(5).SvcAddr
              CustAcT(6).AvgAmt = CustAcT(5).AvgAmt
              CustAcT(5).ConsumpAmt = TMeterConsp#
              CustAcT(5).CustAcct = CustomerRecord
              CustAcT(5).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(5).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(5).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(5).ConsumpAmt Or TMeterConsp# > CustAcT(6).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = CustAcT(6).ConsumpAmt
              CustAcT(7).CustAcct = CustAcT(6).CustAcct
              CustAcT(7).CustName = CustAcT(6).CustName
              CustAcT(7).SvcAddr = CustAcT(6).SvcAddr
              CustAcT(7).AvgAmt = CustAcT(6).AvgAmt
              CustAcT(6).ConsumpAmt = TMeterConsp#
              CustAcT(6).CustAcct = CustomerRecord
              CustAcT(6).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(6).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(6).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(6).ConsumpAmt Or TMeterConsp# > CustAcT(7).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = CustAcT(7).ConsumpAmt
              CustAcT(8).CustAcct = CustAcT(7).CustAcct
              CustAcT(8).CustName = CustAcT(7).CustName
              CustAcT(8).SvcAddr = CustAcT(7).SvcAddr
              CustAcT(8).AvgAmt = CustAcT(7).AvgAmt
              CustAcT(7).ConsumpAmt = TMeterConsp#
              CustAcT(7).CustAcct = CustomerRecord
              CustAcT(7).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(7).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(7).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(7).ConsumpAmt Or TMeterConsp# > CustAcT(8).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = CustAcT(8).ConsumpAmt
              CustAcT(9).CustAcct = CustAcT(8).CustAcct
              CustAcT(9).CustName = CustAcT(8).CustName
              CustAcT(9).SvcAddr = CustAcT(8).SvcAddr
              CustAcT(9).AvgAmt = CustAcT(8).AvgAmt
              CustAcT(8).ConsumpAmt = TMeterConsp#
              CustAcT(8).CustAcct = CustomerRecord
              CustAcT(8).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(8).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(8).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(8).ConsumpAmt Or TMeterConsp# > CustAcT(9).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = CustAcT(9).ConsumpAmt
              CustAcT(10).CustAcct = CustAcT(9).CustAcct
              CustAcT(10).CustName = CustAcT(9).CustName
              CustAcT(10).SvcAddr = CustAcT(9).SvcAddr
              CustAcT(10).AvgAmt = CustAcT(9).AvgAmt
              CustAcT(9).ConsumpAmt = TMeterConsp#
              CustAcT(9).CustAcct = CustomerRecord
              CustAcT(9).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(9).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(9).AvgAmt = AvgR
            ElseIf TMeterConsp# = CustAcT(9).ConsumpAmt Or TMeterConsp# > CustAcT(10).ConsumpAmt Then
              CustAcT(10).ConsumpAmt = TMeterConsp#
              CustAcT(10).CustAcct = CustomerRecord
              CustAcT(10).CustName = QPTrim$(UBCustRec(1).CustName)
              CustAcT(10).SvcAddr = QPTrim$(UBCustRec(1).ServAddr)
              CustAcT(10).AvgAmt = AvgR
            End If
Return
'  TblBreak&(RateRec, MaxStep(RateRec)) = 99999999
'  For LL = 1 To MaxStep(RateRec)
'    ToPrint$ = "Step # " + Str(LL)
'    ToPrint$ = ToPrint$ + "~" + Str(TblBreak&(RateRec, LL - 1)) + "~" + Str(TblBreak&(RateRec, LL))
'    ToPrint$ = ToPrint$ + "~" + Using("#,###,###,###", TotalConsp#(RateRec, LL))
'    ToPrint$ = ToPrint$ + "~" + Using("###,###", TotalCust(RateRec, LL))
'    If TblUnitVal#(RateRec, LL) > 0 Then
'    Else
'      TblUnitVal#(RateRec, LL) = 0
'    End If
'    BigUTotal# = Round#(BigUTotal# + (Round#(TotalConsp#(RateRec, LL) * TblUnitVal#(RateRec, LL))))
'    MinGT# = Round#(MinGT# + Round#(NMinAMT# * TotalCust(RateRec, LL)))
'    BigTotCust = BigTotCust + TotalCust(RateRec, LL)
'    If TotalCust(RateRec, LL) > 0 Then
'      ToPrint$ = ToPrint$ + "~" + Using("###,###,###.##", Round#(TotalConsp#(RateRec, LL) * TblUnitVal#(RateRec, LL))) + "~" + Using("  ###,###.##", Round#(NMinAMT# * TotalCust(RateRec, LL)))
'    Else
'      ToPrint$ = ToPrint$ + "~ ~ "
'    End If
'    Print #UBRpt, RCode$ + "~" + ToPrint$
'    ToPrint$ = ""
'  Next LL
'
'  GBBigUTotal# = Round#(GBBigUTotal# + BigUTotal#)
'  GBMinGT# = Round#(GBMinGT# + MinGT#)
'  GBGTMeterConsp# = Round#(GBGTMeterConsp# + GTMeterConsp#)
'  GBCustTot = GBCustTot + BigTotCust
'  BigUTotal# = 0
'  MinGT# = 0
'  BigTotCust = 0
'  GTMeterConsp# = 0
'Return
  GoTo ExitConsStep
ExitConsStep:
  Close
Exit Sub
End Sub

