VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmWOPrintSel 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Selected Work Orders"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmWOPrintSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   375
      Left            =   6195
      TabIndex        =   2
      Top             =   4605
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
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
      ColDesigner     =   "frmWOPrintSel.frx":08CA
   End
   Begin VB.CommandButton cmdOk 
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
      Left            =   8400
      TabIndex        =   3
      Top             =   7368
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
      TabIndex        =   4
      Top             =   7368
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
            TextSave        =   "10:55 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/4/2008"
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
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4020
      Left            =   2340
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   7548
      _Version        =   196609
      _ExtentX        =   13314
      _ExtentY        =   7091
      _StockProps     =   70
      Caption         =   ""
      Picture         =   "frmWOPrintSel.frx":0C30
      Begin LpLib.fpList fplstworkorders 
         Height          =   3090
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7110
         _Version        =   196608
         _ExtentX        =   12541
         _ExtentY        =   5450
         TextAlias       =   ""
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
         Columns         =   4
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
         ColDesigner     =   "frmWOPrintSel.frx":0C4C
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Alt-C &Clear All"
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
         Left            =   5856
         TabIndex        =   13
         Top             =   3504
         Width           =   1500
      End
      Begin VB.CommandButton cmdMark 
         Caption         =   "Alt-M &Mark All"
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
         Left            =   4272
         TabIndex        =   12
         Top             =   3504
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Spacebar or Click to Toggle, F10 to Continue. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Left            =   144
         TabIndex        =   17
         Top             =   3432
         Width           =   4140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   3180
         TabIndex        =   16
         Top             =   24
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   6084
         TabIndex        =   15
         Top             =   24
         Width           =   1164
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3156
         Left            =   120
         Top             =   240
         Width           =   7308
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   648
         TabIndex        =   14
         Top             =   24
         Width           =   1092
      End
   End
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   6192
      TabIndex        =   0
      Top             =   3552
      Width           =   1692
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   6192
      TabIndex        =   1
      Top             =   4080
      Width           =   1692
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
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3444
      Left            =   2388
      Top             =   2676
      Width           =   7452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Selected Work Orders"
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
      Index           =   0
      Left            =   3624
      TabIndex        =   9
      Top             =   1512
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   3192
      Top             =   1344
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   3192
      Top             =   1224
      Width           =   5772
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
      Height          =   372
      Left            =   4440
      TabIndex        =   7
      Top             =   3600
      Width           =   1668
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
      Height          =   372
      Index           =   0
      Left            =   4536
      TabIndex        =   8
      Top             =   4116
      Width           =   1572
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
      Height          =   372
      Left            =   3792
      TabIndex        =   6
      Top             =   4632
      Width           =   2388
   End
End
Attribute VB_Name = "frmWOPrintSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider

Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdOk.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
        KeyCode = 0
      End If
    End If
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
Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub
Private Sub txtDate1_LostFocus()
  If CheckValDate(txtDate1) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate1.SetFocus
  End If
End Sub
Private Sub txtDate2_LostFocus()
  If CheckValDate(txtDate2) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  frmUBWorkOrderMenu.Show
  Unload frmWOPrintSel
End Sub

Private Sub cmdOk_Click()
 If vaImprint1.Visible = False Then
  If ValidDate Then
  'If QPTrim$(fpcboVendFirst.ColText) <= QPTrim$(fpcboVendLast.ColText) Then
    LoadWOList
    vaImprint1.Visible = True
    fplstworkorders.SetFocus
    txtDate1.Visible = False
    txtDate2.Visible = False
    fpcboRptType.Visible = False
    'Shape3.Visible = False
'  Else
'    MsgBox "First and Last Vendors must be Equal or in Correct Order.", vbOKOnly, "Invalid Vendor Selection"
  End If
Else
  GetSelWOs
 'cmdExit_Click
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
        UBLog "Closed via WOPrintSel by " + PWUser$
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  Me.HelpContextID = hlpPrintSelectedWork
End Sub
Private Sub cmdClear_Click()
  fplstworkorders.Action = ActionDeselectAll
End Sub

Private Sub cmdMark_Click()
  fplstworkorders.Action = ActionSelectAll
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
'Private Sub LoadWOList()
''this fills the listbox with Work Orders for User to Select
'  Dim cnt As Integer, FirstVendRecNum As Long, disdate As String
'  Dim fmt As String, tempstr As String, strInfo As String
'  Dim LdRecLen As Integer, VendorFile As Integer, NumVRecs As Integer
'  Dim APLedgerFile As Integer, NumTran As Long, NextTrans As Long
'  Dim LastVendRecNum As Long, DueDate As Integer, TempAmt As String
'  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
'  Dim APLedgerRec(1) As APLedger81RecType
'  Dim VendorIdx As VendorIdxRecType
'  Dim Vendor As VendorRecType
'  LdRecLen = Len(APLedgerRec(1))
'  fmt = "$##########.##"
'  DueDate = DateDiff("d", "12/31/1979", fpDueDate)
'  fpcboVendFirst.col = 2
'  fpcboVendLast.col = 2
'  FirstVendRecNum = fpcboVendFirst.ColText
'  LastVendRecNum = fpcboVendLast.ColText
'  OpenVendorFile VendorFile, NumVRecs
'  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
'  OpenVendorIdx VendorIdxFile, NumActiveVendors
'  lstInvoices.Clear
'  For cnt = FirstVendRecNum To LastVendRecNum
'    Get VendorIdxFile, cnt, VendorIdx
'    If VendorIdx.RecNum > 0 Then
'    Get VendorFile, VendorIdx.RecNum, Vendor
'    NextTrans& = Vendor.FrstTran
'      Do Until NextTrans& = 0
'      Get APLedgerFile, NextTrans&, APLedgerRec(1)
'      If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
'        If APLedgerRec(1).DueDate <= DueDate Then
'          tempstr = Space$(15)
'          TempAmt$ = Using$(fmt$, Str$(APLedgerRec(1).Amt))
'          lstInvoices.AddItem Left$(Vendor.VNAME, 15) & Chr$(9) & QPTrim$(APLedgerRec(1).DOCNum) & Chr$(9) & Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy") & Chr$(9) & Right$(TempAmt$, 20) & Chr$(9) & NextTrans & Chr$(9) & VendorIdx.RecNum
'          lstInvoices.ListApplyTo = ListApplyToIndividual
'          lstInvoices.col = 3
'          lstInvoices.AlignH = AlignHRight
'        End If
'      End If
'
'      NextTrans& = APLedgerRec(1).NextTrans
'      Loop
'      End If
'    Next
'  Close
'  lstInvoices.Action = ActionSelectAll
'End Sub
  
'Private Sub GetSelInvs()
'    For Pcnt = 0 To fplstworkorders.ListCount - 1
'      If fplstworkorders.Selected(Pcnt) Then
'        cnt = cnt + 1
'        fplstworkorders.col = 4
'        fplstworkorders.Row = Pcnt
'         QPTrim (fplstworkorders.ColList)
'        fplstworkorders.col = 5
'         QPTrim (fplstworkorders.ColList)
'        Debug.Print fplstworkorders.ColList
'        GoSub SelectPrintOne
'
'      End If
'    Next
'
'  Close
'End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub fpCmdOK_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Button = 0
  Call cmdOk_Click
End Sub


'Private Sub fpWRKList_DblClick()
'  Dim TDate As String, cnt As Integer, TransRecNum As Long
'  Dim UBWrkOrd As Integer, WorkOrderRecLen As Integer
'  ReDim WorkOrderRec(1) As WorkOrderRecType
'  WorkOrderRecLen = Len(WorkOrderRec(1))
'  fpWRKList.col = 1                       'switch to the hidden RecNo. column
'  TransRecNum& = Val(fpWRKList.ColText)     'get customer recno
'  UBWrkOrd = FreeFile
'  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
'  Get UBWrkOrd, TransRecNum&, WorkOrderRec(1)
'  Close UBWrkOrd
''
''
'  frmRptWrkOrdDetail.LabelWONum.Caption = Str$(TransRecNum&)
'  If WorkOrderRec(1).CompletedDate <= 0 Then
'    TDate$ = "Open"
'  Else
'    TDate$ = Num2Date$(WorkOrderRec(1).CompletedDate)
'  End If
'  frmRptWrkOrdDetail.LabelEntryDate.Caption = Num2Date$(WorkOrderRec(1).ENTRYDATE)
'  frmRptWrkOrdDetail.LabelCompDate.Caption = TDate$
'  frmRptWrkOrdDetail.LabelCompBy.Caption = Num2Date$(WorkOrderRec(1).CompleteByDate)
'  frmRptWrkOrdDetail.LabelI1.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(1))
'  frmRptWrkOrdDetail.LabelI2.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(2))
'  frmRptWrkOrdDetail.LabelI3.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(3))
'  frmRptWrkOrdDetail.LabelI4.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(4))
'  frmRptWrkOrdDetail.LabelI5.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(5))
'  frmRptWrkOrdDetail.LabelI6.Caption = QPTrim(WorkOrderRec(1).OrdersText.Text(6))
'  frmRptWrkOrdDetail.LabelR1.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
'  frmRptWrkOrdDetail.LabelR2.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
'  frmRptWrkOrdDetail.LabelR3.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
'  frmRptWrkOrdDetail.LabelR4.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
'  frmRptWrkOrdDetail.LabelR5.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
'  frmRptWrkOrdDetail.LabelR6.Caption = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
'frmRptWrkOrdDetail.Show 1
'End Sub


'Private Sub fpWRKList_KeyUp(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape:
'      KeyCode = 0
'      Call cmdOk_Click
'    Case vbKeyReturn
'      KeyCode = 0
'      DoEvents
'      Call fpWRKList_DblClick  'fpCmdOK_Click
'    Case Else:
'  End Select
'
'End Sub
Private Sub LoadWOList()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim NumWrkRecs As Long, Handle As Integer, UBCFile As Integer
  Dim WOFile As Integer, cnt As Long, FromDate As Integer
  Dim ThruDate As Integer, DoneDate As Integer, WorkDate As Integer
  Dim dcnt As Long, Build As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))

  NumWrkRecs& = FileSize(UBPath$ + "UBWRKORD.DAT") / WorkOrderRecLen
  If NumWrkRecs& = 0 Then
    MsgBox "No Work Orders To Display.", vbOKOnly, "No Work Orders"
    GoTo ExitSelect
  End If

  FromDate = Date2Num(txtDate1)
  ThruDate = Date2Num(txtDate2)

'  ShowProcessingScrn "Searching for Open Work Orders"

  WOFile = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As WOFile Len = WorkOrderRecLen
  NumWrkRecs& = LOF(WOFile) / WorkOrderRecLen

  UBCFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCFile Len = UBCustRecLen

  For cnt = 1 To NumWrkRecs&
    Get WOFile, cnt, WorkOrderRec(1)
    DoneDate = WorkOrderRec(1).CompletedDate
    If DoneDate > 0 Then
      GoTo DoneSkip
    End If
    WorkDate = WorkOrderRec(1).ENTRYDATE
    'TempDate$ = Num2Date(WorkDate)
    'MID$(TempDate$, 10, 1) = "9"
    'WorkDate = Date2Num(TempDate$)
    'WorkOrderRec(1).EntryDate = WorkDate
    'PUT WOFile, Cnt, WorkOrderRec(1)

    If WorkDate >= FromDate And WorkDate <= ThruDate Then
      Get UBCFile, WorkOrderRec(1).CustRec, UBCustRec(1)
      dcnt = dcnt + 1
      
      Build$ = Space$(80)
      'LSet Build$ = "  " + Str$(PrevTranRec&)
      Build$ = " " + Num2Date(WorkOrderRec(1).ENTRYDATE) + Chr$(9) + QPTrim$(UBCustRec(1).CustName)
      Build$ = Build$ + Chr$(9) + Str$(Using$("######", WorkOrderRec(1).CustRec))
      Build$ = Build$ + Chr$(9) + Str$(cnt&)
      If Len(QPTrim(Build$)) > 0 Then
        frmWOPrintSel.fplstworkorders.AddItem Build$
      End If
    End If
DoneSkip:
    'ShowPctComp cnt, NumWrkRecs&
  Next
  Close
  'LabelTot.Caption = Str$(dcnt)
  'Me.Show 1
ExitSelect:
Exit Sub

End Sub
Private Sub GetSelWOs()
  Dim UBCustRecLen As Integer, WorkOrderRecLen As Integer
  Dim Dash As String, PrintSingleFlag As Boolean, Copies As Integer
  Dim ReportFile As String, RptHandle As Integer, IdxName As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, IdxNumOfRecs As Long
  Dim NumOfRecs As Long, Handle As Integer, UBCFile As Integer
  Dim UBWOFile As Integer, lcnt As Long, Book As Integer, cnt As Long
  Dim BegRoute As Integer, EndRoute As Integer, Acct As Long
  Dim Header As String, CopyCnt As Integer, MtrCnt As Integer
  Dim Rem1 As String, Rem2 As String, Rem3 As String, Rem4 As String
  Dim Rem5 As String, Rem6 As String, ToPrint As String
  Dim graphicflag As Boolean, dcnt As Long, PCnt As Integer
  Dim Rec As Long, CustNum As Long
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  Rem1$ = ""
  Rem2$ = ""
  Rem3$ = ""
  Rem4$ = ""
  Rem5$ = ""
  Rem6$ = ""
  If fpcboRptType.ListIndex = 0 Then
    graphicflag = True
  Else
    graphicflag = False
  End If
  If graphicflag = True Then
    Dash$ = String$(83, "_")
  Else
    Dash$ = String$(79, "_")
  End If
  FrmShowPctComp.Label1 = "Creating Work Order"
  FrmShowPctComp.Show , Me

  ToPrint$ = ""
  FF$ = Chr$(12)
  'If dcnt > 0 Then
    ReportFile$ = UBPath$ + "WORKORDR.RPT"   'Open Report File
   ' ShowProcessingScrn "Printing Work Orders."

    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle

    UBCFile = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCFile Len = UBCustRecLen

    UBWOFile = FreeFile
    Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWOFile Len = WorkOrderRecLen
    For PCnt = 0 To fplstworkorders.ListCount - 1
      FrmShowPctComp.ShowPctComp PCnt, (fplstworkorders.ListCount - 1)
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        GoTo ExitSelect
      End If

      If fplstworkorders.Selected(PCnt) Then
        cnt = cnt + 1
        fplstworkorders.col = 3
        fplstworkorders.Row = PCnt
        Rec = QPTrim(fplstworkorders.ColList)
        fplstworkorders.col = 2
        CustNum = QPTrim(fplstworkorders.ColList)
        'Debug.Print fplstworkorders.ColList
        Get UBCFile, CustNum&, UBCustRec(1)
        If UBCustRec(1).WOLastTrans > 0 Then
          Get UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
          GoSub SelectPrintOne
        Else
          Close
          Unload FrmShowPctComp
          GoTo ExitSelect
        End If
      End If
    Next

'    For cnt = 1 To dcnt
'      Temp$ = Mid$(MTChoice(Picked(cnt)).V, 41, 6)
'      Temp$ = QPTrim$(Temp$)
'      CustRec& = QPValL(Temp$)
'      GoSub SelectPrintOne:
'      ShowPctComp cnt, dcnt
'    Next
    Close

    Header$ = "Print Selected Work Orders."

    'EntryPoint = 1
    'LPTPort = 1
    If graphicflag = False Then
      ViewPrint ReportFile$, Header$
    Else
      Load frmLoadingRpt
      frmLoadingRpt.setwherefrom frmWOPrintSel
      ARptWorkOrder.GetName ReportFile$
      ARptWorkOrder.startrpt
    End If

    'rintRptFile Header$, ReportFile$, LPTPort, RetCode, EntryPoint

  'End If

ExitSelect:
Exit Sub

SelectPrintOne:
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(1))) > 0 Then
    Rem1$ = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
  Else
    Rem1$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(2))) > 0 Then
    Rem2$ = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
  Else
    Rem2$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(3))) > 0 Then
    Rem3$ = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
  Else
    Rem3$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(4))) > 0 Then
    Rem4$ = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
  Else
    Rem4$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(5))) > 0 Then
    Rem5$ = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
  Else
    Rem5$ = Dash$
  End If

  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(6))) > 0 Then
    Rem6$ = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
  Else
    Rem6$ = "BY: ______________________________   DATE: ____________________"
  End If
 
  If graphicflag = False Then
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "Printed:"; Now
    Print #RptHandle, " "
    Print #RptHandle, Tab(14); "W O R K   O R D E R   :   U T I L I T Y   D E P T ."
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "    Work Order#: "; Using("######", UBCustRec(1).WOLastTrans); Tab(30); "Date Issued: "; Num2Date$(WorkOrderRec(1).ENTRYDATE)
    Print #RptHandle, "      Location#: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Tab(30); "Complete By: "; Num2Date$(WorkOrderRec(1).CompleteByDate)
    Print #RptHandle, "       Account#: "; Acct&; Tab(30); "  Completed: "; Num2Date$(WorkOrderRec(1).CompletedDate)
    Print #RptHandle, "  Customer Name: "; UBCustRec(1).CustName
    Print #RptHandle, "Service Address: "; UBCustRec(1).ServAddr
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Instruction or Description of Work Needed"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(1)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(2)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(3)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(4)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(5)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(6)
    Print #RptHandle, " "
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Remarks Noted by Worker"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, Rem1$
    Print #RptHandle, " "
    Print #RptHandle, Rem2$
    Print #RptHandle, " "
    Print #RptHandle, Rem3$
    Print #RptHandle, " "
    Print #RptHandle, Rem4$
    Print #RptHandle, " "
    Print #RptHandle, Rem5$
    Print #RptHandle, " "
    Print #RptHandle, Rem6$
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "Meter Numbers:"

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #RptHandle, QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      End If
    Next
    Print #RptHandle, FF$;
  Else
    ToPrint$ = Num2Date$(WorkOrderRec(1).ENTRYDATE) + "~"
    ToPrint$ = ToPrint$ + Using("######", UBCustRec(1).WOLastTrans) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~"
    ToPrint$ = ToPrint$ + Str(Acct&) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).CustName + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).ServAddr + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(1) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(2) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(3) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(4) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(5) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(6) + "~"
    ToPrint$ = ToPrint$ + Rem1$ + "~"
    ToPrint$ = ToPrint$ + Rem2$ + "~"
    ToPrint$ = ToPrint$ + Rem3$ + "~"
    ToPrint$ = ToPrint$ + Rem4$ + "~"
    ToPrint$ = ToPrint$ + Rem5$ + "~"
    ToPrint$ = ToPrint$ + Rem6$

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      Else
        ToPrint$ = ToPrint$ + "~ "
      End If
    Next
    ToPrint$ = ToPrint$ + "~" + Num2Date$(WorkOrderRec(1).CompleteByDate) + "~"
    ToPrint$ = ToPrint$ + Num2Date$(WorkOrderRec(1).CompletedDate)
    Print #RptHandle, ToPrint$
    ToPrint$ = ""
  End If
  Return

End Sub
