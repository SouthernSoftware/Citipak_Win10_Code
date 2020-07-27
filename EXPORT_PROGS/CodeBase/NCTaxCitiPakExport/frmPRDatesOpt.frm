VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPRDatesOpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Payroll Dates / Options"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11628
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   528
      Left            =   3936
      TabIndex        =   14
      Top             =   7968
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
      Height          =   528
      Left            =   6480
      TabIndex        =   13
      Top             =   7968
      Width           =   1332
   End
   Begin FPSpread.vaSpread vaSpreadAdditional 
      Height          =   1896
      Left            =   6144
      TabIndex        =   6
      Top             =   5376
      Width           =   4620
      _Version        =   196613
      _ExtentX        =   8149
      _ExtentY        =   3344
      _StockProps     =   64
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      SpreadDesigner  =   "frmPRDatesOpt.frx":0000
   End
   Begin FPSpread.vaSpread vaSpreadEmp2Pay 
      Height          =   1356
      Left            =   4056
      TabIndex        =   4
      Top             =   3504
      Width           =   3564
      _Version        =   196613
      _ExtentX        =   6287
      _ExtentY        =   2392
      _StockProps     =   64
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   10
      SpreadDesigner  =   "frmPRDatesOpt.frx":182B
   End
   Begin LpLib.fpList fplistDef 
      Height          =   288
      Left            =   6816
      TabIndex        =   3
      Top             =   2592
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
      _ExtentY        =   508
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
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
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
      EnableClickEvent=   -1  'True
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmPRDatesOpt.frx":1B5A
   End
   Begin FPSpread.vaSpread vaSpreadDed2Take 
      Height          =   1884
      Left            =   984
      TabIndex        =   5
      Top             =   5376
      Width           =   4620
      _Version        =   196613
      _ExtentX        =   8149
      _ExtentY        =   3323
      _StockProps     =   64
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      SpreadDesigner  =   "frmPRDatesOpt.frx":1DAE
   End
   Begin EditLib.fpDateTime fpMaskBeginDate 
      Height          =   372
      Left            =   4128
      TabIndex        =   1
      Top             =   2016
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   656
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
      Text            =   "10-01-2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm-dd-yyyy"
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
   Begin EditLib.fpDateTime fpMaskEndDate 
      Height          =   372
      Left            =   8256
      TabIndex        =   2
      Top             =   2016
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   656
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
      Text            =   "10-01-2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm-dd-yyyy"
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
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4524
      Left            =   504
      Top             =   3072
      Width           =   10620
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   " Additional Earnings Accounts to Use"
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
      Height          =   345
      Left            =   6300
      TabIndex        =   12
      Top             =   4950
      Width           =   4230
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   " Deductions to Take"
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
      Left            =   2088
      TabIndex        =   11
      Top             =   4944
      Width           =   2364
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1212
      Left            =   504
      Top             =   1872
      Width           =   10620
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   " Employees to Pay:"
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
      Left            =   4632
      TabIndex        =   10
      Top             =   3120
      Width           =   2364
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Set Payroll with Defaults?"
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
      Left            =   3768
      TabIndex        =   9
      Top             =   2640
      Width           =   2796
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
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
      Height          =   348
      Left            =   6264
      TabIndex        =   8
      Top             =   2064
      Width           =   1980
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "Pay Period Beginning Date:"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2064
      Width           =   2652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Payroll Dates / Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2712
      TabIndex        =   0
      Top             =   720
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Index           =   1
      Left            =   1464
      Top             =   360
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1464
      Top             =   240
      Width           =   8652
   End
End
Attribute VB_Name = "frmPRDatesOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdProcess_Click()
'    Call PayDedReport
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
      SendKeys "%x"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim cnt As Integer
  Dim ScrWidth As Long
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Select Case ScrWidth
  Case 1280
    For cnt = 1 To 2
      vaSpreadEmp2Pay.ColWidth(cnt) = vaSpreadEmp2Pay.ColWidth(cnt) + 4.75
      vaSpreadDed2Take.ColWidth(cnt) = vaSpreadDed2Take.ColWidth(cnt) + 6.25
      vaSpreadAdditional.ColWidth(cnt) = vaSpreadAdditional.ColWidth(cnt) + 6.25
    Next
  Case 1024
    For cnt = 1 To 2
      vaSpreadEmp2Pay.ColWidth(cnt) = vaSpreadEmp2Pay.ColWidth(cnt) + 1
      vaSpreadDed2Take.ColWidth(cnt) = vaSpreadDed2Take.ColWidth(cnt) + 1.2
      vaSpreadAdditional.ColWidth(cnt) = vaSpreadAdditional.ColWidth(cnt) + 1.2
    Next
  Case 800
    For cnt = 1 To 2
      vaSpreadEmp2Pay.ColWidth(cnt) = vaSpreadEmp2Pay.ColWidth(cnt) + 0.55
      vaSpreadDed2Take.ColWidth(cnt) = vaSpreadDed2Take.ColWidth(cnt) + 0.75
      vaSpreadAdditional.ColWidth(cnt) = vaSpreadAdditional.ColWidth(cnt) + 0.75
    Next
  Case Else
  End Select
  Call LoadPRDatesOptScreen
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub


Private Sub cmdExit_Click()
  frmPayrollProcessingMenu.Show
  Unload frmPRDatesOpt
End Sub

Private Sub LoadPRDatesOptScreen()
  Dim x As Integer
  Dim DHandle As Integer
  Dim EHandle As Integer
  Dim ERecs As ErnCodeRecType
  Dim DRecs As DedCodeRecType
  Dim DCnt As Long
  Dim ECnt As Long
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 1
  vaSpreadEmp2Pay.Text = "Semi-Monthly"
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 2
  vaSpreadEmp2Pay.Text = "Semi-Annually"
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 3
  vaSpreadEmp2Pay.Text = "Weekly"
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 4
  vaSpreadEmp2Pay.Text = "Monthly"
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 5
  vaSpreadEmp2Pay.Text = "Annually"
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 6
  vaSpreadEmp2Pay.Text = "Bi-Weekly"
  vaSpreadEmp2Pay.Col = 1
  vaSpreadEmp2Pay.Row = 7
  vaSpreadEmp2Pay.Text = "Quarterly"
  
  OpenDedCodeFile DHandle
  DCnt = LOF(DHandle) / Len(DRecs)
  For x = 1 To DCnt
    Get DHandle, x, DRecs
    vaSpreadDed2Take.Col = 1
    vaSpreadDed2Take.Row = x
    vaSpreadDed2Take.Text = DRecs.DCDESC1
  Next x
  Close DHandle
  
  OpenErnCodeFile EHandle
  ECnt = LOF(EHandle) / Len(ERecs)
  For x = 1 To ECnt
     Get EHandle, x, ERecs
     vaSpreadAdditional.Col = 1
     vaSpreadAdditional.Row = x
     vaSpreadAdditional.Text = ERecs.ERNCODE1
  Next x
  Close EHandle
  
  fplistDef.AddItem ("Y")
  fplistDef.AddItem ("N")
  
End Sub
