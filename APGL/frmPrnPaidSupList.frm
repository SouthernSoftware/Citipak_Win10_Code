VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnPaidSupList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paid Supply List"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   12192
   Icon            =   "frmPrnPaidSupList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   6144
      TabIndex        =   2
      Top             =   4800
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   677
      Text            =   ""
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
      ColDesigner     =   "frmPrnPaidSupList.frx":08CA
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
      Left            =   8256
      TabIndex        =   3
      Top             =   7464
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
      Left            =   10032
      TabIndex        =   4
      Top             =   7464
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
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
            TextSave        =   "3:54 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "12/8/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime fpDate1 
      Height          =   372
      Left            =   6162
      TabIndex        =   0
      Top             =   3420
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
   Begin EditLib.fpDateTime fpDate2 
      Height          =   372
      Left            =   6168
      TabIndex        =   1
      Top             =   4116
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
      Left            =   3672
      TabIndex        =   9
      Top             =   4848
      Width           =   2388
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2628
      Left            =   2172
      Top             =   2904
      Width           =   7860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Supply List"
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
      Left            =   3984
      TabIndex        =   8
      Top             =   1488
      Width           =   4332
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1248
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   2496
      Picture         =   "frmPrnPaidSupList.frx":0CB8
      Top             =   3048
      Width           =   288
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
      Height          =   420
      Left            =   4440
      TabIndex        =   7
      Top             =   4188
      Width           =   1572
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
      Left            =   4338
      TabIndex        =   6
      Top             =   3456
      Width           =   1668
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3216
      Top             =   1128
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
Attribute VB_Name = "frmPrnPaidSupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType

Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Sub cmdExit_Click()
  frmAPReportsMenu.Show
  Unload frmPrnPaidSupList
End Sub

Private Sub cmdOk_Click()
 If Oktogo = True Then
  If fpcboRptType.ListIndex = 0 Then
     rptopt = 1
   ElseIf fpcboRptType.ListIndex = 1 Then
     rptopt = 2
   End If
   If rptopt = 1 Then
     PaidSupplyList
   ElseIf rptopt = 2 Then
     PaidSupplyList2
   End If
 End If
End Sub
'Private Sub fpcboAllSel_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboAllSel.ListDown = True
'  End If
'End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
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
      SendKeys "%O"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdOk.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpDate2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
'  fpcboAllSel.AddItem "All"
'  fpcboAllSel.AddItem "Selected"
'  fpcboAllSel.ListIndex = 0
  Me.HelpContextID = hlpPSL

  fpDate1.Text = Format(Now, "mm/dd/yyyy")
  fpDate2.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Function Oktogo()
Dim TempDate1 As Integer, TempDate2 As Integer
    If CheckValDate(fpDate1) = False And CheckValDate(fpDate2) = False Then
      MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
      Oktogo = False
    Else
      TempDate1 = DateDiff("d", "12/31/1979", fpDate1)
      TempDate2 = DateDiff("d", "12/31/1979", fpDate2)
      If TempDate1 > TempDate2 Then
        Oktogo = False
        MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
      Else
        Oktogo = True
      End If
    End If
End Function
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PaidSupplyList()
  Dim BegDate As Integer, EndDate As Integer, NumTran As Long
  Dim Showselected As Boolean, cnt As Long, ToPrint1 As String
  Dim Page As Integer, NumFunds As Integer, PRNFile As String
  Dim ColTitle As String, VendorFile As Integer, APDistRecLen As Integer
  Dim Header As String, A As String, CommaFmt As String, User As String
  Dim APLRecLen As Integer, APLedgerFile As Integer, RunTotal As Double
  Dim APDRecLen As Integer, APDistFile As Integer, NumDistRecs As Long
  Dim RptFile As Integer, RptFileName As String, VRecLen As Integer
  Dim NumVRecs As Integer, TAmt As Double, NumChks As Integer
  Dim Linecnt As Integer, Rec As Long, CheckNumber As Long
  Dim ToPrint As String, NextTrans As Long, DistAmt As Double
  Dim Found As Boolean, Fund As Integer, FundNum As String, FCnt As Integer
  Dim lngCurLow As Long, lngCurHigh As Long, ChkNum As String
  Dim PrintedCheck As Boolean, NextDist As Long, DeptCode As String
  Dim DeptRecNum As Integer, DeptName As String, Newrp As String
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  FrmShowPctComp.Label1 = "Creating Paid Supply List Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  mnuOptions.Enabled = False
 ' Showselected = False
  User$ = QPTrim$(GLUserName$)
'  Page = 0
'  If fpcboAllSel.ListIndex = 1 Then
'    Showselected = True
'  End If
  CommaFmt$ = "##,###,###.##"
  'FF$ = Chr$(12)
 ' MaxLines = 55

  Dim APLedger As APLedger81RecType
  APLRecLen = Len(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  ReDim ChkList(1 To 1) As ChkSortType
  ChkNum$ = Space$(14)

  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FreeFile
  Newrp = "PSLRpt.prn"
  'GetRPTName Newrp
  RptFileName$ = Newrp
  Open RptFileName$ For Output As RptFile

  Dim Vendor As VendorRecType
  VRecLen = Len(Vendor)
  OpenVendorFile VendorFile, NumVRecs

  

  'Get a list of checks
  For cnt = 1 To NumTran&
    FrmShowPctComp.ShowPctComp cnt, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get APLedgerFile, cnt, APLedger
    Get VendorFile, APLedger.VRecNum, Vendor
    If APLedger.TRCode = 3 Then
      If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
        NumChks = NumChks + 1
        ReDim Preserve ChkList(1 To NumChks) As ChkSortType
        ChkList(NumChks).Record = cnt
        TAmt# = Round#(TAmt# + APLedger.Amt)
        RSet ChkNum$ = QPTrim$(APLedger.DOCNum)
        ChkList(NumChks).CHKinfo = ChkNum$
      End If
    End If
  Next
  If NumTran& < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If
  If NumChks > 0 Then
    lngCurLow = LBound(ChkList)
    lngCurHigh = UBound(ChkList)
    QCSort ChkList(), lngCurLow, lngCurHigh
    
    GoSub SearchforPSLTrans
     
    
  Else
    MsgBox "No Checks on file.", vbOKOnly, "No Checks"
  End If
  Close
  Load frmLoadingRpt
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  ARptPaidSupList.GetName RptFileName$
  ARptPaidSupList.totchks = Using("##,###,###.##", RunTotal#)
  ARptPaidSupList.txtDate.Caption = Now
  ARptPaidSupList.txtTown.Caption = User$
  ARptPaidSupList.Label1.Caption = "Paid Supply List"
  ARptPaidSupList.Label17.Caption = fpDate1.Text + " thru " + fpDate2.Text
  ARptPaidSupList.startrpt
  Exit Sub

SearchforPSLTrans:
  FrmShowPctComp.Label1 = "Sorting List"
  FrmShowPctComp.Show , Me
  '--search thru check list
  For cnt = 1 To NumChks
    FrmShowPctComp.ShowPctComp cnt, NumChks
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    ChkNum$ = Space$(14)

    '--Get the pointers we need
    Get APLedgerFile, ChkList(cnt).Record, APLedger
    'dale
    'ThisChk# = APLedger.Amt

    Get VendorFile, APLedger.VRecNum, Vendor
    NextTrans& = Vendor.FrstTran

    '--Look for this check Number in the pd check num field
    CheckNumber& = Val(ChkList(cnt).CHKinfo)
    PrintedCheck = False

    '--Search thru the ledger file to see if trans is flagged for PSL
    Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedger
      If APLedger.PDCheckNum = CheckNumber& Then
        'If Showselected Then
          If APLedger.PSLFlag = "Y" Then
            '--print check info only one time - check may be for multiple invo
            If Not PrintedCheck Then
              GoSub PrintCheckInfo
            End If
            NextDist& = APLedger.FrstDist
            Do
              Get APDistFile, NextDist&, APDist
              If Not Round#(APDist.DistAmt) = -1.79769313486231E+308 Then
                DistAmt# = Round#(DistAmt# + APDist.DistAmt)
                GoSub PrintDist
              End If
              NextDist& = APDist.NextDist
            Loop Until NextDist& = 0
          End If
        'Else

          '112299 Add flag check
'          If APLedger.PSLFlag = "Y" Then
'            If Not PrintedCheck Then
'            GoSub PrintCheckInfo
'            End If
'            NextDist& = APLedger.FrstDist
'            Do
'              Get APDistFile, NextDist&, APDist
'              DistAmt# = Round#(DistAmt# + APDist.DistAmt)
'              NextDist& = APDist.NextDist
'              GoSub PrintDist
'            Loop Until NextDist& = 0
'          End If
'          'Dale
'        End If
      End If
      NextTrans& = APLedger.NextTrans
    Loop
    If PrintedCheck Then GoSub PrintChkTotals
  Next

  Return

PrintCheckInfo:
  PrintedCheck = True

  ToPrint1$ = Space$(40)
  ToPrint1$ = Str$(CheckNumber&) + "~" + Vendor.VNAME + "~" + APLedger.Comment + "~"
  Return

PrintDist:
  DeptCode$ = Mid$(APDist.DistAcctNum, GLFundLen + 2, GLAcctLen)
  DeptRecNum = FindDept(DeptCode$)
  If DeptRecNum > 0 Then
    DeptName$ = GetDeptTitle$(DeptRecNum)
  Else
    DeptName$ = "Undefined Dept" + DeptCode$
  End If
  ToPrint$ = Space$(40)
  ToPrint$ = DeptName$ + "~" + Using(CommaFmt$, Str$(APDist.DistAmt))
  Print #RptFile, ToPrint1$ + ToPrint$
  Return

PrintChkTotals:

  RunTotal# = Round#(RunTotal# + DistAmt#)
  DistAmt# = 0
  Return
CancelExit:
  Exit Sub
End Sub
Private Sub PaidSupplyList2()
  Dim BegDate As Integer, EndDate As Integer, NumTran As Long
  Dim Showselected As Boolean, cnt As Long, FF As String, MaxLines As Integer
  Dim Page As Integer, NumFunds As Integer, PRNFile As String
  Dim ColTitle As String, VendorFile As Integer, APDistRecLen As Integer
  Dim Header As String, A As String, CommaFmt As String, User As String
  Dim APLRecLen As Integer, APLedgerFile As Integer, RunTotal As Double
  Dim APDRecLen As Integer, APDistFile As Integer, NumDistRecs As Long
  Dim RptFile As Integer, RptFileName As String, VRecLen As Integer
  Dim NumVRecs As Integer, TAmt As Double, NumChks As Integer
  Dim Linecnt As Integer, Rec As Integer, CheckNumber As Long
  Dim ToPrint As String, NextTrans As Long, DistAmt As Double
  Dim Found As Boolean, Fund As Integer, FundNum As String, FCnt As Integer
  Dim lngCurLow As Long, lngCurHigh As Long, ChkNum As String
  Dim PrintedCheck As Boolean, NextDist As Long, DeptCode As String
  Dim DeptRecNum As Integer, DeptName As String, Newrp As String
  BegDate = DateDiff("d", "12/31/1979", fpDate1)
  EndDate = DateDiff("d", "12/31/1979", fpDate2)
  FrmShowPctComp.Label1 = "Creating Paid Supply List Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  mnuOptions.Enabled = False
 ' Showselected = False
  User$ = QPTrim$(GLUserName$)
  Page = 0
'  If fpcboAllSel.ListIndex = 1 Then
'    Showselected = True
'  End If
  CommaFmt$ = "##,###,###.##"
  FF$ = Chr$(12)
  MaxLines = 55

  Dim APLedger As APLedger81RecType
  APLRecLen = Len(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  'ReDim ChkList(1 To 1) As GLAcctIndexType      '--borrowing this type
  ReDim ChkList(1 To 1) As ChkSortType
  ChkNum$ = Space$(14)

  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FreeFile
  Newrp = "PSL"
  GetRPTName Newrp
  RptFileName$ = Newrp
  Open RptFileName$ For Output As RptFile

  Dim Vendor As VendorRecType
  VRecLen = Len(Vendor)
  OpenVendorFile VendorFile, NumVRecs

  GoSub PSLPageHdr

  'Get a list of checks
  For cnt = 1 To NumTran&
    FrmShowPctComp.ShowPctComp cnt, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get APLedgerFile, cnt, APLedger
    Get VendorFile, APLedger.VRecNum, Vendor
    If APLedger.TRCode = 3 Then
      If APLedger.TRDATE >= BegDate And APLedger.TRDATE <= EndDate Then
        NumChks = NumChks + 1
        ReDim Preserve ChkList(1 To NumChks) As ChkSortType
        ChkList(NumChks).Record = cnt
        TAmt# = Round#(TAmt# + APLedger.Amt)
        RSet ChkNum$ = QPTrim$(APLedger.DOCNum)
        ChkList(NumChks).CHKinfo = ChkNum$
      End If
    End If
  Next
  If NumTran& < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If
  If NumChks > 0 Then
    lngCurLow = LBound(ChkList)
    lngCurHigh = UBound(ChkList)
    QCSort ChkList(), lngCurLow, lngCurHigh
    
    GoSub SearchforPSLTrans
    Print #RptFile, "Total Checks ", Using("##,###,###.##", RunTotal#);
    Print #RptFile, FF$
  Else
    Print #RptFile, "  No Checks on file."
  End If
  Close


  ViewPrint RptFileName$, Header$
  KillFile RptFileName$
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  Exit Sub

SearchforPSLTrans:
  FrmShowPctComp.Label1 = "Sorting List"
  FrmShowPctComp.Show , Me
  '--search thru check list
  For cnt = 1 To NumChks
    FrmShowPctComp.ShowPctComp cnt, NumChks
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    ChkNum$ = Space$(14)

    '--Get the pointers we need
    Get APLedgerFile, ChkList(cnt).Record, APLedger
    'dale
    'ThisChk# = APLedger.Amt

    Get VendorFile, APLedger.VRecNum, Vendor
    NextTrans& = Vendor.FrstTran

    '--Look for this check Number in the pd check num field
    CheckNumber& = Val(ChkList(cnt).CHKinfo)
    PrintedCheck = False

    '--Search thru the ledger file to see if trans is flagged for PSL
    Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedger
      If APLedger.PDCheckNum = CheckNumber& Then
        'If Showselected Then
          If APLedger.PSLFlag = "Y" Then
            '--print check info only one time - check may be for multiple invo
            If Not PrintedCheck Then
              GoSub PrintCheckInfo
            End If
            NextDist& = APLedger.FrstDist
            Do
              Get APDistFile, NextDist&, APDist
              DistAmt# = Round#(DistAmt# + APDist.DistAmt)
              NextDist& = APDist.NextDist
              GoSub PrintDist
            Loop Until NextDist& = 0
          End If
        'Else

          '112299 Add flag check
'          If APLedger.PSLFlag = "Y" Then
'            If Not PrintedCheck Then
'            GoSub PrintCheckInfo
'            End If
'            NextDist& = APLedger.FrstDist
'            Do
'              Get APDistFile, NextDist&, APDist
'              DistAmt# = Round#(DistAmt# + APDist.DistAmt)
'              NextDist& = APDist.NextDist
'              GoSub PrintDist
'            Loop Until NextDist& = 0
'          End If
'          'Dale
'        End If
      End If
      NextTrans& = APLedger.NextTrans
    Loop
    If PrintedCheck Then GoSub PrintChkTotals
  Next

  Return

PSLPageHdr:
  Page = Page + 1
  Print #RptFile, Tab(40 - (Len(User$) / 2)); User$
  Print #RptFile, Tab(30); " Paid Supply List"
  Print #RptFile, Tab(27); fpDate1.Text + " thru " + fpDate2.Text
  Print #RptFile,
  Print #RptFile, "Chk#     Vendor                     Description"
  Print #RptFile, "                                    Department"
  Print #RptFile, String$(78, "=")
  Linecnt = 7
  Return

PrintCheckInfo:
  PrintedCheck = True

  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 2) = Str$(CheckNumber&)
  Mid$(ToPrint$, 10) = Vendor.VNAME
  Mid$(ToPrint$, 37) = APLedger.Comment
  Print #RptFile, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #RptFile, FF$
    GoSub PSLPageHdr
  End If
  Return

PrintDist:
  DeptCode$ = Mid$(APDist.DistAcctNum, GLFundLen + 2, GLAcctLen)
  DeptRecNum = FindDept(DeptCode$)
  If DeptRecNum > 0 Then
    DeptName$ = GetDeptTitle$(DeptRecNum)
  Else
    DeptName$ = "Undefined Dept" + DeptCode$
  End If
  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 37) = DeptName$
  Mid$(ToPrint$, 67) = Using(CommaFmt$, Str$(APDist.DistAmt))
  Print #RptFile, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #RptFile, FF$
    GoSub PSLPageHdr
  End If
  Return

PrintChkTotals:
  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 67) = "------------"
  Print #RptFile, ToPrint$
  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 67) = Using(CommaFmt$, Str$(DistAmt#))
  Print #RptFile, ToPrint$
  ToPrint$ = Space$(80)
  Print #RptFile, ToPrint$
  Linecnt = Linecnt + 3
  If Linecnt > MaxLines Then
    Print #RptFile, FF$
    GoSub PSLPageHdr
  End If

 '**********Nick Moved RunTotal to Here to Add actual Dist Amts for Grand Tota
  RunTotal# = RunTotal# + DistAmt#
  DistAmt# = 0
  Return
CancelExit:
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
