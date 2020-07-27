VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnCashBal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Balance "
   ClientHeight    =   8640
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   12195
   Icon            =   "frmPrnCashBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   6285
      TabIndex        =   1
      Top             =   4050
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
      _ExtentY        =   714
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
      ColDesigner     =   "frmPrnCashBal.frx":08CA
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
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
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
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
      Left            =   8256
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "5:05 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "1/17/2008"
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
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   6282
      TabIndex        =   0
      Top             =   3360
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
      ButtonColor     =   14737632
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
      Left            =   3528
      TabIndex        =   7
      Top             =   4104
      Width           =   2388
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
      Left            =   4266
      TabIndex        =   6
      Top             =   3408
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   3300
      Picture         =   "frmPrnCashBal.frx":0BF8
      Top             =   3270
      Width           =   360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Cash Balance Report"
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
      Left            =   4272
      TabIndex        =   5
      Top             =   1608
      Width           =   3852
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1836
      Left            =   2688
      Top             =   2976
      Width           =   6828
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   1248
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
Attribute VB_Name = "frmPrnCashBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim FirstFund As String, LastFund As String
Dim ActiveYear As Integer

Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnCashBal
End Sub

Private Sub cmdPrint_Click()
  If CheckValDate(txtDate) = True Then
    If fpcboRptType.ListIndex = 0 Then
      rptopt = 1
    ElseIf fpcboRptType.ListIndex = 1 Then
      rptopt = 2
    End If
    If rptopt = 1 Then
      PrintCashBal
    ElseIf rptopt = 2 Then
      PrintCashBal2
    End If
  Else
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
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
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpCashBalance
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
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
        txtDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrintCashBal()
  Dim PRNFile As Integer, FundCnt As Integer, cnt As Integer
  Dim ReportFile As String, ToPrintB As String, CommaFmt As String
  Dim TotalFmt As String, BankGLAcct As String, cntp As Long, ToPrint As String
  Dim BankAcctBal As Double, TotalCashBal As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim NumFunds As Integer, EndDate As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, AcctRec As Integer, MaxLines As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long
  Dim GLBank(1) As GLBankRecType, BankRecLen As Integer, Linecnt As Integer
  Dim NumBanks As Integer, BankFile As Integer, Newrp As String
  BankRecLen = Len(GLBank(1))
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
'--Bank fields
'Bank(1).Deleted
'Bank(1).BankNum
'Bank(1).BankName
'Bank(1).BankAcct
'Bank(1).GLAcct
'================
  'OpenBankFile BankRecLen, BankFile, NumBanks
'''  BankFile = FreeFile
'''  Open "GLBANK.DAT" For Random Access Read Write Shared As BankFile Len = Bank
'''  NumBanks = LOF(BankFile) \ BankRecLen
  '****
  OpenBankFile BankFile, NumBanks
  If NumBanks = 0 Then
    Close BankFile
    
    Beep
    MsgBox "Error: Banks not defined. Press any key to return to menu.", vbOKOnly, "Cash Balance"
    
    Exit Sub
  End If

  ReDim FundList$(1)
  GetFundList FundList$(), NumFunds
  EndDate = DateDiff("d", "12/31/1979", txtDate)
  'Newrp = "CshBal"
  'GetRPTName Newrp
  ReportFile$ = "CashBal.prn"
  MaxLines = 53
  OpenAcctFile AcctFileNum
  NumGLAccts = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  'GoSub PrintHeader
  FrmShowPctComp.Label1 = "Printing Cash Balance Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdPrint.Enabled = False
  Me.mnuOptions.Enabled = False
  For cnt = 1 To NumBanks
    Get BankFile, cnt, GLBank(1)
    
    If Not GLBank(1).Deleted Then
    'Bank(1).Deleted
    'Bank(1).BankNum
    'Bank(1).BankName
    'Bank(1).BankAcct
    'Bank(1).GLAcct
    ToPrintB$ = GLBank(1).BankName
    For FundCnt = 1 To NumFunds  'FundList$(1)
      BankGLAcct$ = QPTrim$(FundList$(FundCnt) + GLBank(1).GLAcct)
      AcctRec = AcctFind(BankGLAcct$)
      ToPrint$ = Space(80)
      If AcctRec > 0 Then
        Get AcctFileNum, AcctRec, GLAcct
        CalcBal# = Round#(GLAcct.BegBal)
        NextTr& = GLAcct.FrstTran
        Do Until NextTr& = 0
          Get TransFileNum, NextTr&, GLTrans
            If GLTrans.TRDATE <= EndDate Then
                CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            End If
          NextTr& = GLTrans.NextTran
        Loop
        ToPrint$ = ToPrintB$ + "~" + BankGLAcct$ + "~" + Using(CommaFmt$, Str$(CalcBal#))
        Print #PRNFile, ToPrint$
        TotalCashBal# = TotalCashBal# + CalcBal#
        BankAcctBal# = BankAcctBal# + CalcBal#
  
      End If
    Next

    'Print #PRNFile, "Total for "; QPTrim$(GLBank(1).BankName);
    'Print #PRNFile, Tab(30); Using(CommaFmt$, Str$(BankAcctBal#))
    BankAcctBal# = 0
  End If
  
    FrmShowPctComp.ShowPctComp cnt, NumBanks
    If FrmShowPctComp.Out = True Then
      Close
      'FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      Me.mnuOptions.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
   Me.cmdExit.Enabled = True
   Me.cmdPrint.Enabled = True
   Me.mnuOptions.Enabled = True
   EnableCloseButton Me.hwnd, True
   Close BankFile
  ' Print #PRNFile, "Total Cash"; Tab(27); Using(CommaFmt$, Str$(TotalCashBal#))
   Close
   Load frmLoadingRpt
   ARptCashBal.txtRptDate = "Period Ending: " + txtDate
   ARptCashBal.txtDate = Now
   ARptCashBal.txtTown = GLUserName$
   ARptCashBal.GetName ReportFile$
   ARptCashBal.startrpt

   
 '  ViewPrint ReportFile$, "Cash Balance Report"
 '  KillFile ReportFile$
   Exit Sub
'PrintHeader:
'  Print #PRNFile, GLUserName$; Tab(43); "Run Date: " + Date$
'  Print #PRNFile, "Cash Balance Summary"
'  Print #PRNFile, "Ending Date: " + txtDate.Text
'  Print #PRNFile,
'  Print #PRNFile, "Account                           Balance"
'  Print #PRNFile, "-----------------------------------------"
'  Linecnt = 6
'  Return
CancelExit:
'Me.SetFocus
Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub PrintCashBal2()
  Dim PRNFile As Integer, FundCnt As Integer, cnt As Integer
  Dim ReportFile As String, ToPrint As String, CommaFmt As String
  Dim TotalFmt As String, BankGLAcct As String, cntp As Long
  Dim BankAcctBal As Double, TotalCashBal As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim NumFunds As Integer, EndDate As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, AcctRec As Integer, MaxLines As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long
  Dim GLBank(1) As GLBankRecType, BankRecLen As Integer, Linecnt As Integer
  Dim NumBanks As Integer, BankFile As Integer, Newrp As String
  BankRecLen = Len(GLBank(1))
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
'--Bank fields
'Bank(1).Deleted
'Bank(1).BankNum
'Bank(1).BankName
'Bank(1).BankAcct
'Bank(1).GLAcct
'================
  'OpenBankFile BankRecLen, BankFile, NumBanks
'''  BankFile = FreeFile
'''  Open "GLBANK.DAT" For Random Access Read Write Shared As BankFile Len = Bank
'''  NumBanks = LOF(BankFile) \ BankRecLen
  '****
  OpenBankFile BankFile, NumBanks
  If NumBanks = 0 Then
    Close BankFile
    
    Beep
    MsgBox "Error: Banks not defined. Press any key to return to menu.", vbOKOnly, "Cash Balance"
    
    Exit Sub
  End If

  ReDim FundList$(1)
  GetFundList FundList$(), NumFunds
  EndDate = DateDiff("d", "12/31/1979", txtDate)
  'Newrp = "CshBal"
  'GetRPTName Newrp
  ReportFile$ = "CashBal.txt"
  MaxLines = 53
  OpenAcctFile AcctFileNum
  NumGLAccts = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  GoSub PrintHeader
  FrmShowPctComp.Label1 = "Printing Cash Balance Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdPrint.Enabled = False
  Me.mnuOptions.Enabled = False
  For cnt = 1 To NumBanks
    Get BankFile, cnt, GLBank(1)
    
    If Not GLBank(1).Deleted Then
    'Bank(1).Deleted
    'Bank(1).BankNum
    'Bank(1).BankName
    'Bank(1).BankAcct
    'Bank(1).GLAcct
    Print #PRNFile, GLBank(1).BankName
    For FundCnt = 1 To NumFunds  'FundList$(1)
      BankGLAcct$ = QPTrim$(FundList$(FundCnt) + GLBank(1).GLAcct)
      AcctRec = AcctFind(BankGLAcct$)
      If AcctRec > 0 Then
         Get AcctFileNum, AcctRec, GLAcct
         CalcBal# = Round#(GLAcct.BegBal)
         NextTr& = GLAcct.FrstTran
         Do Until NextTr& = 0
            Get TransFileNum, NextTr&, GLTrans
            If GLTrans.TRDATE <= EndDate Then
                CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            End If
            NextTr& = GLTrans.NextTran
         Loop
         Print #PRNFile, BankGLAcct$; Tab(30);
         Print #PRNFile, Using(CommaFmt$, Str$(CalcBal#))
         TotalCashBal# = TotalCashBal# + CalcBal#
         BankAcctBal# = BankAcctBal# + CalcBal#
         Linecnt = Linecnt + 1
      End If
      If Linecnt > MaxLines Then
        Print #PRNFile, Chr$(12)
        GoSub PrintHeader
      End If
    Next

    Print #PRNFile, "Total for "; QPTrim$(GLBank(1).BankName);
    Print #PRNFile, Tab(30); Using(CommaFmt$, Str$(BankAcctBal#))
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFile, Chr$(12)
      GoSub PrintHeader
    End If
    BankAcctBal# = 0
  End If
  
    FrmShowPctComp.ShowPctComp cnt, NumBanks
    If FrmShowPctComp.Out = True Then
      Close
      'FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      Me.mnuOptions.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
   Me.cmdExit.Enabled = True
   Me.cmdPrint.Enabled = True
   Me.mnuOptions.Enabled = True
   EnableCloseButton Me.hwnd, True
   Print #PRNFile,
   Close BankFile
   Print #PRNFile, Tab(28); "--------------"
   Print #PRNFile, "Total Cash"; Tab(27); Using(CommaFmt$, Str$(TotalCashBal#))
   Print #PRNFile, Chr$(12)
   
   
   Close
  
   
   ViewPrint ReportFile$, "Cash Balance Report"
   KillFile ReportFile$
   Exit Sub
PrintHeader:
  Print #PRNFile, GLUserName$; Tab(43); "Run Date: " + Date$
  Print #PRNFile, "Cash Balance Summary"
  Print #PRNFile, "Ending Date: " + txtDate.Text
  Print #PRNFile,
  Print #PRNFile, "Account                           Balance"
  Print #PRNFile, "-----------------------------------------"
  Linecnt = 6
  Return
CancelExit:
'Me.SetFocus
Exit Sub
End Sub

