VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostAPChks 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post A/P Checks"
   ClientHeight    =   8844
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   12192
   Icon            =   "frmPostAPChks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8844
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   4392
      TabIndex        =   0
      Top             =   7272
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
      ColDesigner     =   "frmPostAPChks.frx":08CA
   End
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   2568
      Top             =   2808
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D0D0D0&
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
      Left            =   7032
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7248
      Width           =   1356
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
      Left            =   8712
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7248
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   8592
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   445
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
            TextSave        =   "11:31 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4/21/2005"
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
   Begin VB.Shape Shape4 
      Height          =   1140
      Left            =   1638
      Top             =   6912
      Width           =   8916
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type "
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
      Left            =   2112
      TabIndex        =   9
      Top             =   7344
      Width           =   2388
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   420
      Left            =   4950
      TabIndex        =   8
      Top             =   2856
      Width           =   2292
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Once Posted, The Checks or Register May NOT Be Printed Again. "
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
      Height          =   588
      Index           =   2
      Left            =   3384
      TabIndex        =   7
      Top             =   3672
      Width           =   5196
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Ok to Begin Posting or Exit to Escape Posting Procedure. "
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
      Left            =   2712
      TabIndex        =   6
      Top             =   4296
      Width           =   6780
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Before You Post, Make Sure You Have Printed A Check Register. "
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
      Index           =   0
      Left            =   2532
      TabIndex        =   5
      Top             =   3288
      Width           =   7308
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Post A/P Checks"
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
      Left            =   4548
      TabIndex        =   4
      Top             =   1464
      Width           =   3108
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3216
      Top             =   1224
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3216
      Top             =   1104
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2376
      Top             =   2664
      Width           =   7452
   End
End
Attribute VB_Name = "frmPostAPChks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Vendor As VendorRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim BadAcct As Integer
Private Sub cmdExit_Click()
  Unload frmPostAPChks
End Sub
'_________________________________________________________________
'' Used to test new report format when errors occur and need to print
'' the error logs...............
''Private Sub TempStuff()
''Dim ErrorP As Integer, ErrorFile As String
''Dim ReportFile As String, GLLogFileName As String
''
''  ErrorP = FreeFile
''  ErrorFile = "ErrorLog.PRN"
''  Open ErrorFile$ For Output As ErrorP
''
''  Print #ErrorP, "Error *** Call Software Support***"
''  Print #ErrorP, "Ledger,Vendor,Interface Update  " + Str$(55)
''  Print #ErrorP, "Error Code  " + Str(154)
''  Close
''    ARptErrorLog.GetName ErrorFile$
''    ARptErrorLog.startrpt
''
''     MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
''     ReportFile$ = "TempLog.PRN"
''     ARptErrorLog.GetName ReportFile$
''     ARptErrorLog.startrpt
''
''     MsgBox "Errors Were Found. Review GL Posting Log.", vbOKOnly, "GL Account Error"
''     GLLogFileName = "GLlog.dat"
''     ReportFile$ = "GLlog.dat"
''     ARptErrorLog.GetName ReportFile$
''     ARptErrorLog.startrpt
''End Sub
'-----------------------------------------------------------------------

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub Timer1_Timer()
 ' Label2.Visible = Not Label2.Visible
  '&H0080FFFF&
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Label2.ForeColor = &H80FFFF
    Shape3.BackColor = &HC0&
  Else
    Label2.ForeColor = &HFFFF&
    Shape3.BackColor = &H80&
  End If
  
End Sub

Private Sub cmdOk_Click()
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  Call MainLog("Begin Post APChk.")
  If rptopt = 1 Then
    ChkPostControl
  ElseIf rptopt = 2 Then
    ChkPostControl2
  End If
  MsgBox "A/P Check Posting Complete.", vbOKOnly, "Procedure Complete"
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmPostAPChks
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
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

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpPostAPChk
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Public Function ChkPostControl()
  BadAcct = 0
  APChkPost BadAcct, frmPostAPChks, False
  If BadAcct <> 0 Then
     cmdExit_Click
  Else
    APChkPost BadAcct, frmPostAPChks, True
    cmdExit_Click
  End If
End Function
Public Function ChkPostControl2()
  BadAcct = 0
  APChkPost2 BadAcct, frmPostAPChks, False
  If BadAcct <> 0 Then
     cmdExit_Click
  Else
    APChkPost2 BadAcct, frmPostAPChks, True
    cmdExit_Click
  End If
End Function

Public Sub APChkPost(Bad As Integer, Form As Form, go4it As Boolean)
  Dim Dash As String, Dash2 As String, FF As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer, ErrorP As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long, PageNum As Integer, ToPrint As String
  Dim MaxLines As Integer, CDActive As String, CashAcct As String, CDCash As String
  Dim CDDue As String, PadChars As Integer, DetPad As String, Dist As Integer
  Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
  Dim NumFunds As Integer, APDistRecLen As Integer, RecLen As Integer
  Dim PayListRecLen As Integer, TPayListFile As Integer, LedCnt As Long
  Dim TPayCnt As Integer, Pcnt As Integer, ChkCnt As Integer, Page As String
  Dim CheckDate As String, PrintFileName As String, APIFFileName As String
  Dim PrintFile2 As Integer, PrintFileName2 As String
  Dim GLIFRecLen As Integer, GLIFFile As Integer, TRec As Long
  Dim APLedgerFile As Integer, NumTran As Long, Linecnt As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, fmt As String
  Dim TotalChkAmt As Double, NextDist As Long, ThisFund As String
  Dim FundCnt As Integer, NextDistRec As Long, FirstDist As Long
  Dim Num2Write As Integer, ChkDist As Integer, ChkFDist As Integer
  Dim APAcct As String, UVoidCnt As Integer, NumIFRecs As Integer
  Dim I As Integer, RptFile As Integer, RptFileName As String, JE As String
  Dim JEDr As Double, JECr As Double, JELineCnt As Integer, TPCk As String
  Dim Text1 As String, Text2 As String, ErrorFile As String, TPInv As String
  Dim ReportFile As String, GLLogFileName As String, TempCash As String, TPDist As String
  Dim TPayList2 As TPayListType
  On Local Error GoTo ItsBroke
  FF$ = Chr$(12)
  MaxLines = 53
  ToPrint$ = Space$(78)
  Dash$ = String$(75, "-")
  Dash2$ = String$(61, "-")
  Mid$(Dash2$, 1, 7) = Space$(7)
  PageNum = 0
  TempCash$ = ""
  If go4it = True Then
    FrmShowPctComp.Label1 = "Posting Account Transactions."
  Else
    FrmShowPctComp.Label1 = "Verifying Account Transactions."
  End If
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , Form
  DoEvents
  GetAPAcct APAcct
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  fmt = "##,###,###.##"
  '--For Central Depository - number of characters to pad detail code with
  If CDActive$ = "Y" Then
    PadChars = GLDetLen - GLFundLen
    If PadChars > 0 Then
      DetPad$ = String$(PadChars, "0")
    End If
  End If

  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub

  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APLedgerRec2(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType

  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList2)
  TPayListFile = FreeFile
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    Exit Sub
  End If
  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    Get TPayListFile, Pcnt, TPayList2
    TPayArray(Pcnt).VendorRecNum = TPayList2.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList2.LedgerRecNum
  Next
  Close TPayListFile

  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  If ChkCnt = 0 Then Exit Sub
  ReDim CHKinfo(1 To ChkCnt) As CheckInfoType3
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To ChkCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  CheckDate$ = Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
  '--If we are not using a central depository, then use the
  '--cash account$ assigned with the record (glbank.dat)
  If CDActive$ <> "Y" Then
    If CHKinfo(1).Bankcode <> 0 Then
      If Exist("GLBANK.DAT") Then
        TempCash$ = GetBankGLAcct(CHKinfo(1).Bankcode)
      End If
    End If
  End If
  If TempCash$ <> "" Then
    CashAcct$ = TempCash$
  End If
  ErrorP = FreeFile
  ErrorFile = "ErrorLog.PRN"
  Open ErrorFile$ For Output As ErrorP
  PrintFile = FreeFile
  PrintFileName$ = "APCHKREG.PRN"
  Open PrintFileName$ For Output As PrintFile
  PrintFile2 = FreeFile
  PrintFileName2$ = "APCHKREG2.PRN"
  Open PrintFileName2$ For Output As PrintFile2

  APIFFileName$ = "APCHKIF.DAT"
  If go4it = False Then
    KillFile APIFFileName$
  End If
  ReDim GLifRec(1) As GLTransRecType
  GLIFRecLen = Len(GLifRec(1))
  GLIFFile = FreeFile
  Open APIFFileName$ For Random As GLIFFile Len = GLIFRecLen
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  If go4it = True Then
    'GoSub DoChkRegHeader
  End If
  '--for each check
  For cnt = 1 To ChkCnt
    FrmShowPctComp.ShowPctComp cnt, ChkCnt

    '--if the check was'nt voided
    If CHKinfo(cnt).VoidFlag = False Then
      '--Create a list to hold the check's distribution amounts by fund
      ReDim ThisChkFunds(1 To NumFunds) As Double

      '--update vendor's pointer to their last transaction
      Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
      NumTran& = NumTran& + 1
      If Vendor.LastTran > 0 Then
        Get APLedgerFile, Vendor.LastTran, APLedgerRec(1)
        APLedgerRec(1).NextTrans = NumTran&
        If go4it = True Then
          Put APLedgerFile, Vendor.LastTran, APLedgerRec(1)
        End If
      End If
      Vendor.LastTran = NumTran&
      If go4it = True Then
        Put VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
      End If
      '---Post the check to the apledger file
      APLedgerRec2(1).VIN = Vendor.VIN
      APLedgerRec2(1).VendorCode = Vendor.vnum
      APLedgerRec2(1).VRecNum = CHKinfo(cnt).VendorRecNum
      APLedgerRec2(1).TRDATE = CHKinfo(cnt).chkdate
      APLedgerRec2(1).DOCNum = Str$(CHKinfo(cnt).LastChk)
      APLedgerRec2(1).TRCode = 3
      APLedgerRec2(1).PAYCODE = 3
      APLedgerRec2(1).Amt = CHKinfo(cnt).ChkAmt
      APLedgerRec2(1).NextTrans = 0
      APLedgerRec2(1).Bankcode = CHKinfo(cnt).Bankcode
      If go4it = True Then
        Put APLedgerFile, NumTran&, APLedgerRec2(1)
      End If
      'write out cent dep cash as a single amount
      '--IF Cent Dep write out check here
      If CDActive$ = "Y" Then
        '--Credit Central Depository Cash
        If go4it = False Then
          TRec& = TRec& + 1
          GLifRec(1).TRDATE = CHKinfo(1).chkdate
          GLifRec(1).AcctNum = CDCash$
          GLifRec(1).Desc = Vendor.VNAME
          GLifRec(1).LDesc = "To Bank Code- " + Str$(CHKinfo(cnt).Bankcode)
          GLifRec(1).DrAmt = 0
          GLifRec(1).CrAmt = CHKinfo(cnt).ChkAmt
          GLifRec(1).Src = "CK" + Format$(Now, "mmddyy")
          GLifRec(1).Ref = Str$(CHKinfo(cnt).LastChk)
          Put GLIFFile, TRec&, GLifRec(1)
        End If
      End If

      '--check lines printed to see if we need a page break
'      If Linecnt > MaxLines Then
'        Print #PrintFile, FF$
'        GoSub DoChkRegHeader
'      End If
      TPCk$ = ""
      '--print this check to the report
      RSet ChkRegInfo(1).ChkNum = Str$(CHKinfo(cnt).LastChk)
      RSet ChkRegInfo(1).chkdate = Format(DateAdd("d", (CHKinfo(cnt).chkdate), "12-31-1979"), "mm/dd/yyyy")
      RSet ChkRegInfo(1).ChkAmt = Using(fmt$, Str$(CHKinfo(cnt).ChkAmt))
      LSet ChkRegInfo(1).VendName = Vendor.VNAME
      'Print #PrintFile, Dash$
      TPCk$ = ChkRegInfo(1).ChkNum + "~" + ChkRegInfo(1).chkdate
      TPCk$ = TPCk$ + "~" + ChkRegInfo(1).VendName + "~" + ChkRegInfo(1).ChkAmt
      'Linecnt = Linecnt + 2

      '--For each invoice being paid by this check...
      For LedCnt& = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast

        '--update the invoice record with paid check num and date
        Get APLedgerFile, TPayArray(LedCnt&).LedgerRecNum, APLedgerRec(1)
        APLedgerRec(1).PDCheckDate = CHKinfo(cnt).chkdate
        APLedgerRec(1).PDCheckNum = CHKinfo(cnt).LastChk
        APLedgerRec(1).PAYCODE = 3
        'add this 7/29/02 to keep the bankcode for void chk stuff
        APLedgerRec(1).Bankcode = CHKinfo(cnt).Bankcode
        If go4it = True Then
          Put APLedgerFile, TPayArray(LedCnt&).LedgerRecNum, APLedgerRec(1)
        End If
        '--print the invoice header
        'GoSub InvHeader
        '--check for a page break
        'If Linecnt > MaxLines Then
          'Print #PrintFile, FF$
          'GoSub DoChkRegHeader
'#############!!!!!!!!!!!!!!!!!Put next 3 lines in to show continuation of check
          'Print #PrintFile, Dash$
          'Print #PrintFile, ChkRegInfo(1).ChkNum; " Continued"
          'Linecnt = Linecnt + 2
          'GoSub InvHeader
        'End If
        TPInv$ = ""
        '--Print the invoice to the report
        RSet LedInfo(1).LedDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
        LSet LedInfo(1).LedInvNum = APLedgerRec(1).DOCNum
        RSet LedInfo(1).InvAmt = Using(fmt$, Str$(APLedgerRec(1).Amt))
        TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
        TPInv$ = LedInfo(1).LedDate + "~" + LedInfo(1).LedInvNum + "~" + LedInfo(1).InvAmt
        'Linecnt = Linecnt + 2

        '--Print the distribution header
        'GoSub DistHeader

        '--print the distributions of the invoice
        NextDist& = APLedgerRec(1).FrstDist
        Do Until NextDist& = 0
          Get APDistFile, NextDist&, APDistRec(1)
'          If Linecnt > MaxLines Then
'            Print #PrintFile, FF$
'            GoSub DoChkRegHeader
''#############!!!!!!!!!!!!!!!!!Put next 3 lines in to show continuation of check
'          Print #PrintFile, Dash$
'          Print #PrintFile, ChkRegInfo(1).ChkNum; " Continued"
'          Linecnt = Linecnt + 2
'            GoSub DistHeader
'          End If
          TPDist$ = ""
          RSet DistInfo(1).DistAcct = APDistRec(1).DistAcctNum
          RSet DistInfo(1).DistAmt = Using(fmt$, Str$(APDistRec(1).DistAmt))
          TPDist$ = DistInfo(1).DistAcct + "~" + DistInfo(1).DistAmt
          'Linecnt = Linecnt + 1
          ToPrint$ = TPCk$ + "~" + TPInv$ + "~" + TPDist
          Print #PrintFile, ToPrint$
          ToPrint$ = ""
          '--Summarize the distributions by fund
          ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
          For FundCnt = 1 To NumFunds
            If ThisFund$ = FundList$(FundCnt) Then
              '--Update grand total by fund
              FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
              '--Update check totals by fund
              ThisChkFunds(FundCnt) = Round(ThisChkFunds(FundCnt) + APDistRec(1).DistAmt)
              Exit For
            End If
          Next

          NextDist& = APDistRec(1).NextDist
        Loop '--distribution printing
      Next  '--Invoice being paid

      'Print #PrintFile,

      '--get the record number for a new distribution record
      NextDistRec& = (LOF(APDistFile) \ APDistRecLen) + 1
      FirstDist& = NextDistRec&

      '--see how many distributions we need to write
      Num2Write = 0
      For ChkDist = 1 To NumFunds
'IF ThisChkFunds(ChkDist) > 0 THEN  'what if we have negative entries?
        If ThisChkFunds(ChkDist) <> 0 Then  'what if we have negative entries?
          Num2Write = Num2Write + 1
        End If
      Next

      '--Write out check distribution by fund to the apdist.dat file
      ReDim NewDistRec(1) As APDistRecType
      For ChkFDist = 1 To NumFunds
'IF ThisChkFunds(ChkFDist) > 0 THEN
        If ThisChkFunds(ChkFDist) <> 0 Then
          Dist = Dist + 1
          If Dist = Num2Write Then
            NewDistRec(1).APLedgerRec = NumTran&
            NewDistRec(1).DistAcctNum = FundList$(ChkFDist) + APAcct$
            NewDistRec(1).DistAmt = ThisChkFunds(ChkFDist)
            NewDistRec(1).NextDist = 0
            If go4it = True Then
              Put APDistFile, NextDistRec&, NewDistRec(1)
            End If
            Dist = 0
            Exit For
          Else
            NewDistRec(1).APLedgerRec = NumTran&
            NewDistRec(1).DistAcctNum = FundList$(ChkFDist) + APAcct$
            NewDistRec(1).DistAmt = ThisChkFunds(ChkFDist)
            NewDistRec(1).NextDist = NextDistRec& + 1
            If go4it = True Then
              Put APDistFile, NextDistRec&, NewDistRec(1)
            End If
            NextDistRec& = NextDistRec& + 1
          End If
        End If
      Next

      '--update ledger rec pointers to first and last dist here
      APLedgerRec2(1).FrstDist = FirstDist&
      APLedgerRec2(1).LastDist = (FirstDist& + Num2Write) - 1
      If go4it = True Then
        Put APLedgerFile, NumTran&, APLedgerRec2(1)
      End If
      '--Write out the GL entries to update cash and a/p
      '--creates one entry per check per fund
      For ChkFDist = 1 To NumFunds
'IF ThisChkFunds(ChkFDist) > 0 THEN
        If ThisChkFunds(ChkFDist) <> 0 Then
          '--Credit Cash or Due to Central Depository
          If go4it = False Then
            TRec& = TRec& + 1
            GLifRec(1).Src = "CK" + Format$(Now, "mmddyy")
            GLifRec(1).AcctNum = FundList$(ChkFDist) + CashAcct$
            GLifRec(1).TRDATE = CHKinfo(1).chkdate
            GLifRec(1).Desc = Vendor.VNAME
            GLifRec(1).LDesc = CHKinfo(1).Bankcode
            GLifRec(1).Ref = Str$(APLedgerRec(1).PDCheckNum)
            GLifRec(1).CrAmt = ThisChkFunds(ChkFDist)
            GLifRec(1).DrAmt = 0
            Put GLIFFile, TRec&, GLifRec(1)
            '--Debit A/P
            TRec& = TRec& + 1
            GLifRec(1).AcctNum = FundList$(ChkFDist) + APAcct$
            GLifRec(1).CrAmt = 0
            GLifRec(1).DrAmt = ThisChkFunds(ChkFDist)
            Put GLIFFile, TRec&, GLifRec(1)
            '--update Central Depository
            If CDActive$ = "Y" Then
            '--Debit Central Depository Due From
              TRec& = TRec& + 1
              GLifRec(1).AcctNum = CDDue$ + FundList$(ChkFDist) + DetPad$
              GLifRec(1).DrAmt = ThisChkFunds(ChkFDist)
              GLifRec(1).CrAmt = 0
              Put GLIFFile, TRec&, GLifRec(1)
            End If
          End If
         End If
        Next

      Else
        UVoidCnt = UVoidCnt + 1
    End If
  Next 'check

  '--Finish Check Detail Report
 ' Print #PrintFile, Dash$
 ' Print #PrintFile, FF$
 ' Linecnt = 1
  'PageNum = 0
'  GoSub DoFundHeader
  For cnt = 1 To NumFunds
'IF FundAmts(Cnt) > 0 THEN
    If FundAmts(cnt) <> 0 Then
      Print #PrintFile2, Using("####", Str$(Val(FundList$(cnt)))) + "~" + Using(fmt$, Str$(FundAmts(cnt)))
      'Linecnt = Linecnt + 1
    End If
  Next
  
  Close

  '--Create the A/P Check Journal Entry report
  'If go4it = True Then
  GLIFFile = FreeFile
  Open "APCHKIF.DAT" For Random As GLIFFile Len = GLIFRecLen
  NumIFRecs = LOF(GLIFFile) \ GLIFRecLen

  ReDim SorTtrans(NumIFRecs) As GLTransRecType
  For I = 1 To NumIFRecs
    Get GLIFFile, I, GLifRec(1)
    SorTtrans(I) = GLifRec(1)
  Next
  low = LBound(SorTtrans)
  High = UBound(SorTtrans)
  QCkSort2 SorTtrans(), low, High


  '--print the journal entry report
  RptFile = FreeFile
  RptFileName$ = "APCHKJE.PRN"
  Open RptFileName$ For Output As RptFile
  JE$ = Space$(80)
  'GoSub JEHeader

  For I = 1 To NumIFRecs
    '--Print Entry
    LSet JE$ = ""
    JE$ = SorTtrans(I).AcctNum
    JE$ = JE$ + "~" + Left$(SorTtrans(I).Desc, 19)
    JE$ = JE$ + "~" + SorTtrans(I).Ref
    JE$ = JE$ + "~" + Using(fmt$, Str$(SorTtrans(I).DrAmt))
    JE$ = JE$ + "~" + Using(fmt$, Str$(SorTtrans(I).CrAmt))
    Print #RptFile, JE$

    JEDr# = JEDr# + SorTtrans(I).DrAmt
    JECr# = JECr# + SorTtrans(I).CrAmt

'    JELineCnt = JELineCnt + 1
'    If JELineCnt > 55 Then
'       Print #RptFile, FF$
'       GoSub JEHeader
'    End If
  Next

'  Print #RptFile,
'  Print #RptFile, "Journal Entry Totals";
'  Print #RptFile, Tab(49); Using(fmt$, Str$(JEDr#));
'  Print #RptFile, Tab(63); Using(fmt$, Str$(JECr#))
'  Print #RptFile, FF$
' End If
  Close

  '--Post and clean up
  If BadAcct > 0 Then
    Call MainLog("Errors APChk Post.")
    MsgBox "Errors Were Found, Ok to View", vbOKOnly, "Errors"
    ARptErrorLog.GetName ErrorFile$
    ARptErrorLog.startrpt
    
    'ViewPrint "ErrorFile", "Error Message"
    GoTo ExitPost
  End If
   Post2GL APIFFileName$, BadAcct, frmAPChkProcessMenu, go4it
   If BadAcct > 0 Then
    If go4it = False Then
     Call MainLog("APChk Post Errors - Procedure Stopped.")
     MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
     ReportFile$ = "TempLog.PRN"
     ARptErrorLog.GetName ReportFile$
     ARptErrorLog.startrpt

     'ViewPrint ReportFile$, "Error Log"
     frmCitiCancel.Show
     Unload frmPostAPChks
     Unload frmAPChkProcessMenu
     GoTo ExitPost
    Else
     Call MainLog("APChk Post Errors 2nd gothru GL Acct Error.")
     MsgBox "Errors Were Found. Review GL Posting Log.", vbOKOnly, "GL Account Error"
     GLLogFileName = "GLlog.dat"
     ReportFile$ = "GLlog.dat"
     ARptErrorLog.GetName ReportFile$
     ARptErrorLog.startrpt
     'ViewPrint ReportFile$, "Posting Log"
   End If
   End If
   PostCHKDAT "APCHKINF.DAT", frmAPChkProcessMenu, go4it
   If go4it = True Then
    KillFile APIFFileName$
    KillFile "APCHKINF.DAT"
    KillFile "TPAYLIST.LST"
    KillFile "TPAYLIST2.LST"
    KillFile "TPAYLISTD.LST"
    KillFile "TPAYNot.LST"
   End If
    Erase FundList$, FundAmts, APLedgerRec, APLedgerRec2, APDistRec
    Erase LedInfo, DistInfo, ChkRegInfo, CHKinfo
    Erase GLifRec
  If go4it = True Then
  Load frmLoadingRpt
  Title$ = "A/P Check Register"
  ARptPostAPCks.txtTown.Caption = GLUserName$
  ARptPostAPCks.txtDate.Caption = Now
  ARptPostAPCks.Label1.Caption = Title$
  ARptPostAPCks.txtTotCks = Using("####", Str$(ChkCnt - UVoidCnt))
  ARptPostAPCks.txtTotalAmt = Using("$###,###,###.##", Str$(TotalChkAmt#))
  ARptPostAPCks.GetName PrintFileName$, PrintFileName2$
  ARptPostAPCks.startrpt

  'ViewPrint PrintFileName$, title$
  Title$ = "A/P Check Journal Entries"
  'ViewPrint RptFileName$, title$
  ARptPostChkEntries.txtTown.Caption = GLUserName$
  ARptPostChkEntries.txtDate.Caption = Now
  ARptPostChkEntries.Label1.Caption = Title$
  ARptPostChkEntries.totDeb = Using(fmt$, Str$(JEDr#))
  ARptPostChkEntries.totCred = Using(fmt$, Str$(JECr#))
  ARptPostChkEntries.GetName RptFileName$
  ARptPostChkEntries.startrpt
  Call MainLog("APChk Post Complete.")
  End If
ExitPost:
Exit Sub

'JEHeader:
'  Print #RptFile, "A/P Check Journal Entries"
'  Print #RptFile, "Posting Date: " + CheckDate$
'  Print #RptFile,
'  Mid$(JE$, 1) = "Account Number"
'  Mid$(JE$, 20) = "Description"
'  Mid$(JE$, 40) = "Reference"
'  Mid$(JE$, 49) = "     Debit"
'  Mid$(JE$, 63) = "    Credit"
'  Print #RptFile, JE$
'  LSet JE$ = ""
'  JELineCnt = 4
'Return


'DoFundHeader:
'  PageNum = PageNum + 1
'  Page$ = Using("###", Str$(PageNum))
'
'  Text1$ = "A/P Check Register Detail"
'  Text2$ = "Total Checks Issued        "
'  Print #PrintFile, Text1$
'  Print #PrintFile, Dash$
'  Print #PrintFile,
'  Print #PrintFile, Text2$; Using("####", Str$(ChkCnt - UVoidCnt))
'  Print #PrintFile, "           Totaling  "; Using("$###,###,###.##", Str$(TotalChkAmt#))
'  Print #PrintFile,
'  Print #PrintFile, " By Fund:"
'  Linecnt = 7
'Return

'DoChkRegHeader:
'  PageNum = PageNum + 1
'  Page$ = Using("###", Str$(PageNum))
'  Text1$ = "A/P Check Register Summary                                        Page:" + Page$
'  Print #PrintFile, Text1$
'  Print #PrintFile,
'  Print #PrintFile, " Check No.   Check Date   Vendor Name                         Check Amount"
'  Linecnt = 3
'Return

'InvHeader:
'  Print #PrintFile, Dash2$
'  Print #PrintFile, "       Inv Date    Inv Num                            Inv"
'  Linecnt = Linecnt + 2
'Return

'DistHeader:
'  Print #PrintFile, "                   Dist Acct                         Dist"
'  Linecnt = Linecnt + 1
'Return
ItsBroke:
  BadAcct = BadAcct + 1
  Print #ErrorP, "Error *** Call Software Support***"
  Print #ErrorP, "Ledger,Vendor,Interface Update  " + Str$(cnt)
  Print #ErrorP, "Error Code  " + Str(Err.Number)
  Resume Next
End Sub
Public Sub APChkPost2(Bad As Integer, Form As Form, go4it As Boolean)
  Dim Dash As String, Dash2 As String, FF As String, ChkInfoRecLen As Integer
  Dim VCnt As Integer, cnt As Integer, ChkinfoFile As Integer, ErrorP As Integer
  Dim VendorFile As Integer, NumVRecs As Long, PrintFile As Integer
  Dim Cnt2 As Long, TCheckAmt As Double, Title As String, Temp As Integer
  Dim low As Long, High As Long, PageNum As Integer, ToPrint As String
  Dim MaxLines As Integer, CDActive As String, CashAcct As String, CDCash As String
  Dim CDDue As String, PadChars As Integer, DetPad As String, Dist As Integer
  Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
  Dim NumFunds As Integer, APDistRecLen As Integer, RecLen As Integer
  Dim PayListRecLen As Integer, TPayListFile As Integer, LedCnt As Long
  Dim TPayCnt As Integer, Pcnt As Integer, ChkCnt As Integer, Page As String
  Dim CheckDate As String, PrintFileName As String, APIFFileName As String
  Dim GLIFRecLen As Integer, GLIFFile As Integer, TRec As Long
  Dim APLedgerFile As Integer, NumTran As Long, Linecnt As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, fmt As String
  Dim TotalChkAmt As Double, NextDist As Long, ThisFund As String
  Dim FundCnt As Integer, NextDistRec As Long, FirstDist As Long
  Dim Num2Write As Integer, ChkDist As Integer, ChkFDist As Integer
  Dim APAcct As String, UVoidCnt As Integer, NumIFRecs As Integer
  Dim I As Integer, RptFile As Integer, RptFileName As String, JE As String
  Dim JEDr As Double, JECr As Double, JELineCnt As Integer
  Dim Text1 As String, Text2 As String, ErrorFile As String
  Dim ReportFile As String, GLLogFileName As String, TempCash As String
  Dim TPayList2 As TPayListType
  On Local Error GoTo ItsBroke
  FF$ = Chr$(12)
  MaxLines = 53
  ToPrint$ = Space$(78)
  Dash$ = String$(75, "-")
  Dash2$ = String$(61, "-")
  Mid$(Dash2$, 1, 7) = Space$(7)
  PageNum = 0
  TempCash$ = ""
  If go4it = True Then
    FrmShowPctComp.Label1 = "Posting Account Transactions."
  Else
    FrmShowPctComp.Label1 = "Verifying Account Transactions."
  End If
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , Form
  DoEvents
  GetAPAcct APAcct
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  fmt = "##,###,###.##"
  '--For Central Depository - number of characters to pad detail code with
  If CDActive$ = "Y" Then
    PadChars = GLDetLen - GLFundLen
    If PadChars > 0 Then
      DetPad$ = String$(PadChars, "0")
    End If
  End If

  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub

  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APLedgerRec2(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType
  ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType

  DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))
  PayListRecLen = Len(TPayList2)
  TPayListFile = FreeFile
  Open "TPAYLIST2.LST" For Random Shared As TPayListFile Len = PayListRecLen
  TPayCnt = LOF(TPayListFile) \ 6
  If TPayCnt = 0 Then
    Exit Sub
  End If
  ReDim TPayArray(1 To TPayCnt) As TPayListType
  For Pcnt = 1 To TPayCnt
    Get TPayListFile, Pcnt, TPayList2
    TPayArray(Pcnt).VendorRecNum = TPayList2.VendorRecNum
    TPayArray(Pcnt).LedgerRecNum = TPayList2.LedgerRecNum
  Next
  Close TPayListFile

  ChkinfoFile = FreeFile
  ReDim CHKinfo(1 To 1) As CheckInfoType3
  ChkInfoRecLen = Len(CHKinfo(1))
  ChkCnt = (FileSize("APCHKINF.DAT") \ ChkInfoRecLen)
  If ChkCnt = 0 Then Exit Sub
  ReDim CHKinfo(1 To ChkCnt) As CheckInfoType3
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkInfoRecLen
  For Temp = 1 To ChkCnt
    Get ChkinfoFile, Temp, CHKinfo(Temp)
  Next
  CheckDate$ = Format(DateAdd("d", (CHKinfo(1).chkdate), "12-31-1979"), "mm/dd/yyyy")
  '--If we are not using a central depository, then use the
  '--cash account$ assigned with the record (glbank.dat)
  If CDActive$ <> "Y" Then
    If CHKinfo(1).Bankcode <> 0 Then
      If Exist("GLBANK.DAT") Then
        TempCash$ = GetBankGLAcct(CHKinfo(1).Bankcode)
      End If
    End If
  End If
  If TempCash$ <> "" Then
    CashAcct$ = TempCash$
  End If
  ErrorP = FreeFile
  ErrorFile = "ErrorLog.PRN"
  Open ErrorFile$ For Output As ErrorP
  PrintFile = FreeFile
  PrintFileName$ = "APCHKREG.PRN"
  Open PrintFileName$ For Output As PrintFile
  APIFFileName$ = "APCHKIF.DAT"
  If go4it = False Then
    KillFile APIFFileName$
  End If
  ReDim GLifRec(1) As GLTransRecType
  GLIFRecLen = Len(GLifRec(1))
  GLIFFile = FreeFile
  Open APIFFileName$ For Random As GLIFFile Len = GLIFRecLen
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  If go4it = True Then
    GoSub DoChkRegHeader
  End If
  '--for each check
  For cnt = 1 To ChkCnt
    FrmShowPctComp.ShowPctComp cnt, ChkCnt

    '--if the check was'nt voided
    If CHKinfo(cnt).VoidFlag = False Then
      '--Create a list to hold the check's distribution amounts by fund
      ReDim ThisChkFunds(1 To NumFunds) As Double

      '--update vendor's pointer to their last transaction
      Get VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
      NumTran& = NumTran& + 1
      If Vendor.LastTran > 0 Then
        Get APLedgerFile, Vendor.LastTran, APLedgerRec(1)
        APLedgerRec(1).NextTrans = NumTran&
        If go4it = True Then
          Put APLedgerFile, Vendor.LastTran, APLedgerRec(1)
        End If
      End If
      Vendor.LastTran = NumTran&
      If go4it = True Then
        Put VendorFile, CHKinfo(cnt).VendorRecNum, Vendor
      End If
      '---Post the check to the apledger file
      APLedgerRec2(1).VIN = Vendor.VIN
      APLedgerRec2(1).VendorCode = Vendor.vnum
      APLedgerRec2(1).VRecNum = CHKinfo(cnt).VendorRecNum
      APLedgerRec2(1).TRDATE = CHKinfo(cnt).chkdate
      APLedgerRec2(1).DOCNum = Str$(CHKinfo(cnt).LastChk)
      APLedgerRec2(1).TRCode = 3
      APLedgerRec2(1).PAYCODE = 3
      APLedgerRec2(1).Amt = CHKinfo(cnt).ChkAmt
      APLedgerRec2(1).NextTrans = 0
      APLedgerRec2(1).Bankcode = CHKinfo(cnt).Bankcode
      If go4it = True Then
        Put APLedgerFile, NumTran&, APLedgerRec2(1)
      End If
      'write out cent dep cash as a single amount
      '--IF Cent Dep write out check here
      If CDActive$ = "Y" Then
        '--Credit Central Depository Cash
        If go4it = False Then
          TRec& = TRec& + 1
          GLifRec(1).TRDATE = CHKinfo(1).chkdate
          GLifRec(1).AcctNum = CDCash$
          GLifRec(1).Desc = Vendor.VNAME
          GLifRec(1).LDesc = CHKinfo(cnt).Bankcode
          GLifRec(1).DrAmt = 0
          GLifRec(1).CrAmt = CHKinfo(cnt).ChkAmt
          GLifRec(1).Src = "CK" + Format$(Now, "mmddyy")
          GLifRec(1).Ref = Str$(CHKinfo(cnt).LastChk)
          Put GLIFFile, TRec&, GLifRec(1)
        End If
      End If

      '--check lines printed to see if we need a page break
      If Linecnt > MaxLines Then
        Print #PrintFile, FF$
        GoSub DoChkRegHeader
      End If

      '--print this check to the report
      RSet ChkRegInfo(1).ChkNum = Str$(CHKinfo(cnt).LastChk)
      RSet ChkRegInfo(1).chkdate = Format(DateAdd("d", (CHKinfo(cnt).chkdate), "12-31-1979"), "mm/dd/yyyy")
      RSet ChkRegInfo(1).ChkAmt = Using(fmt$, Str$(CHKinfo(cnt).ChkAmt))
      LSet ChkRegInfo(1).VendName = Vendor.VNAME
      Print #PrintFile, Dash$
      Print #PrintFile, ChkRegInfo(1).ChkNum; ChkRegInfo(1).chkdate; "   ";
      Print #PrintFile, ChkRegInfo(1).VendName; "    "; ChkRegInfo(1).ChkAmt
      Linecnt = Linecnt + 2

      '--For each invoice being paid by this check...
      For LedCnt& = CHKinfo(cnt).ListFirst To CHKinfo(cnt).ListLast

        '--update the invoice record with paid check num and date
        Get APLedgerFile, TPayArray(LedCnt&).LedgerRecNum, APLedgerRec(1)
        APLedgerRec(1).PDCheckDate = CHKinfo(cnt).chkdate
        APLedgerRec(1).PDCheckNum = CHKinfo(cnt).LastChk
        APLedgerRec(1).PAYCODE = 3
        'add this 7/29/02 to keep the bankcode for void chk stuff
        APLedgerRec(1).Bankcode = CHKinfo(cnt).Bankcode
        If go4it = True Then
          Put APLedgerFile, TPayArray(LedCnt&).LedgerRecNum, APLedgerRec(1)
        End If
        '--print the invoice header
        GoSub InvHeader
        '--check for a page break
        If Linecnt > MaxLines Then
          Print #PrintFile, FF$
          GoSub DoChkRegHeader
'#############!!!!!!!!!!!!!!!!!Put next 3 lines in to show continuation of check
          Print #PrintFile, Dash$
          Print #PrintFile, ChkRegInfo(1).ChkNum; " Continued"
          Linecnt = Linecnt + 2
          GoSub InvHeader
        End If

        '--Print the invoice to the report
        RSet LedInfo(1).LedDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
        LSet LedInfo(1).LedInvNum = APLedgerRec(1).DOCNum
        RSet LedInfo(1).InvAmt = Using(fmt$, Str$(APLedgerRec(1).Amt))
        TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
        Print #PrintFile, LedInfo(1).LedDate; "  "; LedInfo(1).LedInvNum; LedInfo(1).InvAmt
        Linecnt = Linecnt + 2

        '--Print the distribution header
        GoSub DistHeader

        '--print the distributions of the invoice
        NextDist& = APLedgerRec(1).FrstDist
        Do Until NextDist& = 0
          Get APDistFile, NextDist&, APDistRec(1)
          If Linecnt > MaxLines Then
            Print #PrintFile, FF$
            GoSub DoChkRegHeader
'#############!!!!!!!!!!!!!!!!!Put next 3 lines in to show continuation of check
          Print #PrintFile, Dash$
          Print #PrintFile, ChkRegInfo(1).ChkNum; " Continued"
          Linecnt = Linecnt + 2
            GoSub DistHeader
          End If
          RSet DistInfo(1).DistAcct = APDistRec(1).DistAcctNum
          RSet DistInfo(1).DistAmt = Using(fmt$, Str$(APDistRec(1).DistAmt))
          Print #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; "   "; DistInfo(1).DistAmt
          Linecnt = Linecnt + 1

          '--Summarize the distributions by fund
          ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
          For FundCnt = 1 To NumFunds
            If ThisFund$ = FundList$(FundCnt) Then
              '--Update grand total by fund
              FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
              '--Update check totals by fund
              ThisChkFunds(FundCnt) = Round(ThisChkFunds(FundCnt) + APDistRec(1).DistAmt)
              Exit For
            End If
          Next

          NextDist& = APDistRec(1).NextDist
        Loop '--distribution printing
      Next  '--Invoice being paid

      Print #PrintFile,

      '--get the record number for a new distribution record
      NextDistRec& = (LOF(APDistFile) \ APDistRecLen) + 1
      FirstDist& = NextDistRec&

      '--see how many distributions we need to write
      Num2Write = 0
      For ChkDist = 1 To NumFunds
'IF ThisChkFunds(ChkDist) > 0 THEN  'what if we have negative entries?
        If ThisChkFunds(ChkDist) <> 0 Then  'what if we have negative entries?
          Num2Write = Num2Write + 1
        End If
      Next

      '--Write out check distribution by fund to the apdist.dat file
      ReDim NewDistRec(1) As APDistRecType
      For ChkFDist = 1 To NumFunds
'IF ThisChkFunds(ChkFDist) > 0 THEN
        If ThisChkFunds(ChkFDist) <> 0 Then
          Dist = Dist + 1
          If Dist = Num2Write Then
            NewDistRec(1).APLedgerRec = NumTran&
            NewDistRec(1).DistAcctNum = FundList$(ChkFDist) + APAcct$
            NewDistRec(1).DistAmt = ThisChkFunds(ChkFDist)
            NewDistRec(1).NextDist = 0
            If go4it = True Then
              Put APDistFile, NextDistRec&, NewDistRec(1)
            End If
            Dist = 0
            Exit For
          Else
            NewDistRec(1).APLedgerRec = NumTran&
            NewDistRec(1).DistAcctNum = FundList$(ChkFDist) + APAcct$
            NewDistRec(1).DistAmt = ThisChkFunds(ChkFDist)
            NewDistRec(1).NextDist = NextDistRec& + 1
            If go4it = True Then
              Put APDistFile, NextDistRec&, NewDistRec(1)
            End If
            NextDistRec& = NextDistRec& + 1
          End If
        End If
      Next

      '--update ledger rec pointers to first and last dist here
      APLedgerRec2(1).FrstDist = FirstDist&
      APLedgerRec2(1).LastDist = (FirstDist& + Num2Write) - 1
      If go4it = True Then
        Put APLedgerFile, NumTran&, APLedgerRec2(1)
      End If
      '--Write out the GL entries to update cash and a/p
      '--creates one entry per check per fund
      For ChkFDist = 1 To NumFunds
'IF ThisChkFunds(ChkFDist) > 0 THEN
        If ThisChkFunds(ChkFDist) <> 0 Then
          '--Credit Cash or Due to Central Depository
          If go4it = False Then
            TRec& = TRec& + 1
            GLifRec(1).Src = "CK" + Format$(Now, "mmddyy")
            GLifRec(1).AcctNum = FundList$(ChkFDist) + CashAcct$
            GLifRec(1).TRDATE = CHKinfo(1).chkdate
            GLifRec(1).Desc = Vendor.VNAME
            GLifRec(1).LDesc = CHKinfo(1).Bankcode
            GLifRec(1).Ref = Str$(APLedgerRec(1).PDCheckNum)
            GLifRec(1).CrAmt = ThisChkFunds(ChkFDist)
            GLifRec(1).DrAmt = 0
            Put GLIFFile, TRec&, GLifRec(1)
            '--Debit A/P
            TRec& = TRec& + 1
            GLifRec(1).AcctNum = FundList$(ChkFDist) + APAcct$
            GLifRec(1).CrAmt = 0
            GLifRec(1).DrAmt = ThisChkFunds(ChkFDist)
            Put GLIFFile, TRec&, GLifRec(1)
            '--update Central Depository
            If CDActive$ = "Y" Then
            '--Debit Central Depository Due From
              TRec& = TRec& + 1
              GLifRec(1).AcctNum = CDDue$ + FundList$(ChkFDist) + DetPad$
              GLifRec(1).DrAmt = ThisChkFunds(ChkFDist)
              GLifRec(1).CrAmt = 0
              Put GLIFFile, TRec&, GLifRec(1)
            End If
          End If
         End If
        Next

      Else
        UVoidCnt = UVoidCnt + 1
    End If
  Next 'check

  '--Finish Check Detail Report
  Print #PrintFile, Dash$
  Print #PrintFile, FF$
  Linecnt = 1
  PageNum = 0
  GoSub DoFundHeader
  For cnt = 1 To NumFunds
'IF FundAmts(Cnt) > 0 THEN
    If FundAmts(cnt) <> 0 Then
      Print #PrintFile, "      Fund "; Using("####", Str$(Val(FundList$(cnt)))); "  Amt  "; Using(fmt$, Str$(FundAmts(cnt)))
      Linecnt = Linecnt + 1
    End If
  Next
  Print #PrintFile, FF$
  Close

  '--Create the A/P Check Journal Entry report
  'If go4it = True Then
  GLIFFile = FreeFile
  Open "APCHKIF.DAT" For Random As GLIFFile Len = GLIFRecLen
  NumIFRecs = LOF(GLIFFile) \ GLIFRecLen

  ReDim SorTtrans(NumIFRecs) As GLTransRecType
  For I = 1 To NumIFRecs
    Get GLIFFile, I, GLifRec(1)
    SorTtrans(I) = GLifRec(1)
  Next
  low = LBound(SorTtrans)
  High = UBound(SorTtrans)
  QCkSort2 SorTtrans(), low, High


  '--print the journal entry report
  RptFile = FreeFile
  RptFileName$ = "APCHKJE.PRN"
  Open RptFileName$ For Output As RptFile
  JE$ = Space$(80)
  GoSub JEHeader
  

  For I = 1 To NumIFRecs
    '--Print Entry
    LSet JE$ = ""
    Mid$(JE$, 1) = SorTtrans(I).AcctNum
    Mid$(JE$, 20) = Left$(SorTtrans(I).Desc, 19)
    Mid$(JE$, 40) = SorTtrans(I).Ref
    Mid$(JE$, 49) = Using(fmt$, Str$(SorTtrans(I).DrAmt))
    Mid$(JE$, 63) = Using(fmt$, Str$(SorTtrans(I).CrAmt))
    Print #RptFile, JE$

    JEDr# = JEDr# + SorTtrans(I).DrAmt
    JECr# = JECr# + SorTtrans(I).CrAmt

    JELineCnt = JELineCnt + 1
    If JELineCnt > 55 Then
       Print #RptFile, FF$
       GoSub JEHeader
    End If
  Next

  Print #RptFile,
  Print #RptFile, "Journal Entry Totals";
  Print #RptFile, Tab(49); Using(fmt$, Str$(JEDr#));
  Print #RptFile, Tab(63); Using(fmt$, Str$(JECr#))
  Print #RptFile, FF$
' End If
  Close

  '--Post and clean up
  If BadAcct > 0 Then
    Call MainLog("Errors APChk Post.")
    MsgBox "Errors Were Found, Ok to View", vbOKOnly, "Errors"
    ViewPrint "ErrorFile", "Error Message"
    GoTo ExitPost
  End If
   Post2GL APIFFileName$, BadAcct, frmAPChkProcessMenu, go4it
   If BadAcct > 0 Then
    If go4it = False Then
     Call MainLog("APChk Post Errors - Procedure Stopped.")
     MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
     ReportFile$ = "TempLog.PRN"
     ViewPrint ReportFile$, "Error Log"
     frmCitiCancel.Show
     Unload frmPostAPChks
     Unload frmAPChkProcessMenu
     GoTo ExitPost
    Else
     Call MainLog("APChk Post Errors 2nd gothru GL Acct Error.")
     MsgBox "Errors Were Found. Review GL Posting Log.", vbOKOnly, "GL Account Error"
     GLLogFileName = "GLlog.dat"
     ReportFile$ = "GLlog.dat"
     ViewPrint ReportFile$, "Posting Log"
   End If
   End If
   PostCHKDAT "APCHKINF.DAT", frmAPChkProcessMenu, go4it
   If go4it = True Then
    KillFile APIFFileName$
    KillFile "APCHKINF.DAT"
    KillFile "TPAYLIST.LST"
    KillFile "TPAYLIST2.LST"
    KillFile "TPAYLISTD.LST"
    KillFile "TPAYNot.LST"
    KillFile "T2Pay.lst"
   End If
    Erase FundList$, FundAmts, APLedgerRec, APLedgerRec2, APDistRec
    Erase LedInfo, DistInfo, ChkRegInfo, CHKinfo
    Erase GLifRec
  If go4it = True Then
  Title$ = "A/P Check Register"
  ViewPrint PrintFileName$, Title$
  Title$ = "A/P Check Journal Entries"
  ViewPrint RptFileName$, Title$
  Call MainLog("APChk Post Complete.")
  End If
ExitPost:
Exit Sub

JEHeader:
  Print #RptFile, "A/P Check Journal Entries"
  Print #RptFile, "Posting Date: " + CheckDate$
  Print #RptFile,
  Mid$(JE$, 1) = "Account Number"
  Mid$(JE$, 20) = "Description"
  Mid$(JE$, 40) = "Reference"
  Mid$(JE$, 49) = "     Debit"
  Mid$(JE$, 63) = "    Credit"
  Print #RptFile, JE$
  LSet JE$ = ""
  Print #RptFile,
  JELineCnt = 5
Return


DoFundHeader:
  PageNum = PageNum + 1
  Page$ = Using("###", Str$(PageNum))

  Text1$ = "A/P Check Register Detail"
  Text2$ = "Total Checks Issued        "
  Print #PrintFile, Text1$
  Print #PrintFile, Dash$
  Print #PrintFile,
  Print #PrintFile, Text2$; Using("####", Str$(ChkCnt - UVoidCnt))
  Print #PrintFile, "           Totaling  "; Using("$###,###,###.##", Str$(TotalChkAmt#))
  Print #PrintFile,
  Print #PrintFile, " By Fund:"
  Linecnt = 7
Return

DoChkRegHeader:
  PageNum = PageNum + 1
  Page$ = Using("###", Str$(PageNum))
  Text1$ = "A/P Check Register Summary                                        Page:" + Page$
  Print #PrintFile, Text1$
  Print #PrintFile,
  Print #PrintFile, " Check No.   Check Date   Vendor Name                         Check Amount"
  Linecnt = 3
Return

InvHeader:
  Print #PrintFile, Dash2$
  Print #PrintFile, "       Inv Date    Inv Num                            Inv"
  Linecnt = Linecnt + 2
Return

DistHeader:
  Print #PrintFile, "                   Dist Acct                         Dist"
  Linecnt = Linecnt + 1
Return
ItsBroke:
  BadAcct = BadAcct + 1
  Print #ErrorP, "Error *** Call Software Support***"
  Print #ErrorP, "Ledger,Vendor,Interface Update"; Str$(cnt)
  Print #ErrorP, "Error Code"; Str(Err.Number)
  Resume Next
End Sub

Public Sub PostCHKDAT(ChkFileName$, formname As Form, go4it As Boolean)
  Dim OSChekRecLen As Integer, ChkRecLen As Integer, ChkNumRec As Integer
  Dim ChkinfoFile As Integer, Temp As Integer, OSChekFile As String
  Dim OSChekHandle As Integer, OSChekNumRec As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  Dim GLIFDate As String
  ReDim OSIFRec(1) As OSChekRecType
  OSChekRecLen = Len(OSIFRec(1))
  ChkinfoFile = FreeFile
  ReDim CheckInfo(1) As CheckInfoType3
  ChkRecLen = Len(CheckInfo(1))
  ChkNumRec = FileSize("APCHKINF.DAT") \ ChkRecLen
  ReDim CheckInfo(1 To ChkNumRec) As CheckInfoType3
  'FGetAH ChkFileName$, CheckInfo(1), ChkRecLen, ChkNumRec
  Open "APCHKINF.DAT" For Random Shared As ChkinfoFile Len = ChkRecLen
  For Temp = 1 To ChkNumRec
    Get ChkinfoFile, Temp, CheckInfo(Temp)
  Next

  OSChekFile$ = "CRCHEK.DAT"
  OSChekHandle = FreeFile
  Open OSChekFile$ For Random As #OSChekHandle Len = OSChekRecLen
  OSChekNumRec = (LOF(OSChekHandle) \ OSChekRecLen) + 1
  OpenVendorFile VendorFile, NumVRecs
  If go4it = True Then
    FrmShowPctComp.Label1 = "Posting Check Rec Trans."
  Else
    FrmShowPctComp.Label1 = "Verifying Check Rec Trans."
  End If
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents

  For cnt = 1 To ChkNumRec
    FrmShowPctComp.ShowPctComp cnt, ChkNumRec
    Get VendorFile, CheckInfo(cnt).VendorRecNum, Vendor
    'GLIFDate$ = Format(DateAdd("d", (CheckInfo(cnt).ChkDate), "12-31-1979"), "mmddyyyy")
    If CheckInfo(cnt).VoidFlag = False Then  'this is to keep cancel checks out of recnocil
      OSIFRec(1).chkdate = CheckInfo(cnt).chkdate
      OSIFRec(1).ChkNum = CheckInfo(cnt).LastChk
      OSIFRec(1).Desc = Vendor.VNAME
      OSIFRec(1).Amt = CheckInfo(cnt).ChkAmt
      OSIFRec(1).Src = 0
      '--Track Checks by Bank.  For Gate City, VA
      OSIFRec(1).Bankcode = CheckInfo(cnt).Bankcode
  
      'OSIFRec(1).TrType="D"
  
      If go4it = True Then
      Put OSChekHandle, OSChekNumRec, OSIFRec(1)
      End If
      OSChekNumRec = OSChekNumRec + 1
    End If
  Next
    
  Close

End Sub

