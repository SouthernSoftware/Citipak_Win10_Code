VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnPONonApproved 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Non-Approved"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnPONonApproved.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   6255
      TabIndex        =   1
      Top             =   4755
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
      ColDesigner     =   "frmPrnPONonApproved.frx":08CA
   End
   Begin LpLib.fpCombo fpcboDepartment 
      Height          =   405
      Left            =   6240
      TabIndex        =   0
      Top             =   3990
      Width           =   2145
      _Version        =   196608
      _ExtentX        =   3784
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnPONonApproved.frx":0C30
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   468
      Left            =   4452
      TabIndex        =   7
      Top             =   5616
      Width           =   1380
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
      Height          =   468
      Left            =   6360
      TabIndex        =   6
      Top             =   5616
      Width           =   1380
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8484
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
            TextSave        =   "4:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "8/29/2008"
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
      Left            =   3660
      TabIndex        =   5
      Top             =   4800
      Width           =   2388
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3204
      Left            =   3336
      Top             =   3336
      Width           =   5532
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   1344
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Orders Non-Approved"
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
      Left            =   3678
      TabIndex        =   4
      Top             =   1584
      Width           =   4836
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
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
      Height          =   324
      Left            =   4596
      TabIndex        =   2
      Top             =   4032
      Width           =   1356
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   1224
      Width           =   7020
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
Attribute VB_Name = "frmPrnPONonApproved"
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
Dim GLTrans   As GLTransRecType
Dim PO As POFORMRecType2
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer

Private Sub cmdExit_Click()
  frmPOProcessMenu.Show
  Unload frmPrnPONonApproved
End Sub

Private Sub cmdOk_Click()
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    PrnNAEditList
  ElseIf rptopt = 2 Then
    PrnNAEditList2
  End If
End Sub

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
        fpcboDepartment.SetFocus
        KeyCode = 0
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdOk_Click
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
  Me.HelpContextID = hlpNApprPO
  DeptList fpcboDepartment
  fpcboDepartment.ListIndex = 0
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

Private Sub PrnNAEditList()
  Dim cnt As Integer, DebitCol As Integer, ReportFile2 As String
  Dim CreditCol As Integer, PrnFileNum2 As Integer
  Dim PRNFile As Integer, HowMany As Integer, CommaFmt As String
  Dim ReportFile As String, ToPrint As String, OverFmt As String
  Dim Header As String, DistSumLine As String, ToPrint2 As String
  Dim TransTotal As Double, TranCnt As Integer, FileName As String
  Dim RegTitle As String, TranCol As Integer, CashCol As Integer
  Dim DeptNumber As String, NYBeg As Integer, NYEnd As Integer
  Dim ThisDist As Double, Accttotal As Double, Over As String
  Dim ThisAcct As String, AcctCnt As Integer, WhatAcct As Integer
  Dim Found As Boolean, Fund As Integer, FundNum As String
  Dim GrdTot As Double, HCnt As Integer, Newrp As String
  Dim POEditFile As Integer, NumEdTrans As Integer, PODept As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, AcctDist As Integer
  Dim IdxFile As Integer, NumIdxRecs As Integer, Transaction As Integer
  Dim AcctFile As Integer, NumAccts As Integer, DidCnt As Integer
  Dim TransFileNum As Integer, NumTrans As Long, PrnFileNum As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, TotTranDist As Double
  Dim Itmdesc As String
  DebitCol = 42
  CreditCol = 58
  CommaFmt$ = "###,###,###.##"   'ten millions
  OverFmt$ = "###,###,###.##"       'ten millions
  DistSumLine$ = "--------------"
  ReDim Title$(5)
  TransTotal# = 0
  TranCnt = 0
  Itmdesc$ = ""
  ToPrint$ = ""
  ToPrint2$ = ""
  Newrp = "POnon"
  GetRPTName Newrp
  FileName$ = Newrp
  RegTitle$ = "Purchase Orders"
  TranCol = CreditCol
  CashCol = DebitCol
  FrmShowPctComp.Label1 = "Creating PO Non-Approved Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  mnuOptions.Enabled = False
  fpcboDepartment.col = 1
  DeptNumber$ = QPTrim$(fpcboDepartment.ColText)

  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile


  NYBeg = GLSetUpRec(1).NYBeg
  NYEnd = GLSetUpRec(1).NYEnd
  Erase GLSetUpRec
  updateaccttots
  OpenAcctIdx IdxFile, NumIdxRecs
  OpenAcctFile AcctFile, NumAccts

  ReDim POChk(1 To NumIdxRecs) As AcctPOChkType

  For cnt = 1 To NumIdxRecs
    Get IdxFile, cnt, AcctIdx
    If AcctIdx.RecNum > 0 Then
      DidCnt = DidCnt + 1
      Get AcctFile, AcctIdx.RecNum, Acct
      'IF INSTR(Acct.Num, "10-5600-1600") > 0 THEN STOP
      If Acct.Deleted = 0 Then
        POChk(DidCnt).Acct = Acct.Num
        POChk(DidCnt).Bgt = Acct.Bgt
        POChk(DidCnt).Encumb = Acct.Encumb
        'POChk(DidCnt).Bal = Acct.Bal
        POChk(DidCnt).NYApp = Acct.NYApp
        POChk(DidCnt).Bal = Acct.YTD
      End If
    End If
  Next
  Close AcctFile, IdxFile


  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileNum = FreeFile
  Open FileName$ For Output As #PrnFileNum
  ReportFile2$ = "POFundSum.prn"
  PrnFileNum2 = FreeFile
  Open ReportFile2$ For Output As #PrnFileNum2
  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  ReDim FundGrdTot#(1 To NumFunds)

  'GoSub PrintHeader

  For Transaction = 1 To NumEdTrans
    Get POEditFile, Transaction, PO
    FrmShowPctComp.ShowPctComp Transaction, NumEdTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    PODept$ = QPTrim$(PO.REQNUM)
    If PO.Deleted <> True Then
      If PODept$ = DeptNumber$ Or DeptNumber$ = "All" Then
        If Left$(PO.PONum, 3) = "N/A" Then
          TranCnt = TranCnt + 1
          TransTotal# = Round#(TransTotal# + PO.POAmt)
          '--Print 1st Line - Transaction details
          ToPrint$ = ""    'SPACE$(78)

          ToPrint$ = QPTrim(PO.VNDRCODE) + " " + Str$(PO.VNDRREC)
          ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy")
          ToPrint$ = ToPrint$ + "~" + PO.PONum + Str(TranCnt)
          ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(PO.POAmt))
          'Print #PrnFileNum, ToPrint$

          '--Blank line between detail and acct'g distributions
          ToPrint$ = ToPrint$ + "~" + "Dept # " + PO.REQNUM
          '--Print Distribution Label
'          LSet ToPrint$ = ""
'          Mid$(ToPrint$, 2) = "Accounting Distribution:"
'          Print #PrnFileNum, ToPrint$
'          '--Print Field Titles
'          LSet ToPrint$ = ""
'          Mid$(ToPrint$, 4) = "Account Number                             Distribution"
'          Print #PrnFileNum, ToPrint$
          '--Print Accounting Distributions
          TotTranDist# = 0
          '--Loop Thru distributions to print and summarize
          For AcctDist = 1 To 36
            ThisAcct$ = QPTrim$(PO.ITEMS(AcctDist).ACCTNO)
            If Len(ThisAcct$) Then
            If AcctFind(PO.ITEMS(AcctDist).ACCTNO) > 0 Then
              For AcctCnt = 1 To DidCnt
                If InStr(POChk(AcctCnt).Acct, ThisAcct$) > 0 Then
                  WhatAcct = AcctCnt
                  ThisDist# = PO.ITEMS(AcctDist).EXT
                  Itmdesc$ = QPTrim$(PO.ITEMS(AcctDist).Desc)
                  POChk(WhatAcct).POTotal = Round#(POChk(WhatAcct).POTotal + ThisDist#)
                  Exit For
                End If
              Next
              '--Print this distribution
              ToPrint2$ = ""
              ToPrint2$ = ThisAcct$

              ToPrint2$ = ToPrint2$ + "~" + Using$(CommaFmt$, Str$(ThisDist#))
              Accttotal# = Round#(POChk(WhatAcct).Bal + POChk(WhatAcct).POTotal + POChk(WhatAcct).Encumb)
              If PO.PODATE < NYBeg Then
                If Accttotal# > POChk(WhatAcct).Bgt Then
                  Over$ = "OVER BDGT:" + Using$(OverFmt$, Str$(Abs(POChk(WhatAcct).Bgt - Accttotal#)))
                  ToPrint2$ = ToPrint2$ + "~" + Over$ + "~" + Itmdesc$
                Else
                  Over$ = "  "
                End If
              Else
                If Accttotal# > POChk(WhatAcct).Bgt Then
                  Over$ = "OVER BDGT:" + Using$(OverFmt$, Str$(Abs(POChk(WhatAcct).NYApp - Accttotal#)))
                  ToPrint2$ = ToPrint2$ + "~" + Over$ + "~" + Itmdesc$
                Else
                  Over$ = "  "
                End If
              End If
              ToPrint2$ = ToPrint2$ + "~" + Over$ + "~" + Itmdesc$
              Print #PrnFileNum, ToPrint$ + "~" + ToPrint2$
              Itmdesc$ = ""
              TotTranDist# = TotTranDist# + ThisDist#


              'Sum by fund
              Found = False
              For Fund = 1 To NumFunds
                FundNum$ = Left$(PO.ITEMS(AcctDist).ACCTNO, GLFundLen)
                If FundNum$ = FundList$(Fund) Then
                  Found = True
                  FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + ThisDist#)
                  Exit For
                End If
              Next

              If Not Found Then 'Oh.Shit = True
                 Unload FrmShowPctComp
                 MsgBox "Error, Invalid Fund Found.", vbOKOnly, "Error"
                 ActivateControls frmPrnPONonApproved
                 Exit Sub
              End If
            End If
            End If              'Active transaction test

          Next  'Distribution


        End If  'Not deleted test
      End If    'Not Correct Dept
    End If      'Not Approved Flag
  Next          'Transaction

  '--Summary

  For cnt = 1 To NumFunds
    If FundGrdTot#(cnt) > 0 Then
      ToPrint$ = ""
      ToPrint$ = FundList$(cnt)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(FundGrdTot#(cnt)))
      Print #PrnFileNum2, ToPrint$
      GrdTot# = Round#(GrdTot# + FundGrdTot#(cnt))
    End If
  Next
  Close
  If NumEdTrans < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
  End If

  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  Load frmLoadingRpt
  ARptPOEdits.totTrans = Using$("####", Str$(TranCnt))
  ARptPOEdits.totGrand = Using$(CommaFmt$, Str$(TransTotal#))
  ARptPOEdits.totFunds = Using$(CommaFmt$, Str$(GrdTot#))
  
  If NumEdTrans < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    ARptPOEdits.Label13.Visible = True
    ARptPOEdits.Label13.Caption = "** No Unapproved Purchase Orders **"
  End If
  ARptPOEdits.txtDate = Now
  ARptPOEdits.txtTown = GLUserName$
  ARptPOEdits.Label1.Caption = "Purchase Orders : NOT YET APPROVED!"
  ARptPOEdits.Label2.Caption = "Dept #: " + DeptNumber$
  ARptPOEdits.Caption = "Purchase Orders : Not Yet Approved"
  ARptPOEdits.GetName FileName$, ReportFile2$
  ARptPOEdits.startrpt
  
Exit Sub
CancelExit:
  Exit Sub
End Sub
Private Sub PrnNAEditList2()
  Dim MaxLines As Integer, cnt As Integer, DebitCol As Integer
  Dim Linecnt As Integer, Page As Integer, CreditCol As Integer
  Dim PRNFile As Integer, HowMany As Integer, CommaFmt As String
  Dim ReportFile As String, ToPrint As String, OverFmt As String
  Dim FF As String, Header As String, DistSumLine As String
  Dim TransTotal As Double, TranCnt As Integer, FileName As String
  Dim RegTitle As String, TranCol As Integer, CashCol As Integer
  Dim DeptNumber As String, NYBeg As Integer, NYEnd As Integer
  Dim ThisDist As Double, Accttotal As Double, Over As String
  Dim ThisAcct As String, AcctCnt As Integer, WhatAcct As Integer
  Dim Found As Boolean, Fund As Integer, FundNum As String
  Dim GrdTot As Double, HCnt As Integer, Newrp As String
  Dim POEditFile As Integer, NumEdTrans As Integer, PODept As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, AcctDist As Integer
  Dim IdxFile As Integer, NumIdxRecs As Integer, Transaction As Integer
  Dim AcctFile As Integer, NumAccts As Integer, DidCnt As Integer
  Dim TransFileNum As Integer, NumTrans As Long, PrnFileNum As Integer
  Dim SetUpRecLen As Integer, SetupFile As Integer, TotTranDist As Double
  Dim Itmdesc As String
  DebitCol = 30
  CreditCol = 40
  CommaFmt$ = "###,###,###.##"   'ten millions
  OverFmt$ = "###,###,###.##"       'ten millions
  DistSumLine$ = "--------------"
  FF$ = Chr$(12)
  ReDim Title$(5)
  MaxLines = 56
  TransTotal# = 0
  TranCnt = 0
  ToPrint$ = Space$(80)
  Itmdesc$ = ""
  Page = 0
  Newrp = "POnon"
  GetRPTName Newrp
  FileName$ = Newrp
  RegTitle$ = "Purchase Orders"
  TranCol = CreditCol
  CashCol = DebitCol
  FrmShowPctComp.Label1 = "Creating PO Non-Approved Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOk.Enabled = False
  mnuOptions.Enabled = False
  fpcboDepartment.col = 1
  DeptNumber$ = QPTrim$(fpcboDepartment.ColText)

  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile


  NYBeg = GLSetUpRec(1).NYBeg
  NYEnd = GLSetUpRec(1).NYEnd
  Erase GLSetUpRec
  updateaccttots
  OpenAcctIdx IdxFile, NumIdxRecs
  OpenAcctFile AcctFile, NumAccts

  ReDim POChk(1 To NumIdxRecs) As AcctPOChkType

  For cnt = 1 To NumIdxRecs
    Get IdxFile, cnt, AcctIdx
    If AcctIdx.RecNum > 0 Then
      DidCnt = DidCnt + 1
      Get AcctFile, AcctIdx.RecNum, Acct
      'IF INSTR(Acct.Num, "10-5600-1600") > 0 THEN STOP
      If Acct.Deleted = 0 Then
        POChk(DidCnt).Acct = Acct.Num
        POChk(DidCnt).Bgt = Acct.Bgt
        POChk(DidCnt).Encumb = Acct.Encumb
        'POChk(DidCnt).Bal = Acct.Bal
        POChk(DidCnt).NYApp = Acct.NYApp
        POChk(DidCnt).Bal = Acct.YTD
      End If
    End If
  Next
  Close AcctFile, IdxFile


  OpenPOEditFile POEditFile, NumEdTrans
  PrnFileNum = FreeFile

  Open FileName$ For Output As #PrnFileNum

  '--Get a list of active funds
  ReDim FundList(1) As String
  GetFundList FundList(), NumFunds
  'REDIM FundSum#(1 TO NumFunds)
  ReDim FundGrdTot#(1 To NumFunds)

  GoSub PrintHeader

  For Transaction = 1 To NumEdTrans
    Get POEditFile, Transaction, PO
    FrmShowPctComp.ShowPctComp Transaction, NumEdTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdOk.Enabled = True
      mnuOptions.Enabled = True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    PODept$ = QPTrim$(PO.REQNUM)
    If PO.Deleted <> True Then
      If PODept$ = DeptNumber$ Or DeptNumber$ = "All" Then
        If Left$(PO.PONum, 3) = "N/A" Then
          TranCnt = TranCnt + 1
          TransTotal# = Round#(TransTotal# + PO.POAmt)
          '--Print 1st Line - Transaction details
          LSet ToPrint$ = ""    'SPACE$(78)

          LSet ToPrint$ = PO.VNDRCODE + " " + Str$(PO.VNDRREC)
          Mid$(ToPrint$, 30) = Format(DateAdd("d", (PO.PODATE), "12-31-1979"), "mm/dd/yyyy")
          Mid$(ToPrint$, 42) = PO.PONum
          Mid$(ToPrint$, 66) = Using$(CommaFmt$, Str$(PO.POAmt))
          Print #PrnFileNum, ToPrint$

          '--Blank line between detail and acct'g distributions
          Print #PrnFileNum, "Dept # "; PO.REQNUM
          '--Print Distribution Label
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 2) = "Accounting Distribution:"
          Print #PrnFileNum, ToPrint$
          '--Print Field Titles
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 2) = "Item Desc               Account Number    Distribution"
          Print #PrnFileNum, ToPrint$
          '--Print Accounting Distributions
          TotTranDist# = 0
          '--Loop Thru distributions to print and summarize
          For AcctDist = 1 To 36
            ThisAcct$ = QPTrim$(PO.ITEMS(AcctDist).ACCTNO)
            If Len(ThisAcct$) Then
            If AcctFind(PO.ITEMS(AcctDist).ACCTNO) > 0 Then
              For AcctCnt = 1 To DidCnt
                If InStr(POChk(AcctCnt).Acct, ThisAcct$) > 0 Then
                  WhatAcct = AcctCnt
                  ThisDist# = PO.ITEMS(AcctDist).EXT
                  Itmdesc$ = Mid$(QPTrim$(PO.ITEMS(AcctDist).Desc), 1, 24)
                  POChk(WhatAcct).POTotal = Round#(POChk(WhatAcct).POTotal + ThisDist#)
                  Exit For
                End If
              Next
              '--Print this distribution
              LSet ToPrint$ = ""
              Mid$(ToPrint$, 2) = Itmdesc$
              Mid$(ToPrint$, 28) = ThisAcct$

              Mid$(ToPrint$, TranCol) = Using$(CommaFmt$, Str$(ThisDist#))
              Accttotal# = Round#(POChk(WhatAcct).Bal + POChk(WhatAcct).POTotal + POChk(WhatAcct).Encumb)
              If PO.PODATE < NYBeg Then
                If Accttotal# > POChk(WhatAcct).Bgt Then
                  Over$ = "OVER BDGT:" + Using$(OverFmt$, Str$(Abs(POChk(WhatAcct).Bgt - Accttotal#)))
                  Mid$(ToPrint$, 56) = Over$
                End If
              Else
                If Accttotal# > POChk(WhatAcct).Bgt Then
                  Over$ = "OVER BDGT:" + Using$(OverFmt$, Str$(Abs(POChk(WhatAcct).NYApp - Accttotal#)))
                  Mid$(ToPrint$, 56) = Over$
                End If
              End If
              Print #PrnFileNum, ToPrint$

              TotTranDist# = TotTranDist# + ThisDist#

              Linecnt = Linecnt + 1
              If Linecnt >= MaxLines Then
                Print #PrnFileNum, FF$
                GoSub PrintHeader
              End If

              'Sum by fund
              Found = False
              For Fund = 1 To NumFunds
                FundNum$ = Left$(PO.ITEMS(AcctDist).ACCTNO, GLFundLen)
                If FundNum$ = FundList$(Fund) Then
                  Found = True
                  FundGrdTot#(Fund) = Round#(FundGrdTot#(Fund) + ThisDist#)
                  Exit For
                End If
              Next

              If Not Found Then 'Oh.Shit = True
                 Unload FrmShowPctComp
                 MsgBox "Error, Invalid Fund Found.", vbOKOnly, "Error"
                 ActivateControls frmPrnPONonApproved
                 Exit Sub
              End If
            End If
            End If              'Active transaction test

          Next  'Distribution

          '--Summary line after last distribution
          LSet ToPrint$ = ""
          Mid$(ToPrint$, TranCol) = DistSumLine$
          Print #PrnFileNum, ToPrint$

          '--Transaction Distribution Totals
          LSet ToPrint$ = ""
          Mid$(ToPrint$, 4) = "Total Distributed"
          Mid$(ToPrint$, TranCol) = Using$(CommaFmt$, Str$(TotTranDist#))
          Print #PrnFileNum, ToPrint$
          Linecnt = Linecnt + 2
          If Linecnt >= MaxLines Then
            Print #PrnFileNum, FF$
            GoSub PrintHeader
          End If

          '--2 blank lines before next distribution
          ToPrint$ = String$(80, "=")
          Print #PrnFileNum, ToPrint$
          Print #PrnFileNum,
          Linecnt = Linecnt + 2
          If Linecnt >= MaxLines Then
            Print #PrnFileNum, FF$
            GoSub PrintHeader
          End If
        End If  'Not deleted test
      End If    'Not Correct Dept
    End If      'Not Approved Flag
  Next          'Transaction
  'PRINT #PrnFileNum, STRING$(80, "-")
  If Linecnt > 45 Then
    Print #PrnFileNum, FF$
  End If

  '--Summary
  LSet ToPrint$ = ""
  LSet ToPrint$ = "File Totals:"
  Print #PrnFileNum, ToPrint$

  LSet ToPrint$ = ""
  LSet ToPrint$ = "Number of Transactions"
  Mid$(ToPrint$, 31) = Using$("####", Str$(TranCnt))
  Print #PrnFileNum, ToPrint$

  LSet ToPrint$ = ""
  LSet ToPrint$ = "Grand Totals"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(TransTotal#))
  Print #PrnFileNum, ToPrint$

  Print #PrnFileNum,
  LSet ToPrint$ = ""
  LSet ToPrint$ = "Summary by Fund:"
  Print #PrnFileNum, ToPrint$

  For cnt = 1 To NumFunds
    If FundGrdTot#(cnt) > 0 Then
      LSet ToPrint$ = ""
      LSet ToPrint$ = "Fund" + " " + FundList$(cnt)
      Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(FundGrdTot#(cnt)))
      Print #PrnFileNum, ToPrint$
      GrdTot# = Round#(GrdTot# + FundGrdTot#(cnt))
    End If
  Next

  LSet ToPrint$ = ""
  LSet ToPrint$ = "Total All Funds"
  Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(GrdTot#))
  Print #PrnFileNum, ToPrint$
  Print #PrnFileNum, FF$

  
  If NumEdTrans < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
    Print #PrnFileNum, "** No Unapproved Purchase Orders **"
  End If
  Close
  ViewPrint FileName$, RegTitle$
  KillFile FileName$
  Me.cmdExit.Enabled = True
  Me.cmdOk.Enabled = True
  EnableCloseButton Me.hwnd, True
  mnuOptions.Enabled = True
  Exit Sub

PrintHeader:
  Page = Page + 1
  Title$(1) = "Purchase Orders : NOT YET APPROVED !!!"
  Title$(2) = "Run Date: " + Date$ + "                                                  Page : " & Page
  Title$(3) = "Dept #: " + DeptNumber$
  Title$(4) = "Vendor                       Date        PO No"
  Title$(5) = String$(80, "-")

  Print #PrnFileNum, GLUserName$
  For HCnt = 1 To 5
    Print #PrnFileNum, Title$(HCnt)
  Next
  Linecnt = 6
  Return
CancelExit:
  Exit Sub
End Sub

Private Sub fpcboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDepartment.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDepartment.ListIndex = -1
    fpcboDepartment.Action = ActionClearSearchBuffer
  End If
  If fpcboDepartment.ListDown <> True Then
    If KeyCode = vbKeyDown Then
        SendKeys "{Tab}"
        KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
