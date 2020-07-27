VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrnOpenPays 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Payables Report"
   ClientHeight    =   8892
   ClientLeft      =   36
   ClientTop       =   492
   ClientWidth     =   12192
   Icon            =   "frmPrnOpenPays.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboDistributions 
      Height          =   384
      Left            =   6252
      TabIndex        =   0
      Top             =   3528
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
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
      Columns         =   0
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
      ColDesigner     =   "frmPrnOpenPays.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   6240
      TabIndex        =   1
      Top             =   4152
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
      ColDesigner     =   "frmPrnOpenPays.frx":0D38
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
      Left            =   8256
      TabIndex        =   3
      Top             =   6408
      Width           =   1236
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
      Height          =   468
      Left            =   6420
      TabIndex        =   2
      Top             =   6408
      Width           =   1236
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "9:50 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11/18/2004"
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
      Left            =   3696
      TabIndex        =   7
      Top             =   4176
      Width           =   2388
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1764
      Left            =   2430
      Top             =   3120
      Width           =   7332
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   1416
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open Payables Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3684
      TabIndex        =   6
      Top             =   1656
      Width           =   4836
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Show Distributions:"
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
      Left            =   3816
      TabIndex        =   5
      Top             =   3576
      Width           =   2196
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   1296
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
Attribute VB_Name = "frmPrnOpenPays"
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
  Unload frmPrnOpenPays
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
  Me.HelpContextID = hlpOpenPay
  fpcboDistributions.AddItem "No"
  fpcboDistributions.AddItem "Yes"
  fpcboDistributions.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
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
        fpcboDistributions.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub cmdOk_Click()
  DeActivateControls frmPrnOpenPays, True
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt = 1 Then
    OpenPayable
  ElseIf rptopt = 2 Then
    OpenPayable2
  End If
  ActivateControls frmPrnOpenPays, True
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub OpenPayable()
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, cnt As Integer
  Dim ToPrintI As String, PrnFundTot As Integer, RptFundTot As String
  Dim NumFunds As Integer, APDistRecLen As Integer, VendorFile As Integer
  Dim PrintFile  As Integer, TPayCnt As Integer, NumVRecs As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Title As String
  Dim TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  Dim Header As String, User As String, ShowDist As Boolean, ToPrintD As String
  Dim NumVendRecs As Integer, VendorIdxFile As Integer, ToPrint As String
  Dim NumActiveVendors As Integer, VCnt As Integer, PRNFile As String
  Dim DoneVHeader As Integer, NumItems As Integer, NextTrans As Long
  Dim InvTotal As Double, DistAcctRec As Integer, AcctName As String
  Dim Vactive As Integer, Newrp As String, ToPrintV As String
  Header$ = "Open Payables Report"
  User$ = QPTrim$(GLUserName$)
  FrmShowPctComp.Label1 = "Creating Open Payables Report"
  FrmShowPctComp.Show , Me
  DoEvents
  If fpcboDistributions.ListIndex = 1 Then
    ShowDist = True
  Else
    ShowDist = False
  End If
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub
  ReDim FundAmts(1 To NumFunds) As Double
  ReDim APLedgerRec(1) As APLedger81RecType
  RecLen = Len(APLedgerRec(1))
  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))
  Dim Vendor As VendorRecType
  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  ReDim VIndex(1 To NumVendRecs) As VendorIdxRecType
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  For VCnt = 1 To NumVendRecs
    Get VendorIdxFile, VCnt, VendorIdx
    VIndex(VCnt).VendorCode = VendorIdx.VendorCode
    VIndex(VCnt).RecNum = VendorIdx.RecNum
  Next
  Close VendorIdxFile
  PrintFile = FreeFile
  Newrp = "Pay"
  GetRPTName Newrp
  PRNFile$ = Newrp
  Open PRNFile$ For Output As PrintFile
  PrnFundTot = FreeFile
  RptFundTot$ = "OPFund.prn"
  Open RptFundTot$ For Output As PrnFundTot
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen
  For VCnt = 1 To NumVendRecs
    FrmShowPctComp.ShowPctComp VCnt, NumVendRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnOpenPays, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    DoneVHeader = 0
    NumItems = 0
    Get VendorFile, VIndex(VCnt).RecNum, Vendor
    NextTrans& = Vendor.FrstTran
    Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedgerRec(1)
      If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
        NumItems = NumItems + 1
          ToPrintV$ = Str(NumItems) + "~" + Vendor.vnum + "~" + Vendor.VNAME + "~"
        GoSub PrintItem
      End If
      NextTrans& = APLedgerRec(1).NextTrans
    Loop
  Next
  GoSub FinishOpenReport
  Close
  Erase FundList$, FundAmts, APLedgerRec, APDistRec
  Load frmLoadingRpt
  'RptParm1$ = "From Vendor: " + Vend1St$ + "   Thru Vendor: " + VendLst$
  'RptParm2$ = "From Date: " + fpDate1.Text + "   Thru Date: " + fpDate2.Text
  'ARptVendHist.txtRptParm1.Caption = RptParm1$
  'ARptVendHist.txtRptParm2.Caption = RptParm2$
  
  If ShowDist = True Then
    ActivateControls frmPrnOpenPays, True
    ARptOpnPays1.txtTown.Caption = GLUserName$
    ARptOpnPays1.txtDate.Caption = Now
    ARptOpnPays1.RptVendTot.DataValue = TotalChkAmt#
    ARptOpnPays1.Label1.Caption = "OPEN PAYABLES REPORT"
    ARptOpnPays1.totvends.DataValue = ChkCnt
    ARptOpnPays1.GetName PRNFile$, RptFundTot$
    ARptOpnPays1.startrpt
  Else
    ActivateControls frmPrnOpenPays, True
    ARptOpnPays2.txtTown.Caption = GLUserName$
    ARptOpnPays2.txtDate.Caption = Now
    ARptOpnPays2.totvends.DataValue = ChkCnt
    'ARptOpnPays2.RptVendTot.DataValue = TotalChkAmt#
    ARptOpnPays2.Label1.Caption = "OPEN PAYABLES REPORT"
    ARptOpnPays2.GetName PRNFile$
    ARptOpnPays2.startrpt
  End If
  Exit Sub

PrintItem:

  ToPrintI$ = Space$(80)
  ToPrintI$ = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment) + "~"
  ToPrintI$ = ToPrintI$ + Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy") + "~"
  ToPrintI$ = ToPrintI$ + Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy") + "~"
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    ToPrintI$ = ToPrintI$ + Left$(APLedgerRec(1).PONum, 10) + "~"
  Else
    ToPrintI$ = ToPrintI$ + Left$(APLedgerRec(1).MPONum, 10) + "~"
  End If
  ToPrintI$ = ToPrintI$ + Using("##,###,###.##", Str$(APLedgerRec(1).Amt)) + "~"

  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  NextDist& = APLedgerRec(1).FrstDist
  If ShowDist Then


    ToPrintD$ = Space$(80)
    Do Until NextDist& = 0
      Get APDistFile, NextDist&, APDistRec(1)
      InvTotal# = InvTotal# + APDistRec(1).DistAmt
      DistAcctRec = AcctFind(APDistRec(1).DistAcctNum)
      AcctName$ = GetAcctTitle(DistAcctRec)

      ToPrintD$ = APDistRec(1).DistAcctNum + "~" + AcctName$ + "~"
      ToPrintD$ = ToPrintD$ + Using("##,###,###.##", Str$(APDistRec(1).DistAmt))
      ToPrintD$ = ToPrintD$ + "~" + Left$(APDistRec(1).DistAcctNum, GLFundLen)
      ToPrint$ = ToPrintV$ + ToPrintI$ + ToPrintD$
      Print #PrintFile, ToPrint$
      ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
      For FundCnt = 1 To NumFunds
        If ThisFund$ = FundList$(FundCnt) Then
          FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
          Exit For
        End If
      Next

      NextDist& = APDistRec(1).NextDist
    Loop

    InvTotal# = 0
  Else
    ToPrint$ = ToPrintV$ + ToPrintI$ + "~~~~"
    Print #PrintFile, ToPrint$
 
  End If
  Return
FinishOpenReport:
  'Print #PrintFile, "Report Totals:"
  'Print #PrintFile, "Vendors with Open Invoices: "; Using("#,###,###,###,###", Str$(ChkCnt))

  'Print #PrintFile, "                  Totaling: "; Using("$##,###,###.##", Str$(TotalChkAmt#))

  If ShowDist Then
    'Print #PrintFile, "Total Open By Fund:"
    For FundCnt = 1 To NumFunds
      If FundAmts(FundCnt) > 0 Then
        Print #PrnFundTot, FundList$(FundCnt) + "~" + Using("$##,###,###.##", Str$(FundAmts(FundCnt)))
      End If
    Next
  End If
  Return

FinishVendor:
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
  Vactive = 0

  Return
CancelExit:
  Exit Sub

End Sub
Private Sub OpenPayable2()
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, cnt As Integer, FF As String, MaxLines As Integer
  Dim Dash As String, DblDash As String, PageNum As Integer
  Dim NumFunds As Integer, APDistRecLen As Integer, VendorFile As Integer
  Dim PrintFile  As Integer, TPayCnt As Integer, NumVRecs As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VRecNum As Long
  Dim ChkCnt As Integer, Linecnt As Integer, Title As String
  Dim Page As String, TotalChkAmt As Double, VendTotal As Double
  Dim NextDist As Long, ThisFund As String, FundCnt As Integer
  Dim Header As String, User As String, ShowDist As Boolean
  Dim NumVendRecs As Integer, VendorIdxFile As Integer, ToPrint As String
  Dim NumActiveVendors As Integer, VCnt As Integer, PRNFile As String
  Dim DoneVHeader As Integer, NumItems As Integer, NextTrans As Long
  Dim InvTotal As Double, DistAcctRec As Integer, AcctName As String
  Dim Vactive As Integer, Newrp As String
  Dash$ = String$(80, "-")
  FF$ = Chr$(12)
  MaxLines = 50
  DblDash$ = String$(78, "=")
  PageNum = 0
  Header$ = "Open Payables Report"
  User$ = QPTrim$(GLUserName$)
  Page = 0
  FrmShowPctComp.Label1 = "Creating Open Payables Report"
  FrmShowPctComp.Show , Me
  DoEvents
 
  If fpcboDistributions.ListIndex = 1 Then
    ShowDist = True
  Else
    ShowDist = False
  End If
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub

  ReDim FundAmts(1 To NumFunds) As Double

  ReDim APLedgerRec(1) As APLedger81RecType
  RecLen = Len(APLedgerRec(1))

  ReDim APDistRec(1) As APDistRecType
  APDistRecLen = Len(APDistRec(1))

  Dim Vendor As VendorRecType

  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  ReDim VIndex(1 To NumVendRecs) As VendorIdxRecType
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  For VCnt = 1 To NumVendRecs
    Get VendorIdxFile, VCnt, VendorIdx
    VIndex(VCnt).VendorCode = VendorIdx.VendorCode
    VIndex(VCnt).RecNum = VendorIdx.RecNum
  Next
  Close VendorIdxFile

  PrintFile = FreeFile
  Newrp = "Pay"
  GetRPTName Newrp
  PRNFile$ = Newrp
  Open PRNFile$ For Output As PrintFile

  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen

  GoSub PrintOpenPayRptHeader
  'QPrintRC "Processing.. Please wait.", 25, 2, -1

  For VCnt = 1 To NumVendRecs
    FrmShowPctComp.ShowPctComp VCnt, NumVendRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnOpenPays, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    DoneVHeader = 0
    NumItems = 0
    Get VendorFile, VIndex(VCnt).RecNum, Vendor
    NextTrans& = Vendor.FrstTran
    Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedgerRec(1)
      If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then

        NumItems = NumItems + 1

        If MaxLines - Linecnt < 5 Then
          Print #PrintFile, FF$
          GoSub PrintOpenPayRptHeader
          GoSub PrintVendHeader
          'GOSUB InvHeader
        End If

        '--Print the Vendor Header First time thru
        If Not DoneVHeader Then
          If Linecnt > MaxLines Then
            Print #PrintFile, FF$
            Linecnt = 0
          End If
          GoSub PrintVendHeader
          'IF ShowDist = False THEN
          '  GOSUB InvHeader
          'END IF
        End If

        If Linecnt > MaxLines Then
          Print #PrintFile, FF$
          GoSub PrintOpenPayRptHeader
          GoSub PrintVendHeader
          'GOSUB InvHeader
        End If

        GoSub PrintItem
        'IF LineCnt > MaxLines THEN
        '  PRINT #PrintFile, FF$
        '  GOSUB PrintOpenPayRptHeader
        '  GOSUB PrintVendHeader
        '  'GOSUB InvHeader
        'END IF

      End If
      NextTrans& = APLedgerRec(1).NextTrans
    Loop

    If DoneVHeader Then
      GoSub FinishVendor
    End If
    If Linecnt > MaxLines Then
      Print #PrintFile, FF$
      GoSub PrintOpenPayRptHeader
    End If
  Next

  Print #PrintFile,             'JB
  'PRINT #PrintFile, DblDash$
  If Linecnt > MaxLines Then
    Print #PrintFile, FF$
  End If
  GoSub FinishOpenReport
  Print #PrintFile, FF$

  Close

  Erase FundList$, FundAmts, APLedgerRec, APDistRec


  Title$ = "Open Payables Report"
  ViewPrint PRNFile$, Title$
  KillFile PRNFile$
  Exit Sub

FinishOpenReport:
  Print #PrintFile, "Report Totals:"
  Print #PrintFile, "Vendors with Open Invoices: "; Using("#,###,###,###,###", Str$(ChkCnt))

  Print #PrintFile, "                  Totaling: "; Using("$##,###,###.##", Str$(TotalChkAmt#))

  If ShowDist Then
    Print #PrintFile, "Total Open By Fund:"
    For FundCnt = 1 To NumFunds
      If FundAmts(FundCnt) > 0 Then
        Print #PrintFile, FundList$(FundCnt); Tab(27); Using("$##,###,###.##", Str$(FundAmts(FundCnt)))

      End If
    Next
  End If
  Return

PrintOpenPayRptHeader:
  Page = Page + 1
  Print #PrintFile, Tab(40 - (Int(Len(User$) / 2))); User$
  Print #PrintFile, Tab(40 - (Int(Len(Header$) / 2))); Header$
  Print #PrintFile,
  Print #PrintFile, "Report Date: "; Date$; Tab(67); "Page #"; Page
  Print #PrintFile,
  Print #PrintFile, "Inv Num/Desc                      Inv Date    Due Date    PO              Amount"

  Print #PrintFile, String$(80, "=")
  Linecnt = 6

  Return


InvHeader:

  'PRINT #PrintFile,
  'PRINT #PrintFile, "Inv Num                   Inv Date    Due Date    PO
  'PRINT #PrintFile, STRING$(78, "-")
  'LineCnt = LineCnt + 2
  '"----------  ----------  -------------------------  ----------    ---------
  Return

PrintVendHeader:
  'PRINT #PrintFile,
  'PRINT #PrintFile, DblDash$
  Print #PrintFile, Vendor.vnum; Tab(15); Vendor.VNAME
  DoneVHeader = -1
  Linecnt = Linecnt + 1

  Return

PrintItem:
  'IF ShowDist THEN
  '  GOSUB InvHeader
  'END IF

  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 1) = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment)
  Mid$(ToPrint$, 35) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 46) = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    Mid$(ToPrint$, 57) = Left$(APLedgerRec(1).PONum, 10)
  Else
    Mid$(ToPrint$, 57) = Left$(APLedgerRec(1).MPONum, 10)
  End If
  Mid$(ToPrint$, 68) = Using("##,###,###.##", Str$(APLedgerRec(1).Amt))
  Print #PrintFile, ToPrint$
  Linecnt = Linecnt + 1

  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  NextDist& = APLedgerRec(1).FrstDist
  If ShowDist Then

    Print #PrintFile,

    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 5) = "Accounting Distribution"
    Print #PrintFile, ToPrint$
    Linecnt = Linecnt + 1

    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 10) = "Acct Number"
    Mid$(ToPrint$, 28) = "Description"
    Mid$(ToPrint$, 56) = "      Amount"
    Print #PrintFile, ToPrint$
    Linecnt = Linecnt + 1

    ToPrint$ = Space$(80)
    Do Until NextDist& = 0
      Get APDistFile, NextDist&, APDistRec(1)
      InvTotal# = InvTotal# + APDistRec(1).DistAmt
      DistAcctRec = AcctFind(APDistRec(1).DistAcctNum)
      AcctName$ = GetAcctTitle(DistAcctRec)

      If Linecnt > MaxLines Then
        Print #PrintFile, FF$
        GoSub PrintOpenPayRptHeader
      End If

      Mid$(ToPrint$, 10) = APDistRec(1).DistAcctNum
      Mid$(ToPrint$, 28) = AcctName$
      Mid$(ToPrint$, 56) = Using("##,###,###.##", Str$(APDistRec(1).DistAmt))
      Print #PrintFile, ToPrint$
      Linecnt = Linecnt + 1

      ThisFund$ = Left$(APDistRec(1).DistAcctNum, GLFundLen)
      For FundCnt = 1 To NumFunds
        If ThisFund$ = FundList$(FundCnt) Then
          FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
          Exit For
        End If
      Next
      NextDist& = APDistRec(1).NextDist
    Loop

    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 56) = String$(13, "-")
    Print #PrintFile, ToPrint$

    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 28) = "Total Distributed"
    Mid$(ToPrint$, 56) = Using("##,###,###.##", Str$(InvTotal#))
    Print #PrintFile, ToPrint$

    Linecnt = Linecnt + 2
    InvTotal# = 0

  End If

  If ShowDist Then
    Print #PrintFile,
    Print #PrintFile, String$(78, "-")
    Linecnt = Linecnt + 2
  End If

  Return


FinishVendor:
  If ShowDist Then
    If NumItems > 1 Then
      Print #PrintFile, QPTrim$(Vendor.VNAME); " Total: "; Tab(66); Using("##,###,###.##", Str(VendTotal#))
      Print #PrintFile, String$(78, "-")
      Linecnt = Linecnt + 2
      'ELSE
      '  PRINT #PrintFile, DblDash$
      '  LineCnt = LineCnt + 1
    End If
  Else
    If NumItems > 1 Then
      Print #PrintFile, Tab(66); "------------"
      Print #PrintFile, QPTrim$(Vendor.VNAME); " Total: "; Tab(66); Using("##,###,###.##", Str(VendTotal#))
      Print #PrintFile, DblDash$
      Linecnt = Linecnt + 3     '2
    Else
      Print #PrintFile, DblDash$
      Linecnt = Linecnt + 1
    End If
  End If
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
  Vactive = 0

  Return
CancelExit:
  Exit Sub

End Sub


Private Sub fpcboDistributions_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeySpace Then
    fpcboDistributions.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDistributions.ListIndex = -1
    fpcboDistributions.Action = ActionClearSearchBuffer
  End If
  If fpcboDistributions.ListDown <> True Then
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
