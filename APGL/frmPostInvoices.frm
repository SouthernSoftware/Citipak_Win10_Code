VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostInvoices 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Invoices"
   ClientHeight    =   8835
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12195
   Icon            =   "frmPostInvoices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   375
      Left            =   2568
      Top             =   3024
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
      Left            =   9792
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   1356
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
      Left            =   8112
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1356
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8580
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   450
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
            TextSave        =   "1:26 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/28/2008"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   348
      Left            =   5172
      TabIndex        =   7
      Top             =   2952
      Width           =   1692
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1344
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Post Invoices"
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
      Left            =   4542
      TabIndex        =   6
      Top             =   1584
      Width           =   3108
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Before You Post, Make Sure You Have Printed An Invoice Register. "
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
      Top             =   3384
      Width           =   7308
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
      TabIndex        =   4
      Top             =   4416
      Width           =   6780
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Once Posted, The Invoice Register May NOT Be Printed Again. "
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
      TabIndex        =   3
      Top             =   3768
      Width           =   5196
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   1224
      Width           =   5772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2172
      Left            =   2400
      Top             =   2784
      Width           =   7452
   End
End
Attribute VB_Name = "frmPostInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim Acct As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim APIED As APInv85Type
Dim ApLedger As APLedger81RecType
Dim APDist As APDistRecType
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim APAcct As String, EncAcct As String
Dim PTRecs() As Long
Private Sub cmdExit_Click()
  Unload frmPostInvoices
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
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdOK.Enabled = False
  PostInvoices
  
  Me.cmdExit.Enabled = True
  Me.cmdOK.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmPostInvoices
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
  DoEvents
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetAPAcct APAcct
  GetEncAcct EncAcct
  StatusBar1.Panels.Item(1).Text = GLUserName
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.HelpContextID = hlpPostInv
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PostInvoices()
  Dim cnt As Integer, Active As Integer, CntD As Integer, PostCnt As Integer
  Dim GotPOs As Boolean, NumFunds As Integer, AP2Post As Integer
  Dim LedgerRecLen As Integer, DistRecLEn As Integer, VendorRecLen As Integer
  Dim APEditFile As Integer, NumEdTrans As Integer, MSrc As String
  Dim PO2Post As Integer, POIFRecNum As Integer, RecordNum As Integer
  Dim APLedgerFile As Integer, NumLdgTran As Long, LogFile As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VendRecNum As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, PrevVendTr As Long
  Dim AcctFileNum As Integer, NumAccts As Integer, FrstVendTR As Long
  Dim DistCnt As Integer, Nextone As Long, Used As Integer, AcctDist As Integer
  Dim Dist As Integer, BadVendor As Integer, Icnt As Integer, Fund As Integer
  Dim FundNum As String, BadPOTrans As Integer, BadGLTrans As Integer
  Dim ToPrint As String, ReportFile As String, GLLogFileName As String
  Dim POLogFileName As String, TempCr As Double, TempDr As Double
  Dim TempAcct As Integer, Editing As Boolean, PTcnt As Integer, RptFile As Integer
  ReDim APDistRec(1) As APDistRecType
  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim Tr2Post(1) As GLTransRecType
  ReDim Preserve PTRecs(1 To 1) As Long
  'On Error GoTo ERRORSTUFF
  LedgerRecLen = Len(APLedgerRec(1))
  DistRecLEn = Len(APDistRec(1))
  VendorRecLen = Len(Vendor)
  '--Verify that there are transactions
  OpenAPEditFile APEditFile, NumEdTrans
  For cnt = 1 To NumEdTrans
    Get APEditFile, cnt, APIED
    If Not APIED.DelFlag Then
      Active = Active + 1
    End If
    If APIED.LOCKED = True Then
      Editing = True
      Exit For
    End If
  Next
  Close
  If Not Editing Then
    If Active = 0 Then            '--No active transactions - tell user and get
      MsgBox "No Transactions To Post.", vbOKOnly, "Post Canceled"
      GoSub OutNow
    End If
    'SetAttr ("APIED.DAT"), vbReadOnly
  Else
    MsgBox "Invoice Edit Is In Process, Please Close Edit Procedure Before Trying To Post.", vbOKOnly, "Canceled"
    GoSub OutNow
  End If
  If MsgBox("Are You Sure You Are Ready To Post?", vbYesNo, "Continue Posting?") = vbNo Then
    GoSub OutNow
  End If
  FrmShowPctComp.Label1 = "Verifying Invoice Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  Call MainLog("Start APInv Post")
  SH_CopyFile "APIED.DAT", "APIEDSV.DAT"
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  ReDim TrFundSum#(1 To NumFunds)

  AP2Post = FreeFile
  Open "APINVIF.dat" For Random As AP2Post Len = Len(Tr2Post(1))
  PO2Post = FreeFile
  'If Exist("Poinvif.dat") Then
    Open "POINVIF.dat" For Random As PO2Post Len = Len(Tr2Post(1))
  'End If
  POIFRecNum = 0
  RecordNum = 0 'Reset Active counter for posting
  MSrc$ = "AP" + Format$(Now, "mmddyy")
  OpenAPEditFile APEditFile, NumEdTrans
  OpenAPLedgerFile APLedgerFile, NumLdgTran&, LedgerRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  OpenVendorFile VendorFile, NumVRecs
  OpenAcctFile AcctFileNum, NumAccts
'
  LogFile = FreeFile
  Open "GLUTIL.LOG" For Append As LogFile
  Print #LogFile, "Invoice Posting started @ " + Time$ + " on " + Date$
'create temp trans files test post first go thru if ok then realy post
    For Icnt = 1 To NumEdTrans     '--Create interface files for gltrans & potran
    FrmShowPctComp.ShowPctComp Icnt, NumEdTrans
'Call MainLog("OK 4")
    ReDim TrFundSum#(1 To NumFunds)
    '--Reinitialize transaction fund summary elements for next loop
    Get APEditFile, Icnt, APIED
    If Not APIED.DelFlag Then  '--If invoice not deleted post it
      GotPOs = False
        If APIED.POFLAG <> 0 Then
          GotPOs = True
        End If
     
      '****************************
      For AcctDist = 1 To 36   '--Write only those with account numbers
        If Len(QPTrim$(APIED.Dist(AcctDist).DACN)) > 0 Then
          RecordNum = RecordNum + 1
          Tr2Post(1).AcctNum = APIED.Dist(AcctDist).DACN
          Tr2Post(1).TRDATE = APIED.DISTDATE
          Tr2Post(1).Desc = APIED.VendName
          Tr2Post(1).LDesc = APIED.INVDESC
          Tr2Post(1).Ref = APIED.InvNum
          Tr2Post(1).DrAmt = APIED.Dist(AcctDist).DAMT
          Tr2Post(1).CrAmt = 0
          Tr2Post(1).Src = MSrc$                '"AP" + ConvDateStr$(DATE$)
          Tr2Post(1).Marked = False
          Put AP2Post, RecordNum, Tr2Post(1)
          '*******************************************************************
          '--If there was a purchase order create a
          '--potrans record to liquidate encumbrance.
          If GotPOs Then
            If APIED.Dist(AcctDist).DISTNUM < 0 Or APIED.Dist(AcctDist).DISTNUM > NumDistRecs& Then
              GoSub LogBadDistTrans
            End If
            If APIED.Dist(AcctDist).DACODE = "T" Then
              
            Get APDistFile, APIED.Dist(AcctDist).DISTNUM, APDistRec(1)
            If APDistRec(1).DistStat <> "T" And APDistRec(1).DistStat <> "L" Then
              'This is to keep records in case errors during pre-post
              'so can get file back to where it was
              PTcnt = PTcnt + 1
              ReDim Preserve PTRecs(1 To PTcnt) As Long
              PTRecs(PTcnt) = APIED.Dist(AcctDist).DISTNUM
              '
              'This is to allow mulitple uses of PO distributions but to
              'unencumber only once. T is tagged, R is reused.
              'update apdistfile with T so will know has been used until
              'post complete, then will show as liquidated.
              'This way creates Trans for potrans only 1st time.
              APDistRec(1).DistStat = "T"
              Put APDistFile, APIED.Dist(AcctDist).DISTNUM, APDistRec(1)
              POIFRecNum = POIFRecNum + 1
              Tr2Post(1).AcctNum = APIED.Dist(AcctDist).DACN
              Tr2Post(1).TRDATE = APIED.DISTDATE
              Tr2Post(1).Desc = APIED.VendName
              Tr2Post(1).LDesc = APIED.INVDESC
              Tr2Post(1).Ref = APIED.InvNum
              Tr2Post(1).DrAmt = 0
              'FOR PO'S USE ORIGINAL AMT FROM PO NOT AMT ON INVOICE! 12-12-03
              Tr2Post(1).CrAmt = APDistRec(1).DistAmt
              'Tr2Post(1).CrAmt = APIED.Dist(AcctDist).DAMT
              Tr2Post(1).Src = MSrc$              '"PO" + ConvDateStr$(DATE$)
              Tr2Post(1).Marked = False
              Put PO2Post, POIFRecNum, Tr2Post(1)
            Else
    'change invoice edit trans code to show already used dist from PO
              APIED.Dist(AcctDist).DACODE = "R"
              Put APEditFile, Icnt, APIED
            End If
          End If
          'make other side of entry
  '''' *************************************************************
          '--IF PO then Create an interface rec for potrans
             If APIED.Dist(AcctDist).DACODE = "T" Then
              POIFRecNum = POIFRecNum + 1
              Tr2Post(1).AcctNum = Left$(APIED.Dist(AcctDist).DACN, GLFundLen) + EncAcct$
              Tr2Post(1).TRDATE = APIED.DISTDATE
              Tr2Post(1).Desc = APIED.VendName
              Tr2Post(1).LDesc = APIED.INVDESC
              Tr2Post(1).Ref = APIED.InvNum
              Tr2Post(1).DrAmt = APIED.Dist(AcctDist).DAMT
              Tr2Post(1).CrAmt = 0
              Tr2Post(1).Src = MSrc$
              Tr2Post(1).Marked = False
              Put PO2Post, POIFRecNum, Tr2Post(1)
             End If
          End If
          '*******************************************************************
          For Fund = 1 To NumFunds              '--Add distribution to the pro
            FundNum$ = Left$(APIED.Dist(AcctDist).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              TrFundSum#(Fund) = Round#(TrFundSum#(Fund) + APIED.Dist(AcctDist).DAMT)
              Exit For
            End If
          Next
        End If  '--test for blank distribution line
      Next      '--Acct'g Distribution
      '--Make the A/P Credits
      For Fund = 1 To NumFunds
        If TrFundSum#(Fund) <> 0 Then
          RecordNum = RecordNum + 1
          Tr2Post(1).AcctNum = FundList$(Fund) + APAcct$
          Tr2Post(1).TRDATE = APIED.DISTDATE   'APEdit.INVDATE
          Tr2Post(1).Desc = APIED.VendName
          Tr2Post(1).LDesc = APIED.INVDESC
          Tr2Post(1).Ref = APIED.InvNum
          Tr2Post(1).DrAmt = 0
          Tr2Post(1).CrAmt = TrFundSum#(Fund)
          Tr2Post(1).Src = MSrc$                '"AP" + ConvDateStr$(DATE$)
          Tr2Post(1).Marked = False
          Put AP2Post, RecordNum, Tr2Post(1)
          '**************************************************************
        End If  '--Fund summary <> 0
      Next      '--fund check for balance
    End If      '--not deleted
  Next          'transaction
  Close

  '--Post Distributions to General Ledger Accts
  Post2GL "APINVIF.dat", BadGLTrans%, frmPostInvoices, False
  'BadGLTrans = 1  used to test error report
  If BadGLTrans > 0 Then
     Call MainLog("APInvPost Errors Procedure Halted")
     MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
     ReportFile$ = "TempLog.PRN"
     frmReportOpt.Show 1
     If rptopt = 1 Then
      ARptErrorLog.GetName ReportFile$
      ARptErrorLog.startrpt
     ElseIf rptopt = 2 Then
      ViewPrint ReportFile$, "Error Log"
     End If
     If Exist("APInvif.dat") Then
        Kill "APINVIF.DAT"
     End If
     If Exist("POInvif.dat") Then
        Kill "POINVIF.DAT"
     'IF BOO-BOO NEED TO FIX INVOICE EDIT FILE BACK THE WAY IT WAS
     'ALSO DISTRIBUTION FILE
     'SO WILL CREATE CORRECT ENTRIES NEXT GO-AROUND IF RETRY
        GoSub StartFresh
     '
     End If
     frmCitiCancel.Show
     Unload frmPostInvoices
     Unload frmInvProcessMenu
     Exit Sub
  End If
''
  '--Post PO liquidations
  If Exist("POINVIF.dat") Then
    Post2PO "POINVIF.dat", BadPOTrans%, frmPostInvoices, False
  End If
    '--Tell user if we have any problems
  If BadPOTrans > 0 Then
     Call MainLog("APInvPost PO Errors - Halted")
     MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
     ReportFile$ = "TempLog.PRN"
     frmReportOpt.Show 1
     If rptopt = 1 Then
       ARptErrorLog.GetName ReportFile$
       ARptErrorLog.startrpt
     ElseIf rptopt = 2 Then
       ViewPrint ReportFile$, "Error Log"
     End If
     If Exist("APInvif.dat") Then
        Kill "APINVIF.DAT"
     End If
     If Exist("POInvif.dat") Then
        Kill "POINVIF.DAT"
     'IF BOO-BOO NEED TO Fix Dist and APied FILE
     'SO Can CREATE CORRECT ENTRIES NEXT GO-AROUND IF RETRY
        GoSub StartFresh
     '
     End If
     frmCitiCancel.Show
     Unload frmPostInvoices
     Unload frmInvProcessMenu
     Exit Sub
  End If
'  If BadVendor > 0 Then
'    MsgBox "Error: Unable to locate vendor. Review Posting Log for details.", vbOKOnly, "Invoice(s) NOT posted."
'  End If
  Post2GL "APINVIF.dat", BadGLTrans%, frmPostInvoices, True

  If BadGLTrans > 0 Then
    Call MainLog("APInvPost Errors Dist NOT POsted.")
    MsgBox "Error: Invoice Distribution(s) NOT posted. Review GL Posting Log.", vbOKOnly, "GL Account Error"
    GLLogFileName = "GLlog.dat"
    ReportFile$ = "GLlog.dat"
    frmReportOpt.Show 1
    If rptopt = 1 Then
      ARptErrorLog.GetName ReportFile$
      ARptErrorLog.startrpt
    ElseIf rptopt = 2 Then
      ViewPrint ReportFile$, "Posting Log"
    End If

  End If
  If Exist("POINVIF.dat") Then
    Post2PO "POINVIF.dat", BadPOTrans%, frmPostInvoices, True
  End If

  If BadPOTrans > 0 Then
    Call MainLog("APInvPost Errors PO Liquidation Not Posted.")
    MsgBox "Error: BadPOTrans, Review PO Posting Log.", vbOKOnly, "PO liquidation(s) NOT posted."
    POLogFileName = "POlog.dat"
    ReportFile$ = "POlog.dat"
     frmReportOpt.Show 1
     If rptopt = 1 Then
       ARptErrorLog.GetName ReportFile$
       ARptErrorLog.startrpt
     ElseIf rptopt = 2 Then
       ViewPrint ReportFile$, "Posting Log"
     End If
 
  End If
  Close
  OpenAPEditFile APEditFile, NumEdTrans
  OpenAPLedgerFile APLedgerFile, NumLdgTran&, LedgerRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  For cnt = 1 To NumEdTrans   '1  '--Post transaction to A/P Ledger, update link
    Get APEditFile, cnt, APIED
    If Not APIED.DelFlag Then
      VendRecNum = APIED.VRecNum
      'PrintHelp "Posting Vendor: " + Str$(VendRecNum)
      If VendRecNum > 0 Then
        PostCnt = PostCnt + 1

        NumLdgTran& = NumLdgTran& + 1
        '--update vendor to transaction link
        OpenVendorFile VendorFile, NumVRecs
        Get VendorFile, VendRecNum, Vendor
        PrevVendTr& = Vendor.LastTran
        FrstVendTR& = Vendor.FrstTran
        If FrstVendTR& = 0 Then
          Vendor.LastTran = NumLdgTran&
          Vendor.FrstTran = NumLdgTran&
        Else
          Vendor.LastTran = NumLdgTran&

          Get APLedgerFile, PrevVendTr&, APLedgerRec(1)
          APLedgerRec(1).NextTrans = NumLdgTran&
          Put APLedgerFile, PrevVendTr&, APLedgerRec(1)
        End If
        Put VendorFile, VendRecNum, Vendor
        Close VendorFile
        '--write transaction to apledger for Invoice Entry
        APLedgerRec(1).VRecNum = VendRecNum
        APLedgerRec(1).VendorCode = APIED.Vendor
        APLedgerRec(1).TRDATE = APIED.InvDate
        APLedgerRec(1).DOCNum = APIED.InvNum
        APLedgerRec(1).PONum = APIED.PONum
        APLedgerRec(1).MPONum = APIED.MPONum
        APLedgerRec(1).DueDate = APIED.DueDate
        APLedgerRec(1).TRCode = 1
        APLedgerRec(1).PAYCODE = Val(APIED.PAYCODE)
        APLedgerRec(1).GLDistDate = APIED.DISTDATE

        '071398 'added Tax amount into invoice total
        'APLedgerRec(1).Amt = APEdit.INVAMT
        APIED.GRANDTOT = Round#(APIED.InvAmt + APIED.STAXAMT + APIED.CTAXAMT)

        '072298
        APLedgerRec(1).Amt = APIED.GRANDTOT
        '071398
        APLedgerRec(1).NextTrans = 0
        APLedgerRec(1).FrstDist = NumDistRecs& + 1
        APLedgerRec(1).Get1099 = APIED.Get1099
        APLedgerRec(1).Comment = APIED.INVDESC '--New File format
        APLedgerRec(1).PSLFlag = APIED.PSLFlag
        '072298
        'APLedgerRec(1).TaxAmt = APEdit.TaxTotal
        APLedgerRec(1).Bankcode = 0
        APLedgerRec(1).TaxAmt = Round#(APIED.STAXAMT + APIED.CTAXAMT)
        For DistCnt = 1 To 36  '2 'LastDist
         If Len(QPTrim$(APIED.Dist(DistCnt).DACN)) > 0 And APIED.Dist(DistCnt).DAMT <> 0 Then
          ReDim APDistRec(1) As APDistRecType
          NumDistRecs& = NumDistRecs& + 1
          APDistRec(1).APLedgerRec = NumLdgTran&
          APDistRec(1).DistAcctRec = APIED.Dist(DistCnt).DACREC
          APDistRec(1).DistAcctNum = APIED.Dist(DistCnt).DACN
          APDistRec(1).DistAmt = APIED.Dist(DistCnt).DAMT
          APDistRec(1).NextDist = NumDistRecs& + 1
          Put APDistFile, NumDistRecs&, APDistRec(1)
          '--could put glupdate here


         End If
        Next     '2
    'rewrite the last valid distrubtion to indicate
        APDistRec(1).NextDist = 0               'Last distrubtion for this invoice
        Put APDistFile, NumDistRecs&, APDistRec(1)
    '--update the last distribution pointer in apledger.dat
        APLedgerRec(1).LastDist = NumDistRecs&
        Put APLedgerFile, NumLdgTran&, APLedgerRec(1)
    'Changed to look for up to "6" PO's
    'For POCnt = 1 To 6      '--Set the PO's Flag to closed
        If APIED.POAPLRecNum > 0 Then
          Get APLedgerFile, APIED.POAPLRecNum, APLedgerRec(1)

           If APIED.POFLAG = 1 Then
            APLedgerRec(1).TRCode = -4
            Put APLedgerFile, APIED.POAPLRecNum, APLedgerRec(1)
           End If
           For CntD = 1 To 36
            If APIED.Dist(CntD).DACODE = "T" Or APIED.Dist(CntD).DACODE = "R" Then
              Get APDistFile, APIED.Dist(CntD).DISTNUM, APDistRec(1)
              APDistRec(1).DistStat = "L"
              Put APDistFile, APIED.Dist(CntD).DISTNUM, APDistRec(1)
            End If
           Next
           Used = 0
           Dist = 0
           Nextone& = APLedgerRec(1).FrstDist
           Do Until Nextone& = 0
            Get APDistFile, Nextone&, APDistRec(1)
            If APDistRec(1).DistStat = "L" Then
              Used = Used + 1
            End If
            Dist = Dist + 1
            Nextone& = APDistRec(1).NextDist
           Loop
           If Used = Dist Then
            APLedgerRec(1).TRCode = -4
            Put APLedgerFile, APIED.POAPLRecNum, APLedgerRec(1)
           End If
        End If


        Else      '--could not find vendor, Mark trans as deleted and log it.
          BadVendor = BadVendor + 1
          APIED.DelFlag = True
          Put APEditFile, cnt, APIED
          GoSub LogBadTrans

        End If    '--test for good vendor
      End If      '--test for not deleted trans
  Next  '1
  Close
  OpenAcctFile AcctFileNum, NumAccts
  OpenAPEditFile APEditFile, NumEdTrans
  OpenAPLedgerFile APLedgerFile, NumLdgTran&, LedgerRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  For cnt = 1 To NumEdTrans
    Get APEditFile, cnt, APIED
    If Not APIED.DelFlag Then
      If APIED.POAPLRecNum > 0 Then
       For CntD = 1 To 36
        If APIED.Dist(CntD).DACODE = "T" Then
          Get APDistFile, APIED.Dist(CntD).DISTNUM, APDistRec(1)
          TempDr = 0
          TempCr = APDistRec(1).DistAmt
          TempAcct = APIED.Dist(CntD).DACREC
          If TempAcct > 0 Then
           Get AcctFileNum, TempAcct, Acct
           Select Case Acct.Typ
            Case "A", "E"                 'asset, exp accts
              Acct.Encumb = Round#(Acct.Encumb + TempDr - TempCr)
                Put AcctFileNum, APIED.Dist(CntD).DACREC, Acct
            Case "L", "R"                 'liab, rev accts
              Acct.Encumb = Round#(Acct.Encumb + TempCr - TempDr)
                Put AcctFileNum, APIED.Dist(CntD).DACREC, Acct
           End Select
          
          End If
        End If
       Next
      End If
    End If
  Next
   Close

  '--All Done with post to apledger & apdist
'  If BadVendor = 0 Then
'    Print #LogFile, "No Posting Errors. Transactions Posted: "; PostCnt
'  End If
  SetAttr "APIED.DAT", vbNormal
  KillFile "APIED.DAT"
  KillFile "APINVIF.DAT"
  KillFile "POINVIF.DAT"
  KillFile "APIEDSV.DAT"
  Call MainLog("APInvPost Completed")
  MsgBox "Posting Procedure Completed", vbOKOnly, "Invoice Posting"

  Exit Sub
StartFresh:
'This is for setting flags in distribution file back to " " blank so
'will recreate the po interface file for potrans, also copy ap edit file
'back to original state.
  Dim upper As Integer
  SH_CopyFile "APIEDSV.DAT", "APIED.DAT"
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  upper = UBound(PTRecs)
  For cnt = 1 To upper
    Get APDistFile, PTRecs(cnt), APDistRec(1)
    If APDistRec(1).DistStat = "T" And APDistRec(1).DistStat <> "L" Then
      APDistRec(1).DistStat = " "
      Put APDistFile, PTRecs(cnt), APDistRec(1)
    End If
  Next          'transaction
  Close
Return
LogBadDistTrans:
  Unload FrmShowPctComp
  Call MainLog("APInvPost Dist Errors - Halted")
  ReportFile$ = "TempDist.PRN"
  RptFile = FreeFile
  Open ReportFile$ For Output As RptFile
  Print #RptFile, "Error With Distribution Listed below, please delete inv, re-enter and try again."
  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 1) = APIED.Vendor
  Mid$(ToPrint$, 12) = Format(DateAdd("d", (APIED.InvDate), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 24) = APIED.InvNum
  Mid$(ToPrint$, 35) = Str$(Round(APIED.InvAmt))
  Mid$(ToPrint$, 46) = APIED.Dist(AcctDist).DACN
  Print #RptFile, ToPrint$
  Close
     MsgBox "Errors Were Found, Please review report before trying again. Contact Software Support if questions.", vbOKOnly, "Errors"
     frmReportOpt.Show 1
     If rptopt = 1 Then
       ARptErrorLog.GetName ReportFile$
       ARptErrorLog.startrpt
     ElseIf rptopt = 2 Then
       ViewPrint ReportFile$, "Error Log"
     End If
     If Exist("APInvif.dat") Then
        Kill "APINVIF.DAT"
     End If
     If Exist("POInvif.dat") Then
        Kill "POINVIF.DAT"
     '
     End If
     'frmCitiCancel.Show
     Unload frmPostInvoices
     'Unload frmInvProcessMenu
     Exit Sub

  Return

LogBadTrans:
  Print #LogFile, "Unable to find Vendor. Transaction deleted."
  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 1) = APIED.Vendor
  Mid$(ToPrint$, 12) = Format(DateAdd("d", (APIED.InvDate), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 24) = APIED.InvNum
  Mid$(ToPrint$, 44) = Str$(Round(APIED.InvAmt))
  Print #LogFile, ToPrint$
  Return
OutNow:
  SetAttr ("APIED.DAT"), vbNormal
  Exit Sub
ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "AP", "Posting Invoices", Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select
  
Proc_Exit:
  Close
  Exit Sub
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

