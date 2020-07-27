Attribute VB_Name = "TaxCommon"
Option Explicit
'  Public MortArr() As Long '8/30/06 used in printing tax bills
'  Public MortCnt As Long '8/30/06 used in printing tax bills
'  Public ZipArr() As Long '8/30/06 used in printing tax bills
'  Public ZipCnt As Long '8/30/06 used in printing tax bills
  Public ComputerName As String
  Public CurrCitiPath As String
  Public StartPath As String
  Public doAlign As Boolean
  Public alnRpt$
  Public RecpDef As Integer
  Public RecpPort As String
  Public NumOfAligns As Integer
  Public BadMaskFlag As Boolean
  Public ScreenW As Long
  Public GTestOK() As Boolean 'used on frmTaxBillGLSetUp
  Public GTestNums() As String 'used on frmTaxBillGLSetUp
  Public GTestDbCrt() As String 'used on frmTaxBillGLSetUp
  Public GTestDesc() As String 'used on frmTaxBillGLSetUp
  Public GCustNum As Long
  Public AddCust As Boolean
  Public EditCust As Boolean
  Public PayDate As String
  Public Twiddle As String
  Public OperNum As Integer
  Public DelAbs As Boolean
  Public PayEntry As Boolean
  Public GPayNum As Integer
  Public EditFlag As Boolean 'used in entering/editing payment transactions
  Public THistRpt As Boolean
  Public BillTrans() As Long
  Public BillDate() As Integer
  Public BillCnt As Integer
  Public RateTblRec As Integer
  Public RealCnt
  Public RealRecs() As Long 'replaced screen local on frmTaxRealProp 8/16/06
  Public PersRecs() As Long 'replaced screen local on frmTaxPersProp 8/16/06
  Public RealProp() As Long
  Public RealRev() As Integer
  Public RptOpt As Integer 'used to determine the type of reports; graphic or text
  Public CycleCnt As Long
  Public CycleUsersName() As String
  Public CycleUsersAcct() As Long
  Public CountyCnt As Long
  Public CountyUsersName() As String
  Public CountyUsersAcct() As Long
  Public ClassCnt() As Long
  Public ClassUsersName() As String
  Public ClassUsersAcct() As Long
  Public ClassRealPin() As String
  Public CntrlDef As Integer
  Public ThisMRec As Integer
  Public FromTX As Boolean
  Public AcctNumList() As Long
  Public NumPreBal() As Double
  Public RefNumCnt As Long
  Public AcctNameList() As Long
  Public RefNameCnt As Long
  Public NamePreBal() As Double
  Public TypeCnt() As Long
  Dim BigNo As Long

  'The modules below are located in NCWORK on the C Drive
  'TAXADVX = CalcInterest (see also TAXINT below)
  'TAXCHIST = Tax customer history report
  'TAXBILL = Tax pre-billing and bill posting
  'TAXABSRT = Move Property Abstract and selecting personal or real estate to delete
  'TAXADRPT = Advertising related transactions, lookups and posting
  'TAXAJUST = Adjustment related entries
  'TAXAUX01 = Various subs...including Lookup
  'TAXAUX02 = Various subs
  'TAXCUST = Adding/Editing customers and their properties
  'TAXDNOTE = Printing late notices
  'TAXINQ = Contains numerous reports including master reports.
  'TAXINT = Calculates tax interest...look for 'APPLY INTEREST' under TAX BILLING FUNCTIONS
  'TAXLNOTE = Late notice code
  'TAXMCHRG = Adjustment transactions
  'TAXMENU = Main menu...also has a 'CheckInterestInfo' sub that is a reminder
  '          to the user to apply interest for the month if they have not already done so
  'TAXPAY = No subs in this module
  'TAXPAY1 = Payment entry, operator verification, tax payment posting, operator editing/adding
  'TAXPAYMN = Looks like a version of TAXPAY1
  'TAXPRNTB = Printing and reprinting tax bills and bills printed report.
  'TAXRE = Many reports...Transaction Journal, Abstracts, Abstract Listing, Ad Listing
  '        Customer Inquiry, Customer Listing, Late Listing, Master Value List
  '        Senior Citizens List
  'TAXREPRT = Contains duplicates subs as TAXRE (probably some differences in code) plus the following:
  '           Mortgage Code Report, MakeAdCostReport, Tax Payment Report.
  'TAXSETUP = MortCode (Mortgage setup?), Relink abstracts, Relink Transactions, GL Tax Billing Setup,
  '           GL Tax Payment Setup, Tax Main Setup
  'TAXINTR = Another adjustment (see TAXMCHRG).
  'TXEXPORT = Looks like the code to create a magnetic file.
  
  Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
  
      Public Const TaxCustFile = "TAXCUST.DAT"
      Public Const CustNameIdxFile = "TAXNMIDX.DAT"
      Public Const SrchNameIdxFile = "SRCHNMIDX.DAT"
      Public Const TaxPayFileName = "TXPAYMNT.DAT"
      Public Const CustPinFile = "TAXCPIN.DAT"
      Public Const TaxPropFile = "TAXPROP.DAT"
      Public Const TaxPersFile = "TAXPERS.DAT"
      Public Const TaxMCodeFile = "TAXMORT.DAT"
      Public Const TaxPenHandling = "TAXPEN.DAT"
'      Public Const TaxPersIdxFile = "TAXPPIN.IDX"
'      Public Const TaxRealIdxFile = "TAXRPIN.IDX"
      Public Const RealOptSearch = "TXROPTSH.DAT"
      Public Const PersOptSearch = "TXPOPTSH.DAT"
      Public Const CustOptSearch = "TXCOPTSH.DAT"
      Public Const SocSecIdxFile = "TXSSIDX.DAT"
      Public Const RealHistFile = "TXRLHIST.DAT"
      Public Const TaxManualBill = "C:\CPWork\TAXMEDIT.DAT" 'added C:\ on 5.16.07
      Public Const TempTaxBillAddOn = "TMPBLADD"
'      Public Const TempBillRecs = "TMPBLREC.DAT"
      Public Const TempBillRecs = "C:\CPWork\TMPBLREC.DAT" 'added C:\ on 11/28/06
      Public Const TaxPersPINFile = "TAXPPIN.DAT"
      Public Const TaxRealPINFile = "TAXRPIN.DAT"
      Public Const TaxBillFile = "TAXTBILL.DAT" 'added C:\ on 5.16.07...took off on 6.21.07
      Public Const TaxBillOPFile = "TAXOPBL.DAT" '5.16.07 ...took off on 6.21.07
      Public Const TaxBillInfoFile = "TAXBINFO.DAT" '5.16.07...took off on 6.21.07
      Public Const TaxIntFile = "TAXINT.DAT"
      Public Const TaxAdvFile = "TAXADV.DAT"
      Public Const TaxTownships = "TXTWNSHP.DAT"
      Public Const TaxPreRptFile = "TAXPREBL.RPT"
      Public Const TaxTransFile = "TAXTRANS.DAT"
      Public Const InterestReportFile = "TAXINT.RPT"
      Public Const TaxAdvReportFile = "TAXADV.RPT"
      Public Const TaxRateTableFile = "TXRTTBLS.DAT"
      Public Const TaxManualBillList = "TXMANLST.DAT"
    'Virginia Added--------------------------------
      Public Const RETaxCustFile = "RETXCUST.DAT"
      Public Const PPTaxCustFile = "PPTXCUST.DAT"
      Public Const RECustPinFile = "VARETXPN.DAT"
      Public Const PPCustPinFile = "VAPPTXPN.DAT"
      Public Const PPTaxBillFile = "TAXPBILL.DAT"
      Public Const RETaxBillFile = "TAXRBILL.DAT"
      Public Const PPTaxPreRptFile = "TXPPREBL.RPT"
      Public Const RETaxPreRptFile = "TXRPREBL.RPT"
      Public Const RETaxIntFile = "TAXREINT.DAT"
      Public Const RETaxPenFile = "TAXREPEN.DAT"
      Public Const PPTaxIntFile = "TAXPPINT.DAT"
      Public Const PPTaxPenFile = "TAXPPPEN.DAT"
'--------------------------------------------------
      Public Const TaxSetupName = "TAXSETUP.DAT"
      Public Const PerTaxName = "TAXPERS.DAT"
      Public Const TaxPropName = "TAXPROP.DAT"
      Public Const MortCodeName = "TAXMORT.DAT"
      Public Const TxGLInterBill = "TAXGLBAC.DAT"
      Public Const TxGLInterPay = "TAXGLACT.DAT"
      Public Const AcctFileName = "GLACCT.DAT"
      Public Const JGLAcctIdxFile = "GLACCT.IDX"
      Public Const InternalPinFile = "TAXINPIN.DAT"
      Public Const MessageName = "TAXMESS.DAT"
      Public Const LateLtrText = "TXLATLTR.DAT"
      Public Const LateLtrPrint = "TXLLPRN.DAT"
      Public Const TaxBill1Name = "TaxBil1.DAT"
      Public Const TBillExpPers = "TBXPERS.DAT"
      Public Const TBillExpReal = "TBXREAL.DAT"
      Public Const TaxIntTickler = "TAXINTCK.DAT"
      Public Const NewTaxTransFile = "NEWTAXTRANS.DAT"
      Public Const ZipIdxFile = "ZIPIDX.DAT"
      Public Const MortIdxFile = "MORTIDX.DAT"
      Public Const CntyEditFile = "CNTYEDIT.DAT"
Public Sub OpenCountyEditFile(CEHandle As Integer)
  Dim CERecLen As Integer
  Dim CERec As AllowCountyEdit
  CERecLen = Len(CERec)
  CEHandle = FreeFile
  Open CntyEditFile For Random Shared As CEHandle Len = CERecLen
End Sub
Public Sub OpenOldPostedReprintFile(PRHandle As Integer, NumOfPRRecs As Long, ThisFile$)
  Dim PRRecLen As Integer
  Dim PRRec As TaxBillTypeOld
  PRRecLen = Len(PRRec)
  PRHandle = FreeFile
  Open ThisFile For Random Shared As PRHandle Len = PRRecLen
  NumOfPRRecs = LOF(PRHandle) / PRRecLen
End Sub
Public Sub OpenMortIdxFile(MortHandle As Integer, NumOfMRecs As Long)
  Dim MortLen As Integer
  Dim MortRec As BillPrintMortIdxType
  MortLen = Len(MortRec)
  MortHandle = FreeFile
  Open MortIdxFile For Random Shared As MortHandle Len = MortLen
  NumOfMRecs = LOF(MortHandle) / MortLen
End Sub
Public Sub OpenZipIdxFile(ZipHandle As Integer, NumOfZRecs As Long)
  Dim ZipLen As Integer
  Dim ZipRec As BillPrintZipIdxType
  ZipLen = Len(ZipRec)
  ZipHandle = FreeFile
  Open ZipIdxFile For Random Shared As ZipHandle Len = ZipLen
  NumOfZRecs = LOF(ZipHandle) / ZipLen
End Sub
Public Sub OpenNewTaxTransFile(TransHandle As Integer)
  Dim TransTaxLen As Integer
  Dim TransTaxRec As TaxTransactionType
  TransTaxLen = Len(TransTaxRec)
  TransHandle = FreeFile
  Open NewTaxTransFile For Random Shared As TransHandle Len = TransTaxLen
End Sub
Public Sub OpenPostedReprintFile(PRHandle As Integer, NumOfPRRecs As Long, ThisFile$)
  Dim PRRecLen As Integer
  Dim PRRec As TaxBillType
  PRRecLen = Len(PRRec)
  PRHandle = FreeFile
  Open ThisFile For Random Shared As PRHandle Len = PRRecLen
  NumOfPRRecs = LOF(PRHandle) / PRRecLen
End Sub
Public Sub OpenTxIntTickFile(TxIntTickHandle As Integer)
  Dim TxIntTickRecLen As Integer
  Dim TxIntTickRec As TaxInterestDateType
  TxIntTickRecLen = Len(TxIntTickRec)
  TxIntTickHandle = FreeFile
  Open TaxIntTickler For Random Shared As TxIntTickHandle Len = TxIntTickRecLen
End Sub
Public Sub OpenTxBillExpRealFile(TxBillExpRealHandle As Integer)
  Dim TxBillExpRealRecLen As Integer
  Dim TxBillExpRealRec As TaxBillExportRealType
  TxBillExpRealRecLen = Len(TxBillExpRealRec)
  TxBillExpRealHandle = FreeFile
  Open TBillExpReal For Random Shared As TxBillExpRealHandle Len = TxBillExpRealRecLen
End Sub
Public Sub OpenTxBillExpPersFile(TxBillExpPersHandle As Integer)
  Dim TxBillExpPersRecLen As Integer
  Dim TxBillExpPersRec As TaxBillExportPersType
  TxBillExpPersRecLen = Len(TxBillExpPersRec)
  TxBillExpPersHandle = FreeFile
  Open TBillExpPers For Random Shared As TxBillExpPersHandle Len = TxBillExpPersRecLen
End Sub
      
Public Sub OpenTxBill1File(TxBill1Handle As Integer)
  Dim TxBill1RecLen As Integer
  Dim TxBill1Rec As TxBill1DefaultsType
  TxBill1RecLen = Len(TxBill1Rec)
  TxBill1Handle = FreeFile
  Open TaxBill1Name For Random Shared As TxBill1Handle Len = TxBill1RecLen
End Sub

Public Sub OpenLatePrnFile(LatePrnHandle As Integer, NumOfLatePrnRecs As Long)
  Dim LatePrnRecLen As Integer
  Dim LatePrnRec As LateListPrintType
  LatePrnRecLen = Len(LatePrnRec)
  LatePrnHandle = FreeFile
  Open LateLtrPrint For Random Shared As LatePrnHandle Len = LatePrnRecLen
  NumOfLatePrnRecs = LOF(LatePrnHandle) / LatePrnRecLen
End Sub
Public Sub OpenLateLtrFile(LateHandle As Integer)
  Dim LateRecLen As Integer
  Dim LateRec As TAXLateLetterType
  LateRecLen = Len(LateRec)
  LateHandle = FreeFile
  Open LateLtrText For Random Shared As LateHandle Len = LateRecLen
End Sub
Public Sub OpenSocSecIdxFile(SSHandle As Integer, NumOfSSFiles As Long)
  Dim SSRecLen As Integer
  Dim SSRec As SocSecIdxType
  SSRecLen = Len(SSRec)
  SSHandle = FreeFile
  Open SocSecIdxFile For Random Shared As SSHandle Len = SSRecLen
  NumOfSSFiles = LOF(SSHandle) / SSRecLen
End Sub
Public Sub OpenRealOptSearchFile(ROSHandle As Integer, NumOfROSFiles As Long)
  Dim ROSRecLen As Integer
  Dim ROSRec As OptRealIdxType
  ROSRecLen = Len(ROSRec)
  ROSHandle = FreeFile
  Open RealOptSearch For Random Shared As ROSHandle Len = ROSRecLen
  NumOfROSFiles = LOF(ROSHandle) / ROSRecLen
End Sub
Public Sub OpenPersOptSearchFile(POSHandle As Integer, NumOfPOSFiles As Long)
  Dim POSRecLen As Integer
  Dim POSRec As OptPersIdxType
  POSRecLen = Len(POSRec)
  POSHandle = FreeFile
  Open PersOptSearch For Random Shared As POSHandle Len = POSRecLen
  NumOfPOSFiles = LOF(POSHandle) / POSRecLen
End Sub
Public Sub OpenCustOptSearchFile(COSHandle As Integer, NumOfCOSFiles As Long)
  Dim COSRecLen As Integer
  Dim COSRec As OptCustIdxType
  COSRecLen = Len(COSRec)
  COSHandle = FreeFile
  Open CustOptSearch For Random Shared As COSHandle Len = COSRecLen
  NumOfCOSFiles = LOF(COSHandle) / COSRecLen
End Sub
Public Sub OpenAdvColRecFile(AdvColRecHandle As Integer, NumOfAdvColRecFiles As Long)
  Dim AdvColRecLen As Integer
  Dim AdvColRec As InterestRecType
  AdvColRecLen = Len(AdvColRec)
  AdvColRecHandle = FreeFile
  Open TaxAdvFile For Random Shared As AdvColRecHandle Len = AdvColRecLen
  NumOfAdvColRecFiles = LOF(AdvColRecHandle) / AdvColRecLen
End Sub
Public Sub OpenTaxManualBillList(TaxManBillListHandle As Integer, TaxManBillListCnt As Long)
  Dim TaxManBillListLen As Integer
  Dim TaxManBillListRec As ManualTaxListType
  TaxManBillListLen = Len(TaxManBillListRec)
  TaxManBillListHandle = FreeFile
  Open TaxManualBillList For Random Shared As TaxManBillListHandle Len = TaxManBillListLen
  TaxManBillListCnt = LOF(TaxManBillListHandle) / Len(TaxManBillListRec)
End Sub
Public Sub OpenTaxManualBillFile(TaxManualBillHandle As Integer, TaxManualBillCnt As Integer)
  Dim TaxManualBillLen As Integer
  Dim TaxManualBillRec As TaxMTransactionType
  TaxManualBillLen = Len(TaxManualBillRec)
  TaxManualBillHandle = FreeFile
  Open TaxManualBill For Random Shared As TaxManualBillHandle Len = TaxManualBillLen
  TaxManualBillCnt = LOF(TaxManualBillHandle) / Len(TaxManualBillRec)
End Sub
Public Sub OpenRealHistFile(RealHistHandle As Integer, RealHistCnt As Long)
  Dim RealHistLen As Integer
  Dim RealHistRec As RealHistoryType
  RealHistLen = Len(RealHistRec)
  RealHistHandle = FreeFile
  Open RealHistFile For Random Shared As RealHistHandle Len = RealHistLen
  RealHistCnt = LOF(RealHistHandle) / Len(RealHistRec)
End Sub
Public Sub OpenTaxRateTables(RateTablesHandle As Integer, RateTablesCnt As Integer)
  Dim RateTablesLen As Integer
  Dim RateTablesRec As OptRevRateTablesType
  RateTablesLen = Len(RateTablesRec)
  RateTablesHandle = FreeFile
  Open TaxRateTableFile For Random Shared As RateTablesHandle Len = RateTablesLen
  RateTablesCnt = LOF(RateTablesHandle) / Len(RateTablesRec)
End Sub
Public Sub OpenTaxMessage(MessHandle As Integer, MsgCnt As Integer)
  Dim MessLen As Integer
  Dim MessRec As TaxMessRecType
  MessLen = Len(MessRec)
  MessHandle = FreeFile
  Open MessageName For Random Shared As MessHandle Len = MessLen
  MsgCnt = LOF(MessHandle) / Len(MessRec)
End Sub
Public Sub OpenTownshipFile(TownshipHandle As Integer, NumOfTownships As Integer)
  Dim TownshipLen As Integer
  Dim TownshipRec As TownshipType
  TownshipLen = Len(TownshipRec)
  TownshipHandle = FreeFile
  Open TaxTownships For Random Shared As TownshipHandle Len = TownshipLen
  NumOfTownships = LOF(TownshipHandle) / Len(TownshipRec)
End Sub
Public Sub OpenTaxBillAddOn(AddOnHandle As Integer)
  Dim AddOnLen As Integer
  Dim AddOnRec As TempTaxBillAddOn
  AddOnLen = Len(AddOnRec)
  AddOnHandle = FreeFile
  Open TempTaxBillAddOn For Random Shared As AddOnHandle Len = AddOnLen
End Sub
Public Sub OpenBillInfoFile(BillInfoHandle As Integer)
  Dim BillInfoLen As Integer
  Dim BillInfoRec As TaxBillInfoType
  BillInfoLen = Len(BillInfoRec)
  BillInfoHandle = FreeFile
  Open TaxBillInfoFile For Random Shared As BillInfoHandle Len = BillInfoLen
End Sub
Public Sub OpenTempBillRecs(TempBillHandle As Integer, TempCnt As Integer)
  Dim TempBillLen As Integer
  Dim TempBillRec As TempPayList
  TempBillLen = Len(TempBillRec)
  TempBillHandle = FreeFile
  Open TempBillRecs For Random Shared As TempBillHandle Len = TempBillLen
  TempCnt = LOF(TempBillHandle) / Len(TempBillRec)
End Sub
Public Sub OpenTaxPenFile(TaxPenHandle As Integer)
  Dim TaxPenLen As Integer
  Dim TaxPenRec As PenaltyHandlingType
  TaxPenLen = Len(TaxPenRec)
  TaxPenHandle = FreeFile
  Open TaxPenHandling For Random Shared As TaxPenHandle Len = TaxPenLen
End Sub
Public Sub OpenTempPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As TaxPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  Operator$ = QPTrim$(Operator)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open "TAXCPR" + Operator$ + ".DAT" For Random Shared As PayHandle Len = PayRecLen
End Sub
Public Sub OpenPayListFile(PayListHandle As Integer, Oper As Integer)
  Dim PayListRec As PayListType
  Dim PayListRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  Operator$ = QPTrim$(Operator$)
  PayListRecLen = Len(PayListRec)
  PayListHandle = FreeFile
  Open "TAXLOP" + Operator$ + ".DAT" For Random Shared As PayListHandle Len = PayListRecLen
End Sub
Public Sub OpenPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As TaxPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open TaxPayFileName For Random Shared As PayHandle Len = PayRecLen
End Sub
Public Sub OpenPersPinFile(PersPinHandle As Integer, NumOfPersPins As Long)
  Dim PersPinLen As Integer
  Dim PersPinRec As PINSearchType
  PersPinLen = Len(PersPinRec)
  PersPinHandle = FreeFile
  Open TaxPersPINFile For Random Shared As PersPinHandle Len = PersPinLen
  NumOfPersPins = LOF(PersPinHandle) / Len(PersPinRec)
End Sub
Public Sub OpenRealPinFile(RealPinHandle As Integer, NumOfRealPins As Long)
  Dim RealPinLen As Integer
  Dim RealPinRec As PINSearchType
  RealPinLen = Len(RealPinRec)
  RealPinHandle = FreeFile
  Open TaxRealPINFile For Random Shared As RealPinHandle Len = RealPinLen
  NumOfRealPins = LOF(RealPinHandle) / Len(RealPinRec)
End Sub
Public Sub OpenRealPropFile(RealPropHandle As Integer, NumOfRealProp As Long)
  Dim RealPropLen As Integer
  Dim RealPropRec As PropertyRecType
  RealPropLen = Len(RealPropRec)
  RealPropHandle = FreeFile
  Open TaxPropFile For Random Shared As RealPropHandle Len = RealPropLen
  NumOfRealProp = LOF(RealPropHandle) / Len(RealPropRec)
End Sub
Public Sub OpenIntPinFile(IntPinHandle As Integer, NumOfIntPins As Long)
  Dim IntPinLen As Integer
  Dim IntPinRec As InternalPinType
  IntPinLen = Len(IntPinRec)
  IntPinHandle = FreeFile
  Open InternalPinFile For Random Shared As IntPinHandle Len = IntPinLen
  NumOfIntPins = LOF(IntPinHandle) / Len(IntPinRec)
End Sub
     
Public Sub OpenCustPinFile(CustPinHandle As Integer, NumOfCustPins As Long)
  Dim CustPinLen As Integer
  Dim CustPinRec As PINRecType
  CustPinLen = Len(CustPinRec)
  CustPinHandle = FreeFile
  Open CustPinFile For Random Shared As CustPinHandle Len = CustPinLen
  NumOfCustPins = LOF(CustPinHandle) / Len(CustPinRec)
End Sub
Public Sub OpenPersPropFile(PersPropHandle As Integer, NumOfPersProp As Long)
  Dim PersPropLen As Integer
  Dim PersPropRec As PersonalRecType
  PersPropLen = Len(PersPropRec)
  PersPropHandle = FreeFile
  Open TaxPersFile For Random Shared As PersPropHandle Len = PersPropLen
  NumOfPersProp = LOF(PersPropHandle) / Len(PersPropRec)
End Sub
Public Sub OpenSrchNameIdxFile(SrchNameIdxHandle As Integer, NumOfNameIdxRec As Long)
  Dim SrchNameIdxLen As Integer
  Dim SrchNameIdxRec As SrchNameIdxType
  SrchNameIdxLen = Len(SrchNameIdxRec)
  SrchNameIdxHandle = FreeFile
  Open SrchNameIdxFile For Random Shared As SrchNameIdxHandle Len = SrchNameIdxLen
  NumOfNameIdxRec = LOF(SrchNameIdxHandle) / Len(SrchNameIdxRec)
End Sub
Public Sub OpenNameIdxFile(NameIdxHandle As Integer, NumOfNameIdxRec As Long)
  Dim NameIdxLen As Integer
  Dim NameIdxRec As CustNameIdxType
  NameIdxLen = Len(NameIdxRec)
  NameIdxHandle = FreeFile
  Open CustNameIdxFile For Random Shared As NameIdxHandle Len = NameIdxLen
  NumOfNameIdxRec = LOF(NameIdxHandle) / Len(NameIdxRec)
End Sub
Public Sub OpenGLAcctFile(GLHandle As Integer, NumOfRecs As Integer)
  Dim GLRec As GLAcctRecType
  Dim GLRecLen As Integer
  GLRecLen = Len(GLRec)
  GLHandle = FreeFile
  Open AcctFileName For Random Shared As GLHandle Len = GLRecLen
  NumOfRecs = LOF(GLHandle) / Len(GLRec)
End Sub
Public Sub OpenGLIdxFile(GLHandle As Integer, NumOfRecs As Integer)
  Dim GLRec As JGLAcctIdxType
  Dim GLRecLen As Integer
  GLRecLen = Len(GLRec)
  GLHandle = FreeFile
  Open JGLAcctIdxFile For Random Shared As GLHandle Len = GLRecLen
  NumOfRecs = LOF(GLHandle) / Len(GLRec)
End Sub
Public Sub OpenTaxGLInterPay(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As TaxAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxGLInterPay For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub OpenTaxGLInterBill(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As TaxAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxGLInterBill For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub OpenTaxBillFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As TaxBillType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open TaxBillFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub
Public Sub OpenTaxBillOverPayFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As TaxTransactionType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open TaxBillOPFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub
      
Public Sub OpenMortCodeFile(MortCodeHandle As Integer, NumOfMortCodes As Integer)
  Dim MortCodeLen As Integer
  Dim MortCodeRec As MortCodeRecType
  MortCodeLen = Len(MortCodeRec)
  MortCodeHandle = FreeFile
  Open MortCodeName For Random Shared As MortCodeHandle Len = Len(MortCodeRec)
  NumOfMortCodes = LOF(MortCodeHandle) / Len(MortCodeRec)
End Sub
      
Public Sub OpenTaxPropFile(TaxPropHandle As Integer, NumOfTaxProps As Long)
  Dim TaxPropLen As Integer
  Dim TaxPropRec As PropertyRecType
  TaxPropLen = Len(TaxPropRec)
  TaxPropHandle = FreeFile
  Open TaxPropName For Random Shared As TaxPropHandle Len = Len(TaxPropRec)
  NumOfTaxProps = LOF(TaxPropHandle) / Len(TaxPropRec)
End Sub
      
Public Sub OpenTaxPersFile(PersTaxHandle As Integer, NumOfPersRecs As Long)
  Dim PersTaxLen As Integer
  Dim PersTaxRec As PersonalRecType
  PersTaxLen = Len(PersTaxRec)
  PersTaxHandle = FreeFile
  Open PerTaxName For Random Shared As PersTaxHandle Len = PersTaxLen
  NumOfPersRecs = LOF(PersTaxHandle) / Len(PersTaxRec)
End Sub
      
Public Sub OpenTaxCustFile(TaxCustHandle As Integer, NumOfTaxCustRec As Long)
  Dim TaxCustLen As Integer
  Dim TaxCustRec As TaxCustType
  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open TaxCustFile For Random Shared As TaxCustHandle Len = TaxCustLen
  NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
End Sub
Public Sub OpenInterestRecFile(InterestRecHandle As Integer, NumOfIntRecFiles As Long)
  Dim InterestRecLen As Integer
  Dim InterestRec As InterestRecType
  InterestRecLen = Len(InterestRec)
  InterestRecHandle = FreeFile
  Open TaxIntFile For Random Shared As InterestRecHandle Len = InterestRecLen
  NumOfIntRecFiles = LOF(InterestRecHandle) / InterestRecLen
End Sub
Public Sub OpenTaxSetUpFile(TaxSetUpHandle As Integer)
  Dim TaxSetUpLen As Integer
  Dim TaxSetUp As TaxMasterType
  TaxSetUpLen = Len(TaxSetUp)
  TaxSetUpHandle = FreeFile
  Open TaxSetupName For Random Shared As TaxSetUpHandle Len = TaxSetUpLen
End Sub
Public Sub OpenTaxTransFile(TaxTransHandle As Integer, NumOfTaxTransRecs As Long)
  Dim TaxTransLen As Integer
  Dim TaxTransRate As TaxTransactionType
  TaxTransLen = Len(TaxTransRate)
  TaxTransHandle = FreeFile
  Open TaxTransFile For Random Shared As TaxTransHandle Len = TaxTransLen
  NumOfTaxTransRecs = LOF(TaxTransHandle) / Len(TaxTransRate)
End Sub
  
Public Sub GetTemp()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  
  'lentemp = Len(Tempfile)
  Tempfile = FreeFile
'  Open "c:\PassTemp.dat" For Random Shared As Tempfile ' Len = lentemp
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp 2/14/08
  Get Tempfile, 1, PassTemp
  PWUser = QPTrim(PassTemp.UserName)
  PWcnt = PassTemp.usernum
  Close

End Sub

Public Sub Terminate2Shell()
   Dim UBFrmCnt As Integer
   ' Loop through the forms collection and unload each form.
   Close
   For UBFrmCnt = Forms.Count - 1 To 0 Step -1
       Unload Forms(UBFrmCnt)
   Next
   DoEvents
   End
End Sub
Public Sub Terminate()
   Dim UBFrmCnt As Integer
   
   If PWcnt = -3 Then GoTo SSPW 'Southern Software Password
   ' Loop through the forms collection and unload each form.
   ClearInUsePRReg PWcnt 'we want this intact so if another user
   'gets in payroll the "inuse" warning will pop up
   Close
SSPW:
   For UBFrmCnt = Forms.Count - 1 To 0 Step -1
      Unload Forms(UBFrmCnt)
   Next
   DoEvents
   Call KillWaste

   End
End Sub

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim Cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For Cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, Cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, Cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
  End Function

Public Function Exist(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  Dim testFile As String
  testFile = UCase$(FileName$)
  On Error GoTo FileError
  FileHandle = FreeFile
  Open testFile For Input Shared As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
  End If
  GoTo ExistExit
FileError:
  Close FileHandle
  Exist = False
  If UCase(Error) <> "FILE NOT FOUND" Then
    MsgBox "Error " & Error$ & " " & testFile, vbOKOnly, "Error"
  End If
ExistExit:
  Err.Clear
  On Error GoTo 0
  DoEvents
End Function
Public Function DirExists(ByVal strDirName As String) As Boolean
  On Error Resume Next
  
  Dim strFileName As String

  strFileName = strDirName & "\Nul"

  If (FileExists(strFileName)) Then
    DirExists = True
  Else
    DirExists = False
  End If
End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
  
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function

Public Sub GetAcctStruct(CitipakPath$, GLFundLen%, GLAcctLen%, GLDetLen%)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetupRec(1) As GLSetupRecType
  'this sub determines the lengths of each piece of the gl number...
  '(ie. 12-345-6789 breaks down to GLFundLen = 2, GLAcctLen = 3
  'and GLDetLen (Dept) = 4)...this data is used in validating
  'GL numbers before they are saved
'  StartPath = StartPath
  SetUpRecLen = Len(GLSetupRec(1))
  If Exist(QPTrim$(CitipakPath) + "GLSETUP.DAT") Then
    SetupFile = FreeFile
    Open CitipakPath + "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  ElseIf Exist(QPTrim$(CitipakPath) + "\GLSETUP.DAT") Then
    SetupFile = FreeFile
    Open QPTrim$(CitipakPath) + "\GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Else
    Exit Sub
  End If
  Get SetupFile, 1, GLSetupRec(1)
  Close SetupFile
  GLFundLen = GLSetupRec(1).FundLen
  GLAcctLen = GLSetupRec(1).AcctLen
  GLDetLen = GLSetupRec(1).DetLen
  Erase GLSetupRec
End Sub

Public Function Date2Num%(TheDate$)
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function
Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

Public Function OldRound#(n As Double)
  OldRound# = Int(n * 100 + 0.500000001) / 100
End Function

Public Sub KillFile(FileName As String)
  On Local Error Resume Next
  If Exist(FileName$) Then
    Kill FileName$
  End If
  
'  On Local Error GoTo ErrorCatch
'  If Exist(FileName$) Then
'    Kill FileName$
'  End If
'
'  Exit Sub
'
'ErrorCatch:
'  Select Case Err
'    Case Is <> 53
'      Call TaxMsg(900, "File deletion permission denied " + CStr(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190.")
'      MainLog ("Killfile error code is " + CStr(Err) + " . " & "Computer: '" & Net_ComputerName & "' Username: '" & Net_UserName & "'  operator: " & CStr(OperNum))
'      Terminate
'    Case 53
'      Resume ExitFillFile
'  End Select
'
'ExitFillFile:
  
End Sub

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmTaxViewPrint.ReportName = ReportFile$
   frmTaxViewPrint.Caption = Title
   frmTaxViewPrint.PgNum = PgNum
   frmTaxViewPrint.cmdAlignment.Visible = False
   If ForceSBar Then
     frmTaxViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmTaxViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmTaxViewPrint.cmdAlignment.Enabled = True
     frmTaxViewPrint.AlignRpt = AlgnRptfile$
    Else
      frmTaxViewPrint.cmdAlignment.Enabled = False
    End If
   frmTaxViewPrint.Show 1
   Unload frmLoadingRpt
'   doAlign = False
End Sub

Public Static Function Using$(ByVal fmt As String, ByVal Number As Double)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(FmtLen)
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
  RSet FmtNumber = TempNumber
  Using = FmtNumber
  
End Function

Public Function ReplaceString$(Text As String, ChangeThis As String, ToThis As String)
  Dim StrLen As Long
  Dim Cnt As Long
  Dim NewText As String
  Dim thischar$
  Dim CTChar$
  Dim TTChar$
  Dim CTLen As Integer
  Dim TTLen As Integer
  Dim BigLen As Integer
  'this function takes the incoming text and rebuilds it one
  'letter at a time until it encounters the text to change
  'at which time it replaces the text to change with the
  'new text
  StrLen = Len(Text)
  CTLen = Len(ChangeThis$)
  TTLen = Len(ToThis$)
  If CTLen > TTLen Then
    BigLen = CTLen
  ElseIf TTLen > CTLen Then
    BigLen = TTLen
  Else
    BigLen = CTLen
  End If
  
  For Cnt = 1 To StrLen 'set up loop to iterate thru entire text
    thischar = Mid$(Text, Cnt, 1) 'step thru text a letter at a time
    CTChar = Mid$(Text, Cnt, CTLen) 'starting with the current letter
    'read ahead the length of the text "change this"
    If CTChar = ChangeThis Then 'if we find the "change this" in the
    'text
      NewText = NewText + ToThis 'assign the length of CTChar to "ToThis"
      'inside the rebuilt new text
      Cnt = Cnt + BigLen - 1 'advance count to compensate for the addition of
      'CTChar
    Else
      NewText = NewText + thischar 'build new text one letter at a time
    End If
  Next
  ReplaceString$ = Trim$(NewText) 'rim out the new text
  Text = ReplaceString$ 'old text is now new text
End Function

Public Sub InsertSSNDashes(ByRef SSN As String)
  Dim ThisLen As Integer
  Dim x As Integer
  Dim NewSSN As String
  
  If InStr(1, SSN, "-") = 4 And InStr(5, SSN, "-") = 7 Then
    Exit Sub
  End If
  ThisLen = Len(SSN)
  ReDim thischar(1 To 9) As String
  For x = 1 To 9
    thischar(x) = Mid(SSN, x, 1)
    If Not IsNumeric(thischar(x)) Or thischar(x) = "" Then
      thischar(x) = " "
    End If
  Next x
  For x = 1 To 9
    NewSSN = NewSSN + thischar(x)
    If x = 3 Or x = 5 Then
      NewSSN = NewSSN + "-"
    End If
  Next x
  
  SSN = NewSSN
  
End Sub

Public Sub CreateCustNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, Cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigName$
  Dim ThisName$
  Dim Thisx As Long
  Dim SmallName$
  Dim TempName As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As CustNameIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  Dim First As Long 'Integer 8/31/09
  Dim Second As Long 'Integer 8/31/09
  Dim Third As Long 'Integer 8/31/09
  Dim Fourth As Long 'Integer 8/31/09
  Dim Fifth As Long 'Integer 8/31/09
  Dim Sixth As Long 'Integer 8/31/09
  Dim Seventh As Long 'Integer 8/31/09
  Dim Eighth As Long 'Integer 8/31/09
  Dim Ninth As Long 'Integer 8/31/09
  
  'on error goto ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).CustName = QPTrim$(CustRec.CustName)
    ThisName = QPTrim$(CustRec.CustName)
    If ThisName > BigName Then
      BigName = ThisName
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigName = BigName + "A"
  SmallName = BigName
  Nextx = 1
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 1830
    DoEvents
  End If
  
  First = ThisCnt * 0.1
  Second = ThisCnt * 0.2
  Third = ThisCnt * 0.3
  Fourth = ThisCnt * 0.4
  Fifth = ThisCnt * 0.5
  Sixth = ThisCnt * 0.6
  Seventh = ThisCnt * 0.7
  Eighth = ThisCnt * 0.8
  Ninth = ThisCnt * 0.9
  
  Do
    For x = Nextx To ThisCnt
      ThisName = TempCustIdx(x).CustName
      If ThisName < SmallName Then
        SmallName = ThisName
        Thisx = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(Thisx)
    TempCustIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do ' NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallName = BigName
    If Nextx > First Then
      First = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2280
        DoEvents
      End If
    End If
    If Nextx > Second Then
      Second = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Third Then
      Third = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Fourth Then
      Fourth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Fifth Then
      Fifth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Sixth Then
      Sixth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Seventh Then
      Seventh = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Eighth Then
      Eighth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Ninth Then
      Ninth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
  Loop
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 3810
    DoEvents
  End If
 
  KillFile "TAXNMIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateCustNameIdx", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Sub CreateSrchNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, Cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigName$
  Dim ThisName$
  Dim Thisx As Long
  Dim SmallName$
  Dim TempName As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As SrchNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As SrchNameIdxType
  Dim ThisCnt As Long 'Integer 8/31/09
  Dim NumOfIdxRecs As Long
  Dim First As Long 'Integer 8/31/09
  Dim Second As Long 'Integer 8/31/09
  Dim Third As Long 'Integer 8/31/09
  Dim Fourth As Long 'Integer 8/31/09
  Dim Fifth As Long 'Integer 8/31/09
  Dim Sixth As Long 'Integer 8/31/09
  Dim Seventh As Long 'Integer 8/31/09
  Dim Eighth As Long 'Integer 8/31/09
  Dim Ninth As Long 'Integer 8/31/09
  
  'on error goto ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.SName) <> "" Then
      Exit For
    End If
  Next x
  
  If x > NumOfCustRecs Then
    KillFile "SRCHNMIDX.DAT"
    Close CustHandle
    Exit Sub
  End If
  
  ReDim TempCustIdx(1 To NumOfCustRecs) As SrchNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).SearchName = QPTrim$(CustRec.SName)
    ThisName = QPTrim$(CustRec.SName)
    If ThisName > BigName Then
      BigName = ThisName
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigName = BigName + "A"
  SmallName = BigName
  Nextx = 1
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 840
    DoEvents
  End If
  
  First = ThisCnt * 0.1
  Second = ThisCnt * 0.2
  Third = ThisCnt * 0.3
  Fourth = ThisCnt * 0.4
  Fifth = ThisCnt * 0.5
  Sixth = ThisCnt * 0.6
  Seventh = ThisCnt * 0.7
  Eighth = ThisCnt * 0.8
  Ninth = ThisCnt * 0.9
  
  Do
    For x = Nextx To ThisCnt
      ThisName = TempCustIdx(x).SearchName
      If ThisName < SmallName Then
        SmallName = ThisName
        Thisx = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(Thisx)
    TempCustIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do ' NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallName = BigName
    If Nextx > First Then
      First = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Second Then
      Second = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Third Then
      Third = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Fourth Then
      Fourth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Fifth Then
      Fifth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Sixth Then
      Sixth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2280
        DoEvents
      End If
    End If
    If Nextx > Seventh Then
      Seventh = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Eighth Then
      Eighth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Ninth Then
      Ninth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
  Loop
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 2820
    DoEvents
  End If
  KillFile "SRCHNMIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenSrchNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    CustIdx.CustRec = CustIdx.CustRec
    CustIdx.SearchName = CustIdx.SearchName
'    If QPTrim$(CustIdx.SearchName) <> "" Then Stop
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateSrchNameIdx", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Sub GetPersRecList(ByRef PersRecs() As Long, GCustRec&, ByRef CustName As String)
  'put routine here to create temp file if adding new cust
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTaxCustRecs As Long
  Dim WhatPers&
  Dim PCnt As Long
  
  OpenTaxCustFile THandle, NumOfTaxCustRecs
  Get THandle, GCustRec, TaxCust
  Close THandle
  
  CustName = QPTrim$(TaxCust.CustName)
  ReDim PersRecs(0 To 0) As Long
  
  OpenPersPropFile PHandle, NumOfPersRecs
  
  WhatPers& = TaxCust.FirstPersRec
  If WhatPers& > 0 Then
    Do
      Get PHandle, WhatPers&, PersRec
      If PersRec.Deleted = True Then GoTo Deleted
      PCnt = PCnt + 1
      ReDim Preserve PersRecs(0 To PCnt) As Long
      PersRecs(PCnt) = WhatPers&
Deleted:
      WhatPers& = PersRec.NextRec
    Loop While WhatPers& > 0
    PersRecs(0) = PCnt
  Else
    PersRecs(0) = 0
  End If
  
  Close PHandle
  
End Sub

Public Sub DelPersAbstract(PersRecs() As Long, WhatPers%, CustRec&)
  Dim PersRec As PersonalRecType
  Dim TaxCust As TaxCustType
  Dim NumOfPersRecs As Long
  Dim Pers2Free&
  Dim NumOfPers&
  Dim PHandle As Integer
  Dim THandle As Integer
  Dim NumOfCustRecs As Long
  Dim FirstPers&
  Dim Cnt&
  Dim DidCnt As Integer
  Dim ThisPers&
  Dim NextPers&
  
  Pers2Free& = PersRecs(WhatPers)
  If Pers2Free& = 0 Then
    Call TaxMsg(900, "ERROR: There are no personal properties to delete. Attempt aborted.")
    Close
    Exit Sub
  End If
  NumOfPers& = PersRecs(0)
  
  OpenPersPropFile PHandle, NumOfPersRecs

  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, CustRec&, TaxCust
  
  FirstPers& = TaxCust.FirstPersRec
  
  'First free the Personal in question
  Get PHandle, Pers2Free&, PersRec
  PersRec.NextRec = 0
  PersRec.CustPin = 0
  PersRec.Deleted = True
  Put PHandle, Pers2Free&, PersRec
  'Personal has been marked deleted
  
  If NumOfPers& = 1 Then        'if this was the cust's only Pers
    PersRecs(0) = 0
    TaxCust.FirstPersRec = 0 'set Pers pointer to 0
    Put THandle, CustRec&, TaxCust         'store cust info
    GoTo DonePersDel            'were finished.
  End If
  
  ReDim TPersRecs(0 To NumOfPers&) As Long
  
  For Cnt& = 1 To NumOfPers&
    ThisPers& = PersRecs(Cnt&)
    If ThisPers& <> Pers2Free& Then
      DidCnt = DidCnt + 1
      TPersRecs(DidCnt) = ThisPers&
    End If
  Next
  
  For Cnt = 1 To DidCnt
    ThisPers& = TPersRecs(Cnt)
    Get PHandle, ThisPers&, PersRec
    If Cnt = 1 Then
      TaxCust.FirstPersRec = ThisPers&
      Put THandle, CustRec&, TaxCust
    End If
    If Cnt < DidCnt Then
      NextPers& = TPersRecs(Cnt + 1)
    Else
      NextPers& = 0
    End If
    PersRec.NextRec = NextPers&
    Put PHandle, ThisPers&, PersRec
  Next
  
DonePersDel:
  Close
  
End Sub

Public Sub GetRealRecList(ByRef RealRecs() As Long, GCustRec&, ByRef CustName As String)
  'put routine here to create temp file if adding new cust
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTaxCustRecs As Long
  Dim WhatPers&
  Dim RCnt As Long
  
  OpenTaxCustFile THandle, NumOfTaxCustRecs
  Get THandle, GCustRec, TaxCust
  Close THandle
  
  CustName = QPTrim$(TaxCust.CustName)
  ReDim RealRecs(0 To 0) As Long
  
  OpenRealPropFile RHandle, NumOfRealRecs
  
  WhatPers& = TaxCust.FirstPropRec
  If WhatPers& > 0 Then
    Do
      Get RHandle, WhatPers&, RealRec
      If RealRec.Deleted = True Then GoTo Deleted
      RCnt = RCnt + 1
      ReDim Preserve RealRecs(0 To RCnt) As Long
      RealRecs(RCnt) = WhatPers&
Deleted:
      WhatPers& = RealRec.NextRec
    Loop While WhatPers& > 0
    RealRecs(0) = RCnt
  Else
    RealRecs(0) = 0
  End If
  
  Close RHandle
  
End Sub

Public Sub DelRealAbstract(RealRecs() As Long, WhatPers%, CustRec&)
  Dim RealRec As PropertyRecType
  Dim TaxCust As TaxCustType
  Dim NumOfRealRecs As Long
  Dim Real2Free&
  Dim NumOfReals&
  Dim RHandle As Integer
  Dim THandle As Integer
  Dim NumOfCustRecs As Long
  Dim FirstReal&
  Dim Cnt&
  Dim DidCnt As Integer
  Dim ThisReal&
  Dim NextReal&
  
  Real2Free& = RealRecs(WhatPers)
  If Real2Free& = 0 Then
    Call TaxMsg(900, "ERROR: There are no real properties to delete. Attempt aborted.")
    Close
    Exit Sub
  End If
  NumOfReals& = RealRecs(0)
  
  OpenRealPropFile RHandle, NumOfRealRecs

  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, CustRec&, TaxCust
  
  FirstReal& = TaxCust.FirstPropRec
  
  'First free the Personal in question
  Get RHandle, Real2Free&, RealRec
  RealRec.NextRec = 0
  RealRec.CustPin = 0
  RealRec.Deleted = True
  Put RHandle, Real2Free&, RealRec
  'Personal has been marked deleted
  
  If NumOfReals& = 1 Then        'if this was the cust's only Pers
    RealRecs(0) = 0
    TaxCust.FirstPropRec = 0 'set Pers pointer to 0
    Put THandle, CustRec&, TaxCust         'store cust info
    GoTo DoneRealDel            'were finished.
  End If
  
  ReDim TRealRecs(0 To NumOfReals&) As Long
  
  For Cnt& = 1 To NumOfReals&
    ThisReal& = RealRecs(Cnt&)
    If ThisReal& <> Real2Free& Then
      DidCnt = DidCnt + 1
      TRealRecs(DidCnt) = ThisReal&
    End If
  Next
  
  For Cnt = 1 To DidCnt
    ThisReal& = TRealRecs(Cnt)
    Get RHandle, ThisReal&, RealRec
    If Cnt = 1 Then
      TaxCust.FirstPropRec = ThisReal& 'last real prop rec
      Put THandle, CustRec&, TaxCust
    End If
    If Cnt < DidCnt Then
      NextReal& = TRealRecs(Cnt + 1)
    Else
      NextReal& = 0
    End If
    RealRec.NextRec = NextReal&
    Put RHandle, ThisReal&, RealRec
  Next
  
DoneRealDel:
  Close
  
End Sub

Public Function ParseBillNum$(Text$)
  Dim BillNum$
  Dim BNumLen As Integer
  Dim thischar$
  Dim GoodPos As Integer
  Dim Cnt As Integer
  
  BillNum$ = QPTrim$(Text$)
  BNumLen = Len(BillNum$)
  If BNumLen > 0 Then
    For Cnt = BNumLen To 1 Step -1
      thischar$ = Mid$(BillNum$, Cnt, 1)
      If InStr("0123456789", thischar$) <= 0 Then
        Exit For
      End If
    Next
    GoodPos = Cnt + 1
    BillNum$ = Mid$(BillNum$, GoodPos)
  End If
  If Not IsNumeric(BillNum$) Then
    BillNum = "-911"
  End If
  ParseBillNum$ = BillNum$
End Function

Public Sub MakeRealPINFile()
  Dim RealPINS As PINSearchType
  Dim RPHandle As Integer
  Dim NumOfRealPins As Long
  Dim RealRec As PropertyRecType
  Dim PHandle As Integer
  Dim NumOfPropRecs As Long
  Dim Cnt&
  
  KillFile TaxRealPINFile
  
  OpenRealPropFile PHandle, NumOfPropRecs
  
  OpenRealPinFile RPHandle, NumOfRealPins
  
  For Cnt& = 1 To NumOfPropRecs&
    Get PHandle, Cnt&, RealRec
    RealPINS.PIN = RealRec.RealPin
    RealPINS.Cust = Cnt&
    Put RPHandle, Cnt&, RealPINS
  Next
  
  Close

End Sub

Public Sub MakePersPINFile()
  Dim PersPINS As PINSearchType
  Dim PPHandle As Integer
  Dim NumOfPersPins As Long
  Dim PersRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPropRecs As Long
  Dim Cnt&
  
  KillFile TaxPersPINFile
  
  OpenPersPropFile PPHandle, NumOfPropRecs
  
  OpenPersPinFile PRHandle, NumOfPersPins
  
  For Cnt& = 1 To NumOfPropRecs&
    Get PPHandle, Cnt&, PersRec
    PersPINS.PIN = PersRec.PropPin
    PersPINS.Cust = Cnt&
    Put PRHandle, Cnt&, PersPINS
  Next
  
  Close

End Sub

'Public Function GetCustBalance#(RecNo&, LastTrans&)
'  Dim TaxTran As TaxTransactionType
'  Dim THandle As Integer
'  Dim NumOfTRecs As Long
'  Dim PrevTransRec&
'  Dim GTOwed#
'  Dim TPaid#
'  Dim GTPaid#
'  Dim PrevTranRec&
'
'  PrevTranRec& = LastTrans
'  OpenTaxTransFile THandle, NumOfTRecs
'
'  Do While PrevTranRec& > 0
'    Get THandle, PrevTranRec&, TaxTran
'    Select Case TaxTran.TranType
'    Case 1    'bill
'      GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
'    Case 2    'payment
'      TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'      GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
'    Case 3    'release
'
'    Case 4    'interest
'      GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
'    Case 6    'collect/add cost
'      GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
'    Case 7    'adjustment
'      GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
'    Case 8    'misc cost
'      GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
'    Case Else
'    End Select
'    PrevTranRec& = TaxTran.LastTrans
'  Loop
'
'    GetCustBalance# = OldRound#(GTOwed# - GTPaid#)
'
'  Close
'
'End Function

Public Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP Len = lenRP '2/14/08
    Get RP, 1, RcptPrnFile
    RcptPrnFile.PaymDate = RcptPrnFile.PaymDate
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
      RecpDef = 1
    End If
    Close RP
  Else
    frmTaxMsg.Label1.Caption = "RECEIPT SETUP FILE NOT FOUND. Payment receipts will not be able to print. Receipt setup can be found on the Citipak main menu."
    frmTaxMsg.Label1.Top = 700
    frmTaxMsg.cmdExit.Text = "ESC OK"
    frmTaxMsg.Show vbModal
    Close
    RecpDef = 99
  End If
End Sub

Public Function BegBalCheck(CustNum As Long, ByVal ONum$, ByRef ThisRec As Integer) As Integer
  Dim OHandle As Integer
  Dim OperRec As CitiPassType 'CMOperRecType
  Dim NumOperRecs As Integer
  Dim x As Integer
  Dim Operator$
  Dim y As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As TaxPaymentRecType
  Dim NumOfPayRecs As Integer
  
  BegBalCheck = 1
  OpenCitiPassFile OHandle, NumOperRecs
  
  If NumOperRecs = 0 Then
    Close
    Exit Function
  End If
  
  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
'      OpIdx(x) = OperRec.OperatorNumber
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle
'  Operator$ = CStr(OPERNUM)
  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    Operator = QPTrim$(Operator)
    If Exist("TAXCPR" + Operator$ + ".DAT") Then
      OpenTempPayFile PayHandle, OpIdx(x) 'look thru all operator files
      NumOfPayRecs = LOF(PayHandle) / Len(EditPayRec)
      For y = 1 To NumOfPayRecs 'if you find this customer already
      'has
        Get PayHandle, y, EditPayRec
        If CustNum = EditPayRec.CustAcct Then
          If EditPayRec.LastPayRec = 0 Then GoTo SkipDeleted
          If QPTrim$(Operator$) = QPTrim$(Str(ONum)) Then
            frmTaxMsgWOpts.Label1.Caption = "An unposted transaction is in progress for this customer. Do you want to edit this transaction?"
            frmTaxMsgWOpts.Label1.Top = 900
            frmTaxMsgWOpts.cmdCont.Text = "F10 Edit"
            frmTaxMsgWOpts.cmdExit.Text = "ESC No"
            frmTaxMsgWOpts.Show vbModal
            If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
              Unload frmTaxMsgWOpts
              MainLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + CStr(CustNum) + " on " + MakeRegDate(EditPayRec.PayDate) + " and opted to continue with the payment edit.")
              BegBalCheck = 2
              ONum = "Operator"
              ThisRec = y
              Close PayHandle
            Else
              Unload frmTaxMsgWOpts
              MainLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + CStr(CustNum) + " on " + MakeRegDate(EditPayRec.PayDate) + " and opted to exit the payment edit.")
              BegBalCheck = 4
            End If
            x = NumOperRecs
            Exit For
          Else
            frmTaxMsg.Label1.Caption = "An unposted transaction is in progress by operator number " + Operator$ + " on " + MakeRegDate(EditPayRec.PayDate) + ". Edit attempt is aborted."
            frmTaxMsg.Label1.Top = 800
            frmTaxMsg.Show vbModal
            BegBalCheck = 4
            MainLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + CStr(CustNum) + " by operator #" + QPTrim$(Operator$) + " on " + MakeRegDate(EditPayRec.PayDate) + " and edit attempt was aborted.")
            Exit For
          End If
        End If
SkipDeleted:
      Next y
    End If
  Next x
  Close PayHandle
End Function

Public Function GetCustBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    GetCustBalance = 0
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenTaxTransFile THandle, NumOfTRecs
  
  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  
'  TaxYear = 2005
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      TaxTran.OperNum = TaxTran.OperNum
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
'      If TaxTran.BelongTo = 16892 And TaxTran.TranType = 2 Then Stop
'      Debug.Print CStr(PrevTranRec)
'      If PrevTranRec& = 11386 Then Stop
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        TaxTran.Revenue.Principle1Pd = TaxTran.Revenue.Principle1Pd
        TaxTran.Revenue.InterestPd = TaxTran.Revenue.InterestPd
        TaxTran.Revenue.CollectionPd = TaxTran.Revenue.CollectionPd
        TaxTran.Revenue.LateListPd = TaxTran.Revenue.LateListPd
        TaxTran.Revenue.Principle1 = TaxTran.Revenue.Principle1
        TaxTran.Revenue.Interest = TaxTran.Revenue.Interest
        TaxTran.Revenue.Collection = TaxTran.Revenue.Collection
        TaxTran.Revenue.LateList = TaxTran.Revenue.LateList
        TaxTran.BelongTo = TaxTran.BelongTo
        TaxTran.TaxYear = TaxTran.TaxYear
        TaxTran.CustomerRec = TaxTran.CustomerRec
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10    'adjust pay down affecting credit balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 11    'adjust prepay down
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 12    'refund total prepay balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    GetCustBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    GetCustBalance# = 0
  End If

  Close THandle

End Function

Public Sub TaxMsg(Top As Integer, Message As String)
  frmTaxMsg.Label1.Caption = Message
  frmTaxMsg.Label1.Top = Top
  frmTaxMsg.Show vbModal
End Sub
Public Sub Savemsg(Top As Integer, Message As String)
  frmTaxSave.Label1.Caption = Message
  frmTaxSave.Label1.Top = Top
  frmTaxSave.Show vbModal
End Sub

Public Function TaxMsgWOpts(Top As Integer, Message As String, CmdF10 As String, CmdESC As String) As String
  frmTaxMsgWOpts.Label1.Caption = Message
  frmTaxMsgWOpts.Label1.Top = Top
  frmTaxMsgWOpts.cmdCont.Text = CmdF10
  frmTaxMsgWOpts.cmdExit.Text = CmdESC
  frmTaxMsgWOpts.Show vbModal
  TaxMsgWOpts = frmTaxMsgWOpts.fptxtChoice.Text
End Function

Public Function CustHasMsg(RecNo&) As Boolean
  
  Dim MsgRec As TaxMessRecType
  Dim MsgHandle As Integer
  Dim x As Integer, y As Integer
  Dim NumMsgRec As Integer
  
  CustHasMsg = False
  OpenTaxMessage MsgHandle, NumMsgRec
'  NumMsgRec = LOF(MsgHandle) / Len(MsgRec)
  If NumMsgRec = 0 Then
    Close MsgHandle
    Exit Function
  End If
  
  If RecNo& > 0 Then
    For x = 1 To NumMsgRec
      Get MsgHandle, x, MsgRec
      If MsgRec.TaxRec = GCustNum Then
        For y = 1 To 15
          If Len(QPTrim$(MsgRec.MessLine(y).Msg)) > 0 Then
            CustHasMsg = True
            Exit For
          End If
        Next y
        Exit For
      End If
    Next x
  End If
  Close MsgHandle
End Function

Public Function RemNulls$(Text As String)
  Dim StrLen As Long
  Dim Cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For Cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, Cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, Cnt, 1) = " "
    End If
  Next
  RemNulls$ = Text
End Function

Public Function TaxMsgW3Opts(Top As Integer, Message As String, CmdF5 As String, CmdF10 As String, CmdESC As String) As String
  frmTaxMsgW3Opts.Label1.Caption = Message
  frmTaxMsgW3Opts.Label1.Top = Top
  frmTaxMsgW3Opts.cmdCont.Text = CmdF10 'continue
  frmTaxMsgW3Opts.cmdExit.Text = CmdESC 'abort
  frmTaxMsgW3Opts.cmdOption.Text = CmdF5 'option
  frmTaxMsgW3Opts.Show vbModal
  TaxMsgW3Opts = frmTaxMsgW3Opts.fptxtChoice.Text
End Function

Public Function GetOverPayBalance(RecNo&) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  GetOverPayBalance = 0
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenTaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.Revenue.PrePaidBal <> 0 Then
        GetOverPayBalance = TaxTran.Revenue.PrePaidBal
        Exit Do
      End If
      PrevTranRec& = TaxTran.LastTrans
    Loop
  End If

  Close THandle

End Function

Public Function RevsAndGLsOK(frm As Form, TaxYear As Integer) As Boolean
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim RevRec As TaxAcctsType
  Dim RRHandle As Integer
  Dim x As Integer
  Dim ThisYear As Integer
  Dim OptRev1 As Integer
  Dim OptRev2 As Integer
  Dim OptRev3 As Integer
  Dim UseIntPrinc As Boolean
  Dim UseIntOpt1 As Boolean
  Dim UseIntOpt2 As Boolean
  Dim UseIntOpt3 As Boolean
  Dim One As Integer
  Dim AHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If TaxMasterRec.AcctgMethod = "N" Then
    RevsAndGLsOK = True
    Exit Function
  End If
  
  One = 1
  AHandle = FreeFile
  Select Case frm.Name
    Case "frmTaxPrebilling"
      Open "Crevglbill.dat" For Output As AHandle
    Case "frmTaxPayMenu"
      Open "C:\CPWork\revglpay.dat" For Output As AHandle
    Case "frmTaxCalcAdCol"
      Open "C:\CPWork\revgladv.dat" For Output As AHandle
    Case "frmTaxCalcInterest"
      Open "C:\CPWork\revglint.dat" For Output As AHandle
    Case "frmTaxManualBillEntry"
      Open "C:\CPWork\revglman.dat" For Output As AHandle
  End Select
  Print #AHandle, One
  Close AHandle
  
  RevsAndGLsOK = True
  
  ThisYear = TaxYear
   
  If QPTrim$(TaxMasterRec.OptRev1) = "" Then
    OptRev1 = 0
  Else
    OptRev1 = 1
  End If
  
  If QPTrim$(TaxMasterRec.OptRev2) = "" Then
    OptRev2 = 0
  Else
    OptRev2 = 1
  End If
  
  If QPTrim$(TaxMasterRec.OptRev3) = "" Then
    OptRev3 = 0
  Else
    OptRev3 = 1
  End If
  
  If Exist("C:\CPWork\revglbill.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXGLBAC.DAT") Then
      x = 1
      GoTo NoFileBill
    End If
    OpenTaxGLInterBill RRHandle
    Get RRHandle, 1, RevRec
    Close RRHandle
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
        If QPTrim$(RevRec.TaxAcct(x).TaxCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).TaxDBAcct) = "" Then
          Exit For
        End If
      End If
    Next x
NoFileBill:
    If x < 52 Then
      RevsAndGLsOK = False
      frmTaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
      frmTaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmTaxMsgW3Opts.cmdExit.Text = "ESC Exit"
      frmTaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmTaxMsgW3Opts.Show vbModal
      If frmTaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        KillFile "C:\CPWork\revglbill.dat"
        Exit Function
      ElseIf frmTaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmTaxMsgWOpts
        frmTaxBillGLSetUp.GThisYear = ThisYear
        frmTaxBillGLSetUp.fpListYear.SearchText = frmTaxBillGLSetUp.GThisYear
        frmTaxBillGLSetUp.fpListYear.ListIndex = frmTaxBillGLSetUp.fpListYear.SearchIndex
        frmTaxBillGLSetUp.Show
        DoEvents
      Else
        Unload frmTaxMsgWOpts
        RevsAndGLsOK = True
        KillFile "C:\CPWork\revglbill.dat"
        MainLog ("ERROR: User warned that billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the billing process anyway.")
      End If
    End If
  End If
  
  If Exist("C:\CPWork\revglpay.dat") Then
    If Not Exist("TAXGLACT.DAT") Then
      x = 1
      GoTo NoFilePay
    End If
    OpenTaxGLInterPay RRHandle
    Get RRHandle, 1, RevRec
    Close RRHandle
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
        If QPTrim$(RevRec.TaxAcct(x).TaxCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).TaxDBAcct) = "" Then
          Exit For
        End If
      End If
    Next x
NoFilePay:
    If x < 52 Then
      RevsAndGLsOK = False
      frmTaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") payment requirements. This needs to be fixed before continuing the payment process. Press F5 if you would like to jump to the payment General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the payment process."
      frmTaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmTaxMsgW3Opts.cmdExit.Text = "ESC Exit"
      frmTaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmTaxMsgW3Opts.Show vbModal
      If frmTaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        KillFile "C:\CPWork\revglpay.dat"
        Exit Function
      ElseIf frmTaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmTaxMsgWOpts
        frmTaxPayGLSetup.GThisYear = ThisYear
        frmTaxPayGLSetup.fpListYear.SearchText = frmTaxPayGLSetup.GThisYear
        frmTaxPayGLSetup.fpListYear.ListIndex = frmTaxPayGLSetup.fpListYear.SearchIndex
        frmTaxPayGLSetup.Show
        DoEvents
'        Unload frm
      Else
        Unload frmTaxMsgWOpts
        KillFile "C:\CPWork\revglpay.dat"
        RevsAndGLsOK = True
        MainLog ("ERROR: User warned that pay revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the payment process anyway.")
      End If
    End If
  End If
  
  If Exist("C:\CPWork\revgladv.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXGLBAC.DAT") Then
      x = 1
      GoTo NoFileAdv
    End If
    OpenTaxGLInterBill RRHandle
    Get RRHandle, 1, RevRec
    Close RRHandle
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(RevRec.TaxAcct(x).AdvCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).AdvDBAcct) = "" Then
          Exit For
        End If
      End If
    Next x
NoFileAdv:
    If x < 52 Then
      RevsAndGLsOK = False
      frmTaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") advertising charge requirements. This needs to be fixed before continuing the advertising charges process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the advertising charges process."
      frmTaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmTaxMsgW3Opts.cmdExit.Text = "ESC Exit"
      frmTaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmTaxMsgW3Opts.Show vbModal
      If frmTaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        KillFile "C:\CPWork\revgladv.dat"
        Exit Function
      ElseIf frmTaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmTaxMsgWOpts
        frmTaxBillGLSetUp.GThisYear = ThisYear
        frmTaxBillGLSetUp.fpListYear.SearchText = frmTaxBillGLSetUp.GThisYear
        frmTaxBillGLSetUp.fpListYear.ListIndex = frmTaxBillGLSetUp.fpListYear.SearchIndex
        frmTaxBillGLSetUp.Show
        DoEvents
'        Unload frm
      Else
        Unload frmTaxMsgWOpts
        RevsAndGLsOK = True
        KillFile "C:\CPWork\revgladv.dat"
        MainLog ("ERROR: User warned that advertising charges revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the advertising charges process anyway.")
      End If
    End If
  End If
  
  If Exist("C:\CPWork\revglint.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXGLBAC.DAT") Then
      x = 1
      GoTo NoFileInt
    End If
    OpenTaxGLInterBill RRHandle
    Get RRHandle, 1, RevRec
    Close RRHandle
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(RevRec.TaxAcct(x).IntCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).IntDBAcct) = "" Then
          Exit For
        End If
      End If
    Next x
NoFileInt:
    If x < 52 Then
      RevsAndGLsOK = False
      frmTaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") interest calculations requirements. This needs to be fixed before continuing the interest calculations process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the interest calculations process."
      frmTaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmTaxMsgW3Opts.cmdExit.Text = "ESC Exit"
      frmTaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmTaxMsgW3Opts.Show vbModal
      If frmTaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        KillFile "C:\CPWork\revglint.dat"
        Exit Function
      ElseIf frmTaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmTaxMsgWOpts
        frmTaxBillGLSetUp.GThisYear = ThisYear
        frmTaxBillGLSetUp.fpListYear.SearchText = frmTaxBillGLSetUp.GThisYear
        frmTaxBillGLSetUp.fpListYear.ListIndex = frmTaxBillGLSetUp.fpListYear.SearchIndex
        frmTaxBillGLSetUp.Show
        DoEvents
'        Unload frm
      Else
        Unload frmTaxMsgWOpts
        RevsAndGLsOK = True
        KillFile "C:\CPWork\revglint.dat"
        MainLog ("ERROR: User warned that interest calculations revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the interest charges process anyway.")
      End If
    End If
  Else
    KillFile "C:\CPWork\revglint.dat"
  End If
  
  If Exist("C:\CPWork\revglman.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXGLBAC.DAT") Then
      x = 1
      GoTo NoFileMan
    End If
    OpenTaxGLInterBill RRHandle
    Get RRHandle, 1, RevRec
    Close RRHandle
    For x = 1 To 51
      If RevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(RevRec.TaxAcct(x).LtLstCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).LtLstDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(RevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
        If QPTrim$(RevRec.TaxAcct(x).TaxCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(RevRec.TaxAcct(x).TaxDBAcct) = "" Then
          Exit For
        End If
      End If
    Next x
NoFileMan:
    If x < 52 Then
      RevsAndGLsOK = False
      frmTaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
      frmTaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmTaxMsgW3Opts.cmdExit.Text = "ESC Exit"
      frmTaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmTaxMsgW3Opts.Show vbModal
      If frmTaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmTaxMsgWOpts
        KillFile "C:\CPWork\revglman.dat"
        Exit Function
      ElseIf frmTaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmTaxMsgWOpts
        frmTaxBillGLSetUp.GThisYear = ThisYear
        frmTaxBillGLSetUp.fpListYear.SearchText = frmTaxBillGLSetUp.GThisYear
        frmTaxBillGLSetUp.fpListYear.ListIndex = frmTaxBillGLSetUp.fpListYear.SearchIndex
        frmTaxBillGLSetUp.Show
        DoEvents
      Else
        Unload frmTaxMsgWOpts
        RevsAndGLsOK = True
        KillFile "C:\CPWork\revglman.dat"
        MainLog ("ERROR: User warned that manual billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the manual billing process anyway.")
      End If
    End If
  End If
  
End Function

Public Sub CheckDirs()
  Dim ThisDir$
  ThisDir = StartPath + "\TAXRPTS"

  If Not DirExists(ThisDir) Then
    frmTaxMsgWOpts.Label1.Caption = "The directory 'TAXRPTS' could not be located in the Citipak directory. Without the 'PRRPTS' directory graphics report printing is not possible. If you wish to create the 'PRRPTS' directory then press F10. Otherwise press ESC and call Southern Software @ 1-800-842-8190 for support."
    frmTaxMsgWOpts.Label1.Top = 500
    frmTaxMsgWOpts.cmdCont.Text = "F10 Make TAXRPTS"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Escape"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MkDir StartPath + "\TAXRPTS"
    Else
      Unload frmTaxMsgWOpts
    End If
  End If

  ThisDir = StartPath + "\TAXRDF"
  
  If Not DirExists(ThisDir) Then
    frmTaxMsgWOpts.Label1.Caption = "The directory 'TAXRDF' could not be located in the Citipak directory. Without the 'TAXRDF' directory graphics reports reprints are not possible. If you wish to create the 'TAXRDF' directory then press F10. Otherwise press ESC and call Southern Software @ 1-800-842-8190 for support."
    frmTaxMsgWOpts.Label1.Top = 500
    frmTaxMsgWOpts.cmdCont.Text = "F10 Make TAXRDF"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Escape"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MkDir StartPath + "\TAXRDF"
    Else
      Unload frmTaxMsgWOpts
    End If
  End If

  ThisDir = StartPath + "\TAXImages"
  
  If Not DirExists(ThisDir) Then
    If TaxMsgWOpts(700, "The directory 'TAXImages' could not be located in the Citipak directory. Press F10 if you would like to create this necessary directory now. Otherwise, press ESC to skip.", "F10 MAKE 'TAXImages'", "ESC SKIP") = "abort" Then
      Unload frmTaxMsgWOpts
    Else
      Unload frmTaxMsgWOpts
      MkDir StartPath + "\TAXImages"
    End If
  End If
 
  ThisDir = StartPath + "\TAXMortExp"
  
  If Not DirExists(ThisDir) Then
    If TaxMsgWOpts(700, "The directory 'TAXMortExp' could not be located in the Citipak directory. Press F10 if you would like to create this necessary directory now. Otherwise, press ESC to skip.", "F10 MAKE 'TAXMortExp'", "ESC SKIP") = "abort" Then
      Unload frmTaxMsgWOpts
    Else
      Unload frmTaxMsgWOpts
      MkDir StartPath + "\TAXMortExp"
    End If
  End If
  
  ThisDir = StartPath + "\TAXBILLBU"
  
  If Not DirExists(ThisDir) Then
    If TaxMsgWOpts(700, "The directory 'TAXBILLBU' could not be located in the Citipak directory. Press F10 if you would like to create this necessary directory now. Otherwise, press ESC to skip.", "F10 MAKE 'TAXBILLBU'", "ESC SKIP") = "abort" Then
      Unload frmTaxMsgWOpts
    Else
      Unload frmTaxMsgWOpts
      MkDir StartPath + "\TAXBILLBU"
    End If
  End If
  
End Sub

Public Function GetPhoneNum(PhoneNum$) As String
  Dim ThisPhone$
  Dim NewPhone$
  Dim ThisLen As Integer
  Dim x As Integer
  
  ThisPhone$ = ReplaceString(PhoneNum$, "-", "")
  ThisPhone$ = ReplaceString(ThisPhone$, "(", "")
  ThisPhone$ = ReplaceString(ThisPhone$, ")", "")
  ThisPhone$ = ReplaceString(ThisPhone$, " ", "")
  
  NewPhone = ""
  ThisLen = Len(ThisPhone)
  If ThisLen = 10 Then
    For x = 1 To 12
      If x = 4 Or x = 8 Then
        NewPhone = NewPhone + "-"
      ElseIf x < 4 Then
        NewPhone = NewPhone + Mid(ThisPhone, x, 1)
      ElseIf x < 8 And x > 4 Then
        NewPhone = NewPhone + Mid(ThisPhone, x - 1, 1)
      ElseIf x > 8 Then
        NewPhone = NewPhone + Mid(ThisPhone, x - 2, 1)
      End If
    Next x
  ElseIf ThisLen = 7 Then
    For x = 1 To 12
      If x <= 3 Then
        NewPhone = NewPhone + "0"
      ElseIf x = 4 Or x = 8 Then
        NewPhone = NewPhone + "-"
      ElseIf x <= 7 Then
        NewPhone = NewPhone + Mid(ThisPhone, x - 4, 1)
      Else
        NewPhone = NewPhone + Mid(ThisPhone, x - 5, 1)
      End If
    Next x
  End If
    
  GetPhoneNum = NewPhone

End Function

Public Function InPayBatchYN(CustRec As Long) As Boolean
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim CitiPass As CitiPassType
  Dim x As Integer, y As Integer
  Dim TaxPaymentRec As TaxPaymentRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  
  InPayBatchYN = False
  If Len(Dir$("Citipass.dat")) Then
    OpenCitiPassFile CitiPassFile, NumPassRecs
    If NumPassRecs = 0 Then
      Close CitiPassFile
      Exit Function
    End If
    ReDim OPNums(1 To NumPassRecs) As Integer
    ReDim OPNames(1 To NumPassRecs) As String
    If Not CitiPassFile = -1 Then
      For x = 1 To NumPassRecs
        Get CitiPassFile, x, CitiPass
        OPNums(x) = CitiPass.PassNum
        OPNames(x) = QPTrim$(CitiPass.UserName)
      Next x
    End If
  Else
    Exit Function
  End If
  Close CitiPassFile
  For x = 1 To NumPassRecs
    If Exist("TAXCPR" + CStr(OPNums(x)) + ".DAT") Then
      OpenTempPayFile PHandle, OPNums(x)
      NumOfPRecs = LOF(PHandle) / Len(TaxPaymentRec)
      For y = 1 To NumOfPRecs
        Get PHandle, y, TaxPaymentRec
        If TaxPaymentRec.CustAcct = CustRec Then
          InPayBatchYN = True
          Call TaxMsg(700, "This customer, " + QPTrim$(TaxPaymentRec.CustName) + ", is currently included in an unposted payment file for operator " + OPNames(x) + ". Please post this payment file before continuing with this adjustment.")
          Close PHandle
          Exit Function
        End If
      Next y
      Close PHandle
    End If
  Next x
  
End Function

Public Sub CreateOptCustIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, Cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigDesc$
  Dim ThisDesc$
  Dim Thisx As Long
  Dim SmallDesc$
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As OptCustIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As OptCustIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  Dim First As Long 'Integer 8/31/09
  Dim Second As Long 'Integer 8/31/09
  Dim Third As Long 'Integer 8/31/09
  Dim Fourth As Long 'Integer 8/31/09
  Dim Fifth As Long 'Integer 8/31/09
  Dim Sixth As Long 'Integer 8/31/09
  Dim Seventh As Long 'Integer 8/31/09
  Dim Eighth As Long 'Integer 8/31/09
  Dim Ninth As Long 'Integer 8/31/09
  Dim First1 As Long 'Integer 8/31/09
  Dim Second1 As Long 'Integer 8/31/09
  Dim Third1 As Long 'Integer 8/31/09
  Dim Fourth1 As Long 'Integer 8/31/09
  Dim Fifth1 As Long 'Integer 8/31/09
  Dim Sixth1 As Long 'Integer 8/31/09
  Dim Seventh1 As Long 'Integer 8/31/09
  Dim Eighth1 As Long 'Integer 8/31/09
  Dim Ninth1 As Long 'Integer 8/31/09
  
  'on error goto ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.OptSrchDesc) <> "" Then
      Exit For
    End If
  Next x
  
  If x > NumOfCustRecs Then
    KillFile "TXCOPTSH.DAT"
    Close CustHandle
    Exit Sub
  End If
      
  ReDim TempCustIdx(1 To NumOfCustRecs) As OptCustIdxType
  
  BigDesc = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
'    If QPTrim$(CustRec.OptSrchDesc) = "" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).OptDesc = QPTrim$(CustRec.OptSrchDesc)
    TempCustIdx(ThisCnt).CustPin = CustRec.PIN
    ThisDesc = QPTrim$(CustRec.OptSrchDesc)
    If ThisDesc > BigDesc Then
      BigDesc = ThisDesc
    End If
BadNum:
  Next x
  Close CustHandle
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 3810
    DoEvents
  End If
  
  BigDesc = BigDesc + "A"
  SmallDesc = BigDesc
  Nextx = 1
  
  First = ThisCnt * 0.1
  First1 = ThisCnt * 0.15
  Second = ThisCnt * 0.2
  Second1 = ThisCnt * 0.25
  Third = ThisCnt * 0.3
  Third1 = ThisCnt * 0.35
  Fourth = ThisCnt * 0.4
  Fourth1 = ThisCnt * 0.45
  Fifth = ThisCnt * 0.5
  Fifth1 = ThisCnt * 0.55
  Sixth = ThisCnt * 0.6
  Sixth1 = ThisCnt * 0.65
  Seventh = ThisCnt * 0.7
  Seventh1 = ThisCnt * 0.75
  Eighth = ThisCnt * 0.8
  Eighth1 = ThisCnt * 0.85
  Ninth = ThisCnt * 0.9
  Ninth1 = ThisCnt * 0.95
  
  Do
    For x = Nextx To ThisCnt
      ThisDesc = QPTrim$(TempCustIdx(x).OptDesc)
      If ThisDesc = "" Then ThisDesc = SmallDesc + "A"
'      If ThisDesc = "" Then
'        If Nextx = 1 Then Thisx = 1
'        GoTo EmptyStr 'added 3/27/06
'      End If
      If ThisDesc <= SmallDesc Then
        SmallDesc = ThisDesc
        Thisx = x
      End If
    Next x
EmptyStr:
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(Thisx)
    TempCustIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallDesc = BigDesc
    If Nextx > First Then
      First = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > First1 Then
      First1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Second Then
      Second = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Second1 Then
      Second1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Third Then
      Third = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Third1 Then
      Third1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Fourth Then
      Fourth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Fourth1 Then
      Fourth1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Fifth Then
      Fifth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Fifth1 Then
      Fifth1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Sixth Then
      Sixth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Sixth1 Then
      Sixth1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Seventh Then
      Seventh = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    If Nextx > Seventh1 Then
      Seventh1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Eighth Then
      Eighth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Eighth1 Then
      Eighth1 = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Ninth Then
      Ninth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
    DoEvents
  
'    BigNo = BigNo + 1
'    If BigNo > 40000 Then
'      Stop
'    End If
   
  Loop
  
  KillFile "TXCOPTSH.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  
  OpenCustOptSearchFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 1830
    DoEvents
  End If
  
  Close CustIdxHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateOptCustIdx", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Sub CreateOptRealIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, Cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigDesc$
  Dim ThisDesc$
  Dim Thisx As Long
  Dim SmallDesc$
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim RealIdx As OptRealIdxType
  Dim RealIdxHandle As Integer
  Dim RealIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As OptRealIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  
  'on error goto ERRORSTUFF
  
  OpenRealPropFile RRHandle, NumOfRRREcs
  
  ReDim TempRealIdx(1 To NumOfRRREcs) As OptRealIdxType
  
  BigDesc = "A"
  For x = 1 To NumOfRRREcs
    Get RRHandle, x, RealRec
    If RealRec.Deleted = -1 Then GoTo BadNum
    If QPTrim$(RealRec.OptSearch) = "" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempRealIdx(ThisCnt).RealRec = x
    TempRealIdx(ThisCnt).OptDesc = QPTrim$(RealRec.OptSearch)
    TempRealIdx(ThisCnt).RealPin = QPTrim$(RealRec.RealPin)
    ThisDesc = QPTrim$(RealRec.OptSearch)
    If ThisDesc > BigDesc Then
      BigDesc = ThisDesc
    End If
BadNum:
  Next x
  Close RRHandle
  
'  BigDesc = BigDesc + "A"
'  SmallDesc = BigDesc
  SmallDesc = ""
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisDesc = TempRealIdx(x).OptDesc
'      If ThisDesc < SmallDesc Then
      If ThisDesc > SmallDesc Then
        SmallDesc = ThisDesc
        Thisx = x
      End If
    Next x
    HoldThis = TempRealIdx(Nextx)
    TempRealIdx(Nextx) = TempRealIdx(Thisx)
    TempRealIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallDesc = "" 'BigDesc
  Loop
  
  KillFile "TXROPTSH.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenRealOptSearchFile RealIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    RealIdx = TempRealIdx(x)
    Put RealIdxHandle, x, RealIdx
  Next x
  Close RealIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateOptRealIdx", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
End Sub
Public Sub CreateOptPersIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, Cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigDesc$
  Dim ThisDesc$
  Dim Thisx As Long
  Dim SmallDesc$
  Dim PersRec As PersonalRecType
  Dim PPHandle As Integer
  Dim NumOfPPREcs As Long
  Dim PersIdx As OptPersIdxType
  Dim PersIdxHandle As Integer
  Dim PersIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As OptPersIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  
  'on error goto ERRORSTUFF
  
  OpenPersPropFile PPHandle, NumOfPPREcs
  
  ReDim TempPersIdx(1 To NumOfPPREcs) As OptPersIdxType
  
  BigDesc = "A"
  For x = 1 To NumOfPPREcs
    Get PPHandle, x, PersRec
    If PersRec.Deleted = -1 Then GoTo BadNum
    If QPTrim$(PersRec.OptSearch) = "" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempPersIdx(ThisCnt).PersRec = x
    TempPersIdx(ThisCnt).OptDesc = QPTrim$(PersRec.OptSearch)
    TempPersIdx(ThisCnt).PersPin = QPTrim$(PersRec.PropPin)
    ThisDesc = QPTrim$(PersRec.OptSearch)
    If ThisDesc > BigDesc Then
      BigDesc = ThisDesc
    End If
BadNum:
  Next x
  Close PPHandle
  
  SmallDesc = ""
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisDesc = TempPersIdx(x).OptDesc
'      If ThisDesc < SmallDesc Then
      If ThisDesc > SmallDesc Then
        SmallDesc = ThisDesc
        Thisx = x
      End If
    Next x
    HoldThis = TempPersIdx(Nextx)
    TempPersIdx(Nextx) = TempPersIdx(Thisx)
    TempPersIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallDesc = "" 'BigDesc
  Loop
  
  KillFile "TXPOPTSH.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenPersOptSearchFile PersIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    PersIdx = TempPersIdx(x)
    Put PersIdxHandle, x, PersIdx
  Next x
  Close PersIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateOptPersIdx", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
End Sub

Public Sub CreateSSIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, Cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim Thisx As Long
  Dim SmallNum As Double
  Dim TempName As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim SSIdx As SocSecIdxType
  Dim SSIdxHandle As Integer
  Dim SSIdxRecLen As Long
  Dim NumOfSSIdxRecs As Long
  Dim RecNum As Long
  Dim HoldThis As SocSecIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  Dim SSN$
  Dim First As Long 'Integer 8/31/09
  Dim Second As Long 'Integer 8/31/09
  Dim Third As Long 'Integer 8/31/09
  Dim Fourth As Long 'Integer 8/31/09
  
  'on error goto ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  
  ReDim TempSSIdx(1 To NumOfCustRecs) As SocSecIdxType
  
  BigNum = 0
  frmTaxSaveAnimation.Show
  frmTaxSaveAnimation.L1.Left = 3810
  DoEvents
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    If QPTrim$(CustRec.CSSN) = "" Then CustRec.CSSN = "111111111"
    SSN = ReplaceString(CustRec.CSSN, "-", "")
    SSN = ReplaceString(SSN, " ", "")
    SSN = QPTrim(SSN)
    If SSN = "" Then GoTo BadNum
    If Not IsNumeric(SSN) Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempSSIdx(ThisCnt).CustRec = x
    TempSSIdx(ThisCnt).SSNum = CDbl(SSN) 'CDbl(CustRec.CSSN)
    ThisNum = CDbl(SSN) ' CDbl(CustRec.CSSN)
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close CustHandle
  
  frmTaxSaveAnimation.Show
  frmTaxSaveAnimation.L1.Left = 840
  DoEvents
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  First = ThisCnt * 0.2
  Second = ThisCnt * 0.4
  Third = ThisCnt * 0.6
  Fourth = ThisCnt * 0.8
  
  Do
    For x = Nextx To ThisCnt
      ThisNum = TempSSIdx(x).SSNum
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        Thisx = x
      End If
    Next x
    HoldThis = TempSSIdx(Nextx)
    TempSSIdx(Nextx) = TempSSIdx(Thisx)
    TempSSIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
    If Nextx > First Then
      First = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 1830
        DoEvents
      End If
    End If
    If Nextx > Second Then
      Second = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 2820
        DoEvents
      End If
    End If
    If Nextx > Third Then
      Third = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 3810
        DoEvents
      End If
    End If
    If Nextx > Fourth Then
      Fourth = ThisCnt + 1
      If frmTaxSaveAnimation.Visible = True Then
        frmTaxSaveAnimation.Show
        frmTaxSaveAnimation.L1.Left = 840
        DoEvents
      End If
    End If
  Loop
  
  KillFile "TXSSIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenSocSecIdxFile SSIdxHandle, NumOfSSIdxRecs
  For x = 1 To ThisCnt
    SSIdx = TempSSIdx(x)
    Put SSIdxHandle, x, SSIdx
  Next x
  
  Close SSIdxHandle
  
  frmTaxSaveAnimation.Show
  frmTaxSaveAnimation.L1.Left = 1830
  DoEvents
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateSSIdx", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Function GetCustBalanceForYear(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    GetCustBalanceForYear = 0
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle
'  If RecNo = 1585 Then Stop
  OpenTaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.TaxYear <> TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        TaxTran.Revenue.Principle1Pd = TaxTran.Revenue.Principle1Pd
        TaxTran.Revenue.InterestPd = TaxTran.Revenue.InterestPd
        TaxTran.Revenue.CollectionPd = TaxTran.Revenue.CollectionPd
        TaxTran.Revenue.LateListPd = TaxTran.Revenue.LateListPd
        TaxTran.Revenue.PenaltyPd = TaxTran.Revenue.PenaltyPd
        TaxTran.Revenue.Future1Pd = TaxTran.Revenue.Future1Pd
        TaxTran.Revenue.Future2Pd = TaxTran.Revenue.Future2Pd
        
        TaxTran.Revenue.Principle1 = TaxTran.Revenue.Principle1
        TaxTran.Revenue.Interest = TaxTran.Revenue.Interest
        TaxTran.Revenue.Collection = TaxTran.Revenue.Collection
        TaxTran.Revenue.LateList = TaxTran.Revenue.LateList
        TaxTran.Revenue.Penalty = TaxTran.Revenue.Penalty
        TaxTran.Revenue.Future1 = TaxTran.Revenue.Future1
        TaxTran.Revenue.Future2 = TaxTran.Revenue.Future2
        TaxTran.BelongTo = TaxTran.BelongTo
        TaxTran.TaxYear = TaxTran.TaxYear
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10    'adjust pay down affecting credit balance
        TPaid# = OldRound#(TPaid# - TaxTran.Amount)
        GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      Case 11    'adjust prepay down
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 12    'refund total prepay balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    GetCustBalanceForYear# = OldRound#(GTOwed# - GTPaid#)
  Else
    GetCustBalanceForYear# = 0
  End If

  Close THandle

End Function

Public Function GetCustBalanceForRange(RecNo&, FirstTaxYear As Integer, LastTaxYear As Integer) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    GetCustBalanceForRange = 0
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenTaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.TaxYear < FirstTaxYear Or TaxTran.TaxYear > LastTaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10    'adjust pay down affecting credit balance
        TPaid# = OldRound#(TPaid# - TaxTran.Amount)
        GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      Case 11    'adjust prepay down
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 12    'refund total prepay balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    GetCustBalanceForRange# = OldRound#(GTOwed# - GTPaid#)
  Else
    GetCustBalanceForRange# = 0
  End If

  Close THandle

End Function

Public Function Check4IntMonth() As Boolean
  Dim IntDateRec As TaxInterestDateType
  Dim IDHandle As Integer
  Dim ThisMonth$
  Dim SaveMonth$
  Dim ThisDate$
  Dim SaveDate$
  Dim DateInt As Integer
  Dim ThisYear$
  Dim SaveYear$
  
  Check4IntMonth = True
  If Exist("TAXINTCK.DAT") Then
    OpenTxIntTickFile IDHandle
    Get IDHandle, 1, IntDateRec
    Close IDHandle
    SaveDate = MakeRegDate(IntDateRec.InterestDate)
    DateInt = Date2Num(Date)
    ThisDate = MakeRegDate(DateInt)
    ThisMonth = Mid(ThisDate, 1, 2)
    SaveMonth = Mid(SaveDate, 1, 2)
    ThisYear = Mid(ThisDate, 7, 4)
    SaveYear = Mid(SaveDate, 7, 4)
    If SaveYear = ThisYear And SaveMonth = ThisMonth Then
      Exit Function
    Else
      Check4IntMonth = False
      Exit Function
    End If
  End If
End Function

Public Sub CheckInt()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  If Exist(CurrCitiPath + "TAXSETUP.DAT") Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle
    If TaxMasterRec.WarnInt = "Y" Then
      If Check4IntMonth = False Then
        If Check4PayBatch = True Then
          frmTaxUnpostedPayList.Label1.Caption = "You have NOT applied interest this month but there are unposted payments outstanding that prevent calculating interest. The following operators have unposted payments."
          frmTaxUnpostedPayList.Show vbModal
          Exit Sub
        End If
        If TaxMsgWOpts(900, "You have NOT applied interest this month. Do you want to apply interest now?", "F10 YES", "ESC NO") <> "abort" Then
          Unload frmTaxMsgWOpts
          MainLog ("User warned that interest had not been applied this month and elected to jump to Interest Menu.")
          frmTaxInterestMenu.Show
          DoEvents
          Unload frmTaxMainMenu
          Exit Sub
        Else
          Unload frmTaxMsgWOpts
          MainLog ("User warned that interest had not been applied this month and elected to skip interest charges.")
        End If
      End If
    End If
  End If

End Sub

Public Sub DeActivateControls(fmx As Form, Optional OP As Boolean, Optional whole As Boolean)
  Dim Cnt As Integer, x As Control
  For Cnt = 0 To fmx.Count - 1
  Set x = fmx.Controls.Item(Cnt)
    If TypeOf x Is CommandButton Then
      x.Enabled = False
    End If
    If TypeOf x Is fpBtn Then
      x.Enabled = False
    End If
    If TypeOf x Is fpCombo Then
      x.Enabled = False
    End If
    If TypeOf x Is fpDateTime Then
      x.Enabled = False
    End If
    If TypeOf x Is fpMask Then
      x.Enabled = False
    End If
    If TypeOf x Is fpList Then  'adding thinking may help keep item in list as selected but didn't
      x.Enabled = True
    End If
    If TypeOf x Is TextBox Then
      x.Enabled = False
    End If
    If TypeOf x Is fpText Then
      x.Enabled = False
    End If
    If TypeOf x Is Menu Then
      x.Enabled = True
    End If
  Next Cnt
  If OP = True Then
    fmx.mnuOptions.Enabled = False
  End If
  EnableCloseButton fmx.hwnd, False
  'Whole as in the whole screen
  If whole = True Then
    fmx.Enabled = False
  End If
End Sub

Public Sub ActivateControls(fmx As Form, Optional OP As Boolean)
  Dim x As Control, Cnt As Integer
  For Cnt = 0 To fmx.Count - 1
  Set x = fmx.Controls.Item(Cnt)
    If TypeOf x Is CommandButton Then
      x.Enabled = True
    End If
    If TypeOf x Is fpBtn Then
      x.Enabled = True
    End If
    If TypeOf x Is fpCombo Then
      x.Enabled = True
    End If
    If TypeOf x Is fpDateTime Then
      x.Enabled = True
    End If
    If TypeOf x Is fpMask Then
      x.Enabled = True
    End If
    If TypeOf x Is fpList Then
      x.Enabled = True
    End If
    If TypeOf x Is TextBox Then
      x.Enabled = True
    End If
    If TypeOf x Is fpText Then
      x.Enabled = True
    End If

  Next Cnt
  If OP = True Then
    fmx.mnuOptions.Enabled = True
  End If
  EnableCloseButton fmx.hwnd, True
    fmx.Enabled = True
End Sub

Public Function QPStripCom$(Address$)
  Dim x As String, StrLen As Long, Cnt As Long, thischar As String
  x$ = QPTrim$(Address$)
  StrLen = Len(x$)
  For Cnt = 1 To StrLen
    thischar = Mid$(x$, Cnt, 1)
    If thischar = "," Then
      Mid$(x$, Cnt, 1) = " "
    End If
  Next

  QPStripCom$ = Trim$(x$)

End Function

Public Function InsertZipDash(Zip$) As String
  Dim ZipLen As Integer
  Dim Thisch$
  Dim x As Integer
  Dim ThisZip$
  
  If Mid(Zip$, 6, 1) = "-" Then
    InsertZipDash = Zip$
    Exit Function
  End If
  
  ZipLen = Len(QPTrim$(Zip$))
  If ZipLen <= 5 Then
    InsertZipDash = Zip$
    Exit Function
  End If
  
  For x = 1 To ZipLen
    If x = 6 Then
      Thisch = "-" + Mid(Zip, x, 1)
    Else
      Thisch = Mid(Zip, x, 1)
    End If
    If x <> 6 Then
      If Not IsNumeric(Thisch) Then
        InsertZipDash = Zip$
        Exit Function
      End If
    Else
      If Not IsNumeric(Mid(Thisch, 2, 1)) Then
        InsertZipDash = Zip$
        Exit Function
      End If
    End If
    ThisZip = ThisZip + Thisch
  Next x
  InsertZipDash = ThisZip
End Function

Public Function GetRealBalance(PIN$) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  Dim x As Long
  
  PIN = QPTrim$(PIN)
  If PIN$ = "" Or PIN$ = "0" Then
    GetRealBalance = 0
    Exit Function
  End If
  
  OpenTaxTransFile THandle, NumOfTRecs

  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTran
    If PIN = QPTrim$(TaxTran.RealPin) Then
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Revenue.PrePaidUsed + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Revenue.PrePaidUsed + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount - TaxTran.Revenue.PrePaidAmt + TaxTran.DiscAmt) 'added prepaidamt on 1/29/08
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount - TaxTran.Revenue.PrePaidAmt + TaxTran.DiscAmt) 'added prepaidamt on 1/29/08
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)  'took out prepaidamt on 5/14/08 then put it back in on 1/26/09 see fix notes
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)  'took out prepaidamt on 5/14/08 then put it back in on 1/26/09 see fix notes
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10 'adjust pay down affecting credit balance
        TPaid# = OldRound#(TPaid# - TaxTran.Amount)
        GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      Case 11 'adjust prepay down 'added 1/29/08
        TPaid# = OldRound#(TPaid# - TaxTran.Amount)
        GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case Else
      End Select
      GetRealBalance# = OldRound#(GTOwed# - Abs(GTPaid#)) 'added Abs 6/18/2010
    End If
  Next x
  If GetRealBalance < 0 Then GetRealBalance = 0 'added 2/6/09

  Close THandle

End Function

Public Sub KillWaste()
  KillFile "C:\CPWork\manualedit.dat"
  KillFile "C:\CPWork\lateltr.dat"
  KillFile "C:\CPWork\ratetbls.dat"
  KillFile "C:\CPWork\revglbill.dat"
  KillFile "C:\CPWork\revgladv.dat"
  KillFile "C:\CPWork\revglint.dat"
  KillFile "C:\CPWork\revglman.dat"
  KillFile "C:\CPWork\taxbillGL.dat"
  KillFile "C:\CPWork\taxpayGL.dat"
  KillFile "C:\CPWork\revglpay.dat"
  KillFile "C:\CPWork\manualbill.dat"
  KillFile "C:\CPWork\custinq.dat"
  KillFile "C:\CPWork\txpyment.dat"
  KillFile "C:\CPWork\editpyment.dat"
  KillFile "C:\CPWork\billlist.dat"
  KillFile "C:\CPWork\addtbl.dat"
  KillFile "C:\CPWork\detail1.dat"
  KillFile "C:\CPWork\detail2.dat"
  KillFile "C:\CPWork\detail3.dat"
  KillFile "C:\CPWork\txadjust.dat"
  KillFile "C:\CPWork\mllbls.dat"
  KillFile "C:\CPWork\custtranshist.dat"
  KillFile "C:\CPWork\realhist.dat"
  KillFile "frombillpost.dat"
  KillFile "C:\CPWork\prepayrefund.dat"
End Sub

Public Function GetPersBalance(PIN$) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  Dim x As Long
  
  PIN = QPTrim$(PIN)
  If PIN$ = "" Or PIN$ = "0" Then
    GetPersBalance = 0
    Exit Function
  End If
  
  OpenTaxTransFile THandle, NumOfTRecs

  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTran
    If PIN = QPTrim$(TaxTran.PersPin) Then
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Revenue.PrePaidUsed + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Revenue.PrePaidUsed + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10 'adjust pay down affecting credit balance
        TPaid# = OldRound#(TPaid# - TaxTran.Amount)
        GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case Else
      End Select
      GetPersBalance# = OldRound#(GTOwed# - GTPaid#)
    End If
  Next x

  Close THandle

End Function

Public Function IsCurrentOwner(RealPin$, CustPin As Long) As Boolean
  Dim PropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  
  IsCurrentOwner = False
  RealPin$ = QPTrim$(RealPin)
  If RealPin$ = "0" Or RealPin$ = "-1" Then '-1 = Mock
    IsCurrentOwner = True
    Exit Function
  End If
  
  OpenRealPropFile RHandle, NumOfRealRecs
  If NumOfRealRecs = 0 Then
    IsCurrentOwner = True
    Close RHandle
    Exit Function
  End If
  
  For x = 1 To NumOfRealRecs
    Get RHandle, x, PropRec
    If QPTrim$(PropRec.RealPin) = RealPin$ Then
      If PropRec.CustPin = CustPin Then
        IsCurrentOwner = True
        Exit For
      End If
      Exit For
    End If
  Next x
  Close RHandle
  
End Function
Public Function GetRPTName(Newrp As String)
  Dim Part As Double
  Part = Timer
  Newrp = Newrp + QPTrim(Str(CLng(Part)))
End Function

Public Function AddDashesToGLNumber(ByVal GLNum$, Fund As Integer, Dept As Integer, Detail As Integer)
  Dim NewGLNum As String
    
  GLNum$ = ReplaceString(GLNum$, "-", "")
  NewGLNum = Mid(GLNum$, 1, Fund) + "-"
  NewGLNum = NewGLNum + Mid(GLNum$, Fund + 1, Dept) + "-"
  NewGLNum = NewGLNum + Mid(GLNum$, Fund + Dept + 1, Detail)
  AddDashesToGLNumber = NewGLNum
  
End Function

Public Function Check4UnpostedBilling(ThisRec As Long) As Boolean
  Dim TaxBill As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim TaxMRec As TaxMTransactionType
  Dim TMHandle As Integer
  Dim NumOfTMRecs As Integer
  
  Check4UnpostedBilling = False
  OpenTaxBillFile TBHandle, NumOfTBRecs
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TaxBill
    If TaxBill.RealPropRecord = ThisRec Then
      Exit For
    ElseIf TaxBill.PersPropRecord = ThisRec Then
      Exit For
    End If
  Next x
  Close TBHandle
  
  If x <= NumOfTBRecs Then
    Check4UnpostedBilling = True
    Call TaxMsg(750, "This property is currently in processing for a new tax bill and cannot be deleted until bill posting completes.")
    Exit Function
  End If
    
  OpenTaxManualBillFile TMHandle, NumOfTMRecs
  For x = 1 To NumOfTMRecs
    Get TMHandle, x, TaxMRec
    If TaxMRec.RealRec = ThisRec Then
      Exit For
    ElseIf TaxMRec.PersRec = ThisRec Then
      Exit For
    End If
  Next x
  Close TMHandle
  
  If x <= NumOfTMRecs Then
    Check4UnpostedBilling = True
    Call TaxMsg(750, "This property is currently in processing for a new manual tax bill and cannot be deleted until bill posting completes.")
  End If
    
End Function

Public Function CheckTaxYear(ByRef ThisTYear As Integer) As Boolean
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear As Integer
  Dim ThisDate$
  
  CheckTaxYear = True
  ThisDate = Date2Num(Date)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  OpenTaxTransFile TRHandle, NumOfTRRecs
  TaxYear = TaxMasterRec.TaxYear
  For x = 1 To NumOfTRRecs
    Get TRHandle, x, TransRec
    If TransRec.TaxYear > TaxYear And ThisDate <= TransRec.DiscXDate Then
      ThisTYear = TransRec.TaxYear
      CheckTaxYear = False
      Exit For
    End If
  Next x
  
  Close TRHandle
  
End Function

Public Function Check4PayBatch() As Boolean
  Dim OHandle As Integer
  Dim OperRec As CitiPassType
  Dim NumOperRecs As Integer
  Dim x As Integer
  Dim Operator$
  
  Check4PayBatch = False
  OpenCitiPassFile OHandle, NumOperRecs
  
  If NumOperRecs = 0 Then
    Close
    Exit Function
  End If
  
  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle

  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    Operator = QPTrim$(Operator)
    If Exist("TAXCPR" + Operator$ + ".DAT") Then
      Check4PayBatch = True
      Exit For
    End If
  Next x

End Function

Public Function Check4CustInPayBatch(CustRec As Long, ByRef OpNum$) As Boolean
  Dim OHandle As Integer
  Dim OperRec As CitiPassType
  Dim NumOperRecs As Integer
  Dim x As Integer, y As Integer
  Dim Operator$
  Dim TaxPaymentRec As TaxPaymentRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  
  Check4CustInPayBatch = False
  OpenCitiPassFile OHandle, NumOperRecs
  
  If NumOperRecs = 0 Then
    Close
    Exit Function
  End If
  
  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle

  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    Operator = QPTrim$(Operator)
    If Exist("TAXCPR" + Operator$ + ".DAT") Then
      OpenTempPayFile PHandle, CLng(Operator)
      NumOfPRecs = LOF(PHandle) / Len(TaxPaymentRec)
      For y = 1 To NumOfPRecs
        Get PHandle, y, TaxPaymentRec
        If TaxPaymentRec.LastPayRec > 0 Then
          If TaxPaymentRec.CustAcct = CustRec Then
            Check4CustInPayBatch = True
            OpNum = Operator
            Exit For
          End If
        End If
      Next y
      Close PHandle
      If y <= NumOfPRecs Then Exit For
    End If
  Next x

End Function

Public Function FindCustInBatchFile(CustNum As Long) As String
  Dim TaxInt As Boolean
  Dim TaxAdv As Boolean
  Dim TaxBill As Boolean
  Dim IntHandle As Integer
  Dim IntRec As InterestRecType
  Dim NumOfIntRecs As Long
  Dim x As Long
  Dim AdvRec As InterestRecType
  Dim AHandle As Integer
  Dim NumOfARecs As Long
  Dim BillRec As TaxBillType
  Dim BillHandle As Integer
  Dim NumOfBillRecs As Long
  
  TaxInt = False
  TaxAdv = False
  TaxBill = False
  
  If Exist(TaxIntFile) Then TaxInt = True
  If Exist(TaxAdvFile) Then TaxAdv = True
  If Exist(TaxBillFile) Then TaxBill = True

  If TaxInt = True Then
    OpenInterestRecFile IntHandle, NumOfIntRecs
    For x = 1 To NumOfIntRecs
      Get IntHandle, x, IntRec
      If IntRec.DelFlag = True Then GoTo SkipR
      If IntRec.CustRec = CustNum Then
        FindCustInBatchFile = "1"
        Exit For
      End If
SkipR:
    Next x
    Close IntHandle
  End If

  If TaxAdv = True Then
    OpenAdvColRecFile AHandle, NumOfARecs
    For x = 1 To NumOfARecs
      Get AHandle, x, AdvRec
      If AdvRec.DelFlag = True Then GoTo SkipAdv
      If AdvRec.CustRec = CustNum Then
        FindCustInBatchFile = FindCustInBatchFile + "2"
        Exit For
      End If
SkipAdv:
    Next x
  End If
    
  If TaxBill = True Then
    OpenTaxBillFile BillHandle, NumOfBillRecs
    For x = 1 To NumOfBillRecs
      Get BillHandle, x, BillRec
      If BillRec.CustRec = CustNum Then
        If BillRec.TotalBillDue > 0 Then
          FindCustInBatchFile = FindCustInBatchFile + "3"
        End If
        Exit For
      End If
    Next x
  End If

  If FindCustInBatchFile = "" Then FindCustInBatchFile = "0"
End Function

Public Function CountReprintFiles(FileName$) As Integer
  Dim x As Integer
  Dim DirCnt As Integer
  Dim MyPath$
  Dim MyName$
  
  CountReprintFiles = 0
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    MyName = Dir
    If Len(MyName) > 4 Then
      If InStr(MyName, FileName$) > 0 Then
        DirCnt = DirCnt + 1
      End If
    End If
  Loop
  
  If DirCnt > 0 Then
    CountReprintFiles = DirCnt
  End If

End Function

Public Sub CreateCustNameIdx1(NewName As String, RecNum As Long)
  Dim x As Long
  Dim CustRecNo As Long
  Dim CustCnt As Long
  Dim ThisName$
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  Dim Inserted As Boolean
  Dim NewCnt As Long
  Dim CurrName$
  
  'on error goto ERRORSTUFF
  
  NewName = QPTrim$(NewName)
  ReDim TempCustIdx(1 To 1) As CustNameIdxType
  Inserted = False
  NewCnt = 0
  OpenNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To NumOfIdxRecs
    Get CustIdxHandle, x, CustIdx
    If CustIdx.CustRec = RecNum Then GoTo BadNum
    If QPTrim$(CustIdx.CustName) = "" And Inserted = False Then GoTo AddNew
    If NewName < CustIdx.CustName And Inserted = False Then
AddNew:
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As CustNameIdxType
      TempCustIdx(NewCnt).CustName = NewName
      TempCustIdx(NewCnt).CustRec = RecNum
      Inserted = True
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As CustNameIdxType
      TempCustIdx(NewCnt).CustName = CustIdx.CustName
      TempCustIdx(NewCnt).CustRec = CustIdx.CustRec
    Else
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As CustNameIdxType
      TempCustIdx(NewCnt).CustName = CustIdx.CustName
      TempCustIdx(NewCnt).CustRec = CustIdx.CustRec
    End If
BadNum:
  Next x
  Close CustIdxHandle
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 1830
    DoEvents
  End If
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 3810
    DoEvents
  End If
 
  KillFile "TAXNMIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To NewCnt
    CustIdx.CustName = TempCustIdx(x).CustName
    CustIdx.CustRec = TempCustIdx(x).CustRec
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateCustNameIdx1", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
End Sub

Public Sub CreateOptCustIdx1(NewName As String, RecNum As Long)
  Dim x As Long
  Dim CustIdx As OptCustIdxType
  Dim CustIdxHandle As Integer
  Dim NumOfIdxRecs As Long
  Dim Inserted As Boolean
  Dim NewCnt As Long
  Dim CurrName$
  
'  'on error goto ERRORSTUFF
  
  NewName = QPTrim$(NewName)
  ReDim TempCustIdx(1 To 1) As OptCustIdxType
  Inserted = False
  NewCnt = 0
  OpenCustOptSearchFile CustIdxHandle, NumOfIdxRecs
  
  For x = 1 To NumOfIdxRecs
    Get CustIdxHandle, x, CustIdx
    If CustIdx.CustRec = RecNum Then GoTo BadNum
    If QPTrim$(CustIdx.OptDesc) = "" And Inserted = False Then GoTo AddNew
    If NewName < CustIdx.OptDesc And Inserted = False Then
AddNew:
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As OptCustIdxType
      TempCustIdx(NewCnt).OptDesc = NewName
      TempCustIdx(NewCnt).CustRec = RecNum
      Inserted = True
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As OptCustIdxType
      TempCustIdx(NewCnt).OptDesc = CustIdx.OptDesc
      TempCustIdx(NewCnt).CustRec = CustIdx.CustRec
    Else
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As OptCustIdxType
      TempCustIdx(NewCnt).OptDesc = CustIdx.OptDesc
      TempCustIdx(NewCnt).CustRec = CustIdx.CustRec
    End If
BadNum:
  Next x
  Close CustIdxHandle
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 3810
    DoEvents
  End If
  
  KillFile "TXCOPTSH.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  
  OpenCustOptSearchFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To NewCnt
    CustIdx.OptDesc = TempCustIdx(x).OptDesc
    CustIdx.CustRec = TempCustIdx(x).CustRec
    Put CustIdxHandle, x, CustIdx
  Next x
  
  Close CustIdxHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateOptCustIdx1", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Sub CreateSrchNameIdx1(NewName As String, RecNum As Long)
  Dim x As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As SrchNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim NumOfIdxRecs As Long
  Dim Inserted As Boolean
  Dim NewCnt As Long
  Dim CurrName$
  
  'on error goto ERRORSTUFF
  
  NewName = QPTrim$(NewName)
  ReDim TempCustIdx(1 To 1) As SrchNameIdxType
  Inserted = False
  NewCnt = 0
  
  OpenSrchNameIdxFile CustIdxHandle, NumOfIdxRecs
  
  For x = 1 To NumOfIdxRecs
    Get CustIdxHandle, x, CustIdx
    If CustIdx.CustRec = RecNum Then GoTo BadNum
    If QPTrim$(CustIdx.SearchName) = "" And Inserted = False Then GoTo AddNew
    If NewName < CustIdx.SearchName And Inserted = False Then
AddNew:
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As SrchNameIdxType
      TempCustIdx(NewCnt).SearchName = NewName
      TempCustIdx(NewCnt).CustRec = RecNum
      Inserted = True
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As SrchNameIdxType
      TempCustIdx(NewCnt).SearchName = CustIdx.SearchName
      TempCustIdx(NewCnt).CustRec = CustIdx.CustRec
    Else
      NewCnt = NewCnt + 1
      ReDim Preserve TempCustIdx(1 To NewCnt) As SrchNameIdxType
      TempCustIdx(NewCnt).SearchName = CustIdx.SearchName
      TempCustIdx(NewCnt).CustRec = CustIdx.CustRec
    End If
BadNum:
  Next x
  Close CustIdxHandle
  
  If frmTaxSaveAnimation.Visible = True Then
    frmTaxSaveAnimation.Show
    frmTaxSaveAnimation.L1.Left = 840
    DoEvents
  End If
  
  KillFile "SRCHNMIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenSrchNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To NewCnt
    CustIdx.CustRec = TempCustIdx(x).CustRec
    CustIdx.SearchName = TempCustIdx(x).SearchName
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateSrchNameIdx1", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
End Sub

Public Sub CreateSSIdx1(NewSSN As String, RecNum As Long)
  Dim x As Long
  Dim SSIdx As SocSecIdxType
  Dim SSIdxHandle As Integer
  Dim SSIdxRecLen As Long
  Dim NumOfSSIdxRecs As Long
  Dim SSN$
  Dim Inserted As Boolean
  Dim NewCnt As Long
  Dim CurrName$
  
  'on error goto ERRORSTUFF
  
  ReDim TempSSIdx(1 To 1) As SocSecIdxType
  OpenSocSecIdxFile SSIdxHandle, NumOfSSIdxRecs
  NewSSN = ReplaceString(NewSSN, "-", "")
  NewSSN = ReplaceString(NewSSN, " ", "")
  NewSSN = QPTrim(NewSSN)
  
'  frmTaxSaveAnimation.Show
'  frmTaxSaveAnimation.L1.Left = 3810
  DoEvents
  For x = 1 To NumOfSSIdxRecs
    Get SSIdxHandle, x, SSIdx
    SSN = ReplaceString(CStr(SSIdx.SSNum), "-", "")
    SSN = ReplaceString(SSN, " ", "")
    SSN = QPTrim(SSN)
    If Not IsNumeric(SSN) Then GoTo BadNum
    If SSIdx.CustRec = RecNum Then GoTo BadNum
    If SSIdx.SSNum = 0 And Inserted = False Then GoTo AddNew
    If NewSSN < SSN And Inserted = False Then
AddNew:
      NewCnt = NewCnt + 1
      ReDim Preserve TempSSIdx(1 To NewCnt) As SocSecIdxType
      TempSSIdx(NewCnt).SSNum = CDbl(NewSSN)
      TempSSIdx(NewCnt).CustRec = RecNum
      Inserted = True
      NewCnt = NewCnt + 1
      ReDim Preserve TempSSIdx(1 To NewCnt) As SocSecIdxType
      TempSSIdx(NewCnt).SSNum = SSIdx.SSNum
      TempSSIdx(NewCnt).CustRec = SSIdx.CustRec
    Else
      NewCnt = NewCnt + 1
      ReDim Preserve TempSSIdx(1 To NewCnt) As SocSecIdxType
      TempSSIdx(NewCnt).SSNum = SSIdx.SSNum
      TempSSIdx(NewCnt).CustRec = SSIdx.CustRec
    End If
BadNum:
  Next x
  Close SSIdxHandle
  
 ' frmTaxSaveAnimation.Show
 ' frmTaxSaveAnimation.L1.Left = 840
  DoEvents
  
  
  KillFile "TXSSIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenSocSecIdxFile SSIdxHandle, NumOfSSIdxRecs
  For x = 1 To NewCnt
    SSIdx.SSNum = TempSSIdx(x).SSNum
    SSIdx.CustRec = TempSSIdx(x).CustRec
    Put SSIdxHandle, x, SSIdx
  Next x
  
  Close SSIdxHandle
  
'  frmTaxSaveAnimation.Show
'  frmTaxSaveAnimation.L1.Left = 1830
  DoEvents
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateSSIdx1", Erl)
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
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

