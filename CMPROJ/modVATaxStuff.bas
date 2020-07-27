Attribute VB_Name = "modVATaxStuff"
      Public Const VATaxCustFile = "TAXCUST.DAT"
      Public Const VACustNameIdxFile = "TAXNMIDX.DAT"
      Public Const VASrchNameIdxFile = "SRCHNMIDX.DAT"
      Public Const VATaxPayFileName = "TXPAYMNT.DAT"
      Public Const VACustPinFile = "TAXCPIN.DAT"
      Public Const VATaxPropFile = "TAXPROP.DAT"
      Public Const VATaxPersFile = "TAXPERS.DAT"
      Public Const VATaxMCodeFile = "TAXMORT.DAT"
      Public Const VATaxPenHandling = "TAXPEN.DAT"
      Public Const VARealOptSearch = "TXROPTSH.DAT"
      Public Const VACustOptSearch = "TXCOPTSH.DAT"
      Public Const VASocSecIdxFile = "TXSSIDX.DAT"
      Public Const VARealHistFile = "TXRLHIST.DAT"
      Public Const VATaxManualBill = "TAXMEDIT.DAT"
      Public Const VATempTaxBillAddOn = "TMPBLADD.DAT"
      Public Const VATempRealBillRecs = "C:\CPWork\TMPBLRREC.DAT"
      Public Const VATempPersBillRecs = "C:\CPWork\TMPBLPREC.DAT"
      Public Const VATaxPersPINFile = "TAXPPIN.DAT"
      Public Const VATaxRealPINFile = "TAXRPIN.DAT"
      Public Const VATaxBillOPFile = "TAXOPBL.DAT"
      Public Const VARealTaxBillInfoFile = "TAXREALBINFO.DAT"
      Public Const VATaxPPenFile = "TAXPPEN.DAT"
      Public Const VATaxRPenFile = "TAXRPEN.DAT"
      Public Const VATaxRIntFile = "TAXRINT.DAT"
      Public Const VATaxPIntFile = "TAXPINT.DAT"
      Public Const VATaxAdvFile = "TAXADV.DAT"
      Public Const TaxTownships = "TXTWNSHP.DAT"
      Public Const VATaxPreRptFile = "TAXPREBL.RPT"
      Public Const VATaxTransFile = "TAXTRANS.DAT"
    'Virginia Added--------------------------------
      Public Const RETaxCustFile = "RETXCUST.DAT"
      Public Const PPTaxCustFile = "PPTXCUST.DAT"
      Public Const REVACustPinFile = "VARETXPN.DAT"
      Public Const PPVACustPinFile = "VAPPTXPN.DAT"
      Public Const PPTaxBillFile = "TAXPBILL.DAT"
      Public Const RealTaxBillFile = "TAXRBILL.DAT"
      Public Const PPVATaxPreRptFile = "TXPPREBL.RPT"
      Public Const REVATaxPreRptFile = "TXRPREBL.RPT"
      Public Const TxRGLInterBill = "TAXRGLBAC.DAT"
      Public Const TxPGLInterBill = "TAXPGLBAC.DAT"
      Public Const TxRGLInterPay = "TAXRGLACT.DAT"
      Public Const TxPGLInterPay = "TAXPGLACT.DAT"
      Public Const TaxBillRealName = "TAXBILRLSR.DAT"
      Public Const TaxBillPersName = "TAXBILPLSR.DAT"
      Public Const TBillExpPers = "TBXPERS.DAT"
      Public Const TBillExpReal = "TBXREAL.DAT"
      Public Const TaxIntTickler = "TAXINTCK.DAT"
      Public Const PersTaxBillFile = "TAXPBILL.DAT"
      Public Const PersVATaxBillOPFile = "TAXPERSOPBL.DAT"
      Public Const RealVATaxBillOPFile = "TAXREALOPBL.DAT"
      Public Const PersVATempTaxBillAddOn = "TMPPERSBLADD.DAT"
      Public Const PersTaxBillInfoFile = "TAXPERSBINFO.DAT"
      Public Const TaxBillPostDateFile = "TXBLPSTDTE.DAT"
      Public Const PPTRARemovalFile = "PPTRARMVL.DAT"
      Public Const DMVInfoFile = "TAXDMVIF.DAT"
  Public RPayEntry As Boolean
  Public PPayEntry As Boolean

Type VAWinRTAXGLAcctRecType
  TaxYear      As Integer        'protected
  TaxDBAcct     As String * 14
  TaxCRAcct     As String * 14
  IntDBAcct     As String * 14
  IntCRAcct     As String * 14
  AdvDBAcct     As String * 14
  AdvCRAcct     As String * 14
  Fill1         As String * 1     'protected
  LtLstDBAcct   As String * 14
  LtLstCRAcct   As String * 14
  PenDBAcct     As String * 14
  PenCRAcct     As String * 14
  Opt1DBAcct    As String * 14
  Opt1CRAcct    As String * 14
  Opt2DBAcct    As String * 14
  Opt2CRAcct    As String * 14
  Opt3DBAcct    As String * 14
  Opt3CRAcct    As String * 14
End Type

Type VAWinPTAXGLAcctRecType
  TaxYear      As Integer        'protected
  PersDBAcct     As String * 14
  PersCRAcct     As String * 14
  MTDBAcct     As String * 14
  MTCRAcct     As String * 14
  MCDBAcct     As String * 14
  MCCRAcct     As String * 14
  Fill1         As String * 1     'protected
  FEDBAcct   As String * 14
  FECRAcct   As String * 14
  MHDBAcct     As String * 14
  MHCRAcct     As String * 14
  IntDBAcct     As String * 14
  IntCRAcct     As String * 14
  PenDBAcct     As String * 14
  PenCRAcct     As String * 14
  Opt1DBAcct    As String * 14
  Opt1CRAcct    As String * 14
  Opt2DBAcct    As String * 14
  Opt2CRAcct    As String * 14
  Opt3DBAcct    As String * 14
  Opt3CRAcct    As String * 14
End Type

Type VATaxRAcctsType
  TaxAcct(1 To 51) As VAWinRTAXGLAcctRecType
End Type

Type VATaxPAcctsType
  TaxAcct(1 To 51) As VAWinPTAXGLAcctRecType
End Type

Type VATaxGLPrePayType
  TaxDBAcct     As String * 14
  TaxCRAcct     As String * 14
  Filler        As String * 70
End Type

Type VAPINRecType
  PIN As Long
End Type

Type VATaxMasterType      'Master Default Information in Setup
  Name As String * 35
  Add1 As String * 35
  Add2 As String * 35
  'ADD3 As String * 35
  'Change the add3 line to break out individual city,st,zip on 013103.
  City As String * 25
  'use taxst for state in address
  'State As String * 2
  Zip As String * 10
  TaxSt As String * 2
  'TaxForm As String * 20
  'Change taxform above to 2 byte integer
  TaxForm As Integer
  'add lateform 031303
  RTaxYear As Integer
  LateForm As Integer
'  pad As String * 16  'left from taxform string of 20
'change above pad to use for following changes as of 3-28-03
'  pad     As String * 3
  WarnInt As String * 1  'Flag to Warn if interest not applied
'  DisFlag As String * 1  'set discount flag if want interest calculated
  MinBill As Double      'amount to not print bills
  'CurRate As Single
  'PastRate As Single
  'PenRate As Single
  'use the 3 rates above (12) for other stuff
 'change rcptport to pad up above - will set printer ports when sign on
  'RcptPort As Integer
  AcctgMethod As String * 1
  'add interface option 031301
  MinTxOpt As Integer '1/26/05 '1) if the taxpayer is charged nothing if
  'their tax bill is equal to or less than this amt...2) the taxpayer is charged at least this
  'amt even if they owe nothing
  TownState As String * 2 '1/26/05
  CurrRYrInt(1 To 5) As Double  '12/14/05
  CurrRYrIntInUse As Double '12/14/05
  CurrPYrInt(1 To 5) As Double  '12/14/05
  CurrPYrIntInUse As Double '12/14/05
  PastYrInt As Double '1/26/05
  PenPct As Double '1/26/05
  PenIdx As Integer
  CntrlDepYN As String * 1
  PriorYrMltRevYN As String * 1
  OverPayGLNum As String * 14
  PenPrncTaxYN As String * 1
  PenIntYN As String * 1
  PenAdvYN As String * 1
  PenLateLstYN As String * 1
  PenOpt1YN As String * 1
  PenOpt2YN As String * 1
  PenOpt3YN As String * 1
  IntPrncTaxYN As String * 1
  IntIntYN As String * 1
  IntAdvYN As String * 1
  IntLateLstYN As String * 1
  IntOpt1YN As String * 1
  IntOpt2YN As String * 1
  IntOpt3YN As String * 1
  OptRev1 As String * 20
  OptRev2 As String * 20
  OptRev3 As String * 20
  DiscRXDate As Integer      'discount amount to calc on payment screen
  DisRPct As Double
  DiscPXDate As Integer      'discount amount to calc on payment screen
  DisPPct As Double
  OptSrchCust As String * 15
  OptSrchProp As String * 15
  CountyName(1 To 5) As String * 20
  CountyNum(1 To 5) As Integer
  UseCountyYN As String * 1
  RealPersSplit As String * 1
  CycleNum(1 To 5) As Long
  CycleName(1 To 5) As String * 20
  UseCyclesYN As String * 1
  CDCashGL  As String * 14
  CDSubGL  As String * 14
  ClassName(1 To 6) As String * 15
  MultiYear As Integer
  PPTRADisc As Double
  MaxVehTaxVal As Double
  LawChngDate As Integer 'on or about 9/6/2006 the VA law changes such that delinquents
  'no longer receive PPTRA discounts
  MinVehTaxVal As Double
  PPTRAYN As String * 1
  PenPenaltyYN As String * 1
  IntPenaltyYN As String * 1
  
  '---------------------------added for 2.05
  POptRev1 As String * 20
  POptRev2 As String * 20
  POptRev3 As String * 20
  PenPersYN As String * 1
  IntPersYN As String * 1
  PersPayOrder As Integer
  PenMTYN As String * 1
  IntMTYN As String * 1
  MTPayOrder As Integer
  PenMCYN As String * 1
  IntMCYN As String * 1
  MCPayOrder As Integer
  PenFEYN As String * 1
  IntFEYN As String * 1
  FEPayOrder As Integer
  PenMHYN As String * 1
  IntMHYN As String * 1
  MHPayOrder As Integer
  PenPIntYN As String * 1
  IntPIntYN As String * 1
  PIntPayOrder As Integer
  PenPPenYN As String * 1
  IntPPenYN As String * 1
  PPenPayOrder As Integer
  PenPOpt1YN As String * 1
  IntPOpt1YN As String * 1
  POpt1PayOrder As Integer
  PenPOpt2YN As String * 1
  IntPOpt2YN As String * 1
  POpt2PayOrder As Integer
  PenPOpt3YN As String * 1
  IntPOpt3YN As String * 1
  POpt3PayOrder As Integer
  '------------------------------------------------------
  Padding As String * 72
  PTaxYear As Integer
End Type

Type VATaxInterestDateType
  RInterestDate As Integer
  PInterestDate As Integer
End Type

Type VATax1997PPRateType
   Rate As Single
End Type

Type VATaxValuesType
  Value    As Double
  OthVal   As Double
  ExmVal   As Double
End Type

Type VATaxCustType
  Acct       As Long
  OPENDATE   As Integer
'  FName      As String * 15
'  LName      As String * 25
  CustName   As String * 50
  SName      As String * 10
  HPHONE     As String * 14
  WPHONE     As String * 14
  CSSN       As String * 11
  OSSN       As String * 11
  Addr1      As String * 35
  Addr2      As String * 35
  City       As String * 20
  State      As String * 2
  Zip        As String * 10
  Active     As String * 1    'Y if Active N if Inactive
  Interest   As String * 1    'Y/N to Charge Interest
  TaxExempt  As String * 1    'Y/N to Charge Taxes Period
  Penalty    As String * 1    'Y/N to Charge Penalty
  Employer   As String * 25
  Bankrupt   As String * 1    'Y/N to Charge Collect/Adv
  TownShip   As String * 25
'032400
  LateNotice As String * 1    'Y/N Allow late notice
'*  PAD1              As String * 202
'*Change Pad1 from 202 and added PrePayment Balance field
'*Also pointer to prepay transaction
'*added these 2 fields on 3/25/03 pks
  PrePayBal    As Double
  PrePayTrans  As Long
'032900 for New Market Va to Track Who Has Been Submitted to DMV
'  DMV1999           As String * 1'remmed out on 3/2/05
'  DMV2000           As String * 1'remmed out on 3/2/05
'  DMV2001           As String * 1'remmed out on 3/2/05
'  DMV2002           As String * 1'remmed out on 3/2/05
  CountyAcctString  As String * 18    'County Account in String Format when lo
  CountyAcct    As Long        'County Account Number to Link to County Record
  LastTrans     As Long        'Pointer to last transaction
  FirstPropRec  As Long        'Pointer to first property rec
  FirstPersRec  As Long        'Pointer to first personal rec
  PIN           As Long        'Cust internal id number.
  Deleted       As Integer     'deleted flag
  FileVer       As Integer     'this is the file struct version number
  OptSrchDesc   As String * 15 '3/1/05
  ServiceAdd    As String * 35
  DrvrsLic      As String * 10
  DeliveryPt      As String * 2
  PostalRt     As String * 4
  Cycle        As Long
  CycleName    As String * 20
  County4BillNum As Long 'used as option for billing
  County4BillName As String * 20
  Pad1         As String * 190  '*remainder after additional fields
End Type

Type VAHistRecInfoType
  TranRec    As Long
  TranType   As Integer
  TranDate   As Integer
  BelongTo   As Long
  Printed    As Integer
End Type

Type VAWinRevSourceType
  Principle1    As Double                 'Va Personal Prop
  Principle2    As Double    'For Va Only     Mach/Tools
  Principle3    As Double    'For Va Only     Merch Cap
  Principle4    As Double    'For Va Only     Farm Equip
  Principle5    As Double    'For Va Only     Mobile Homes
  Interest      As Double
  Penalty       As Double
  Collection    As Double
  Future1       As Double
  Future2       As Double
  Principle1Pd  As Double
  Principle2Pd  As Double    'For Va Only
  Principle3Pd  As Double    'For Va Only
  Principle4Pd  As Double    'For Va Only
  Principle5Pd  As Double    'For Va Only
  InterestPd    As Double
  PenaltyPd     As Double
  CollectionPd  As Double
  Future1Pd     As Double
  Future2Pd     As Double
  RevOpt1       As Double
  RevOpt1Pd     As Double
  RevOpt2       As Double
  RevOpt2Pd     As Double
  RevOpt3       As Double
  RevOpt3Pd     As Double
  LateList      As Double
  LateListPd    As Double
  PrePaidAmt    As Double
  PrePaidUsed   As Double
  PrePaidBal    As Double
  pad           As String * 80
End Type

Type VATaxTransactionType
  TransDate    As Integer          'Transaction Date
  TaxYear      As Integer          'Must Contain Full 4 digit Tax Year Here
  TranType     As Integer          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
                                   '10=DwnAdjPay 11=UpAdjPay
                                   '22=PrePayment 23=Refund Prepayment added 3-25-03
  BillType     As String * 1       'R=Real P=Personal Property C=Combined (NC/
  Amount       As Double           'Total Transaction Amount
  Revenue      As VAWinRevSourceType    'See Revenue Source Type File above
  Description  As String * 30      'Description of Transaction
  Posted2GL    As String * 1       'I/F to G/L Yes or No
  CustomerRec  As Long             'Pointer Back to Customer Record
  LastTrans    As Long             'Points to Previous Trans in History
  'actually Previous pointer
  BelongTo     As Long             'Points to Record of Bill this Transaction
  DMVSubmitted As String * 1       'Y if Sent to DMV
  DMVBatch     As Integer          'Records which batch contained the DMV Tran
  Altered      As Integer          'Flag <> 0  If TR altered at any time
' Padding      As String * 123     'Allow for Future Expansion
'changed padding 123 above on 3-25-03 to allow flag to indicate
'applied prepayment on regular payment transaction
  FromPrePay   As String * 1       'Y if from Prepayment Balance
  Padding      As String * 74     '
  PersPin      As String * 20 'added for 2.05
  RealPin      As String * 20 'added for 2.05
  CustPin      As Long 'added for 2.05
  InternalPin  As Long
  DiscXDate    As Integer 'added for 2.05
  DiscAmt      As Double 'added for 2.05
  OperNum      As Integer
  PersVal      As Double
  PPTRAVal     As Double
  PPTRADisc    As Double
  CntyPara     As String * 20
  CyclPara     As String * 20
  TShpPara     As String * 25
  PPTRARmvl    As Double
  PPTRARmvlDate As Integer
End Type

Type VAInterestRecType
  CustRec            As Long                 'Acct #
  CustName           As String * 40
  TaxYear            As Integer
  Amount             As Double
  BillNumber         As String * 10
  CurYear            As Integer
'end of form
  BillRec            As Long
  DelFlag            As Integer
'  PropRec            AS LONG
  InfoTxt            As String * 30
  NewOwnerName       As String * 50
  NewOwnerAcct       As Long
  CustPin            As Long
  RealPin            As String * 20
  PersPin            As String * 20
  BillType           As String * 1
  Padding            As String * 128
End Type

Type VAPenaltyRecType
  CustRec            As Long                 'Acct #
  CustName           As String * 40
  TaxYear            As Integer
  Amount             As Double
  BillNumber         As String * 10
  CurYear            As Integer
'end of form
  BillRec            As Long
  DelFlag            As Integer
'  PropRec            AS LONG
  InfoTxt            As String * 30
  NewOwnerName       As String * 50
  NewOwnerAcct       As Long
  CustPin            As Long
  RealPin            As String * 20
  PersPin            As String * 20
  BillType           As String * 1
  Balance            As Double
  Padding            As String * 120
End Type

Type VATaxMTransactionType
  Account      As Long
  TransDate    As Integer
  TaxYear      As Integer
  Desc         As String * 30
  TaxAmount    As Double
  IntAmount    As Double
  AdColAmount  As Double
  LateList     As Double
  Penalty      As Double
  Personal     As Double
  MachTools    As Double
  MerchCap     As Double
  FarmEquip    As Double
  MobHomes     As Double
  OptRev1      As Double
  OptRev2      As Double
  OptRev3      As Double
  BillType     As String * 1   'R=REAL P=PERS C=COMB
  SName        As String * 50
  TName        As String * 50
  RealRec      As Long
  PersRec      As Long
  BillNum      As Long
  Class        As String * 1
  Deleted      As Integer
  OverPayUsed  As Double 'if credit balance is applied
  Padding      As String * 94
End Type

Type VAMortCodeRecType
  MORTCODE As String * 8
  BName    As String * 32
  Add1     As String * 32
  Add2     As String * 32
  Add3     As String * 32
  Contact  As String * 32
  PHONE    As String * 14
'Add deleted field 021003
  Deleted  As Integer
  XFileNme As String * 8
  pad      As String * 252
End Type

Type VAMortRecType
    MORTCODE As String * 8
    MortRec  As Integer
End Type

Type VAPINSearchType
  PIN   As String * 20
  Cust  As Long
End Type

Type VAFLen2
  V As String * 64
End Type

'This is Temporary File used for listing customers for selection
Type VASortCustList
  Acct    As Long
  LName   As String * 25
  FName   As String * 15
  SSN     As String * 11
  PAddr   As String * 30
  PIN     As Long
End Type

Type VASortStruct
  who As String * 14
  RecNum As Integer
End Type

Type VAPropertyRecType
    RealPin  As String * 20
    PROPDATE As Integer
    GISPOS   As String * 20
    Map      As String * 6
    BLOCK    As String * 6
    LOTNUMB  As String * 6
    LOTACRE  As String * 1
    PropSize As Double
    PROPDISC As String * 1
    LateList As String * 1
    OptRev1Chrg As Integer
    OptRev2Chrg As Integer
    OptRev3Chrg As Integer
    TownShip As String * 30
    MORTCODE As String * 8
    PROPVALU As Double
    EXMPSENI As Double
    EXMPOTHR As Double
    PROPNOT1 As String * 31
    PROPNOT2 As String * 31
    PROPNOT3 As String * 31
    Fill1    As String * 4
    CustPin  As Long
    NextRec  As Long
    LastYrPrinted As Integer
    Deleted  As Integer
    PropAddr As String * 30
    InternalPin As Long
    LienYN As String * 1
    LienDesc As String * 30
    Mock As String * 1
    Image As String * 10
    OptSearch As String * 20
    ICPDesc As String * 15
    BldgVal As Double
    Blank  As String * 100
End Type

Type VAPersonalRecType
   PropPin  As String * 20
   PROPDATE As Integer
   PersVal  As Double
   MHVALUE  As Double
   MCVALUE  As Double
   CVALUE   As Double
   MTVALUE  As Double
   EXMPSENI As Double
   EXMPOTHR As Double
   DISCOV   As String * 1
   LateList As String * 1
   DESC1    As String * 30
   DESC2    As String * 30
   DESC3    As String * 30
   Desc4    As String * 30
   Desc5    As String * 30
   CustPin        As Long
   NextRec        As Long
   LastYrPrinted  As Integer
   Deleted        As Integer
   VehTaxYear     As Integer
   DMVSubmitted   As String * 1
   InternalPin    As Long
   TaxBillYear    As Integer
   PPTRAYN        As String * 1
   ProRate        As String * 1
   ProrateVal     As Integer
   Vin            As String * 25
   MakeMod        As String * 25
   Weight         As Double
   ModYear        As Integer
   OptRev1Chrg    As Integer
   OptRev2Chrg    As Integer
   OptRev3Chrg    As Integer
   Blank          As String * 105
End Type

Type VARETaxBillType
     CustRec            As Long                 'Acct #
     CustName           As String * 40
     CustAdd1           As String * 35
     CustAdd2           As String * 35
     CustAdd3           As String * 35
     CustZip            As String * 10
     RDesc1             As String * 30
     RDesc2             As String * 30
     RealPin            As String * 20
     RealValue          As Double
     TotalValue         As Double
     ExptValue          As Double
     RealTaxDue         As Double
     BldgValue          As Double
     LateTaxDue         As Double
     TotalBillDue       As Double
     BillNumber         As Long                          'Recpt #
     TaxYear            As Integer
     BillPrinted        As Integer            '-1 = printed
     RealPropRecord     As Long
     PriorYrBalance     As Double
     RealTaxRate        As Double
     CustPin            As Long         'additional Protection for relinking
     TownShip           As String * 20
     MORTCODE           As String * 8
     LotOrAcre          As String * 1
     LASize             As String * 9
     MortRec            As Integer
     RDesc3             As String * 30
     InternalPin        As Long 'added 5/12/05
     OptRevTax1         As Double 'added 5/12/05
     OptRevTax2         As Double 'added 5/12/05
     OptRevTax3         As Double 'added 5/12/05
     OverPayAmt         As Double 'added 5/24/05
     DueDate            As Integer
     PostDate           As Integer
     TransRec           As Long
     Padding            As String * 101
End Type

Type VARETaxBillInfoType
    TaxYear  As Integer
    BillNum  As Long
    REALRATE As Double
    LATEPCT  As Double
    PRNORDER As String * 20
    CountyPara         As String * 20 'added 5/19/05
    TwnShpPara         As String * 30 'added 5/19/05
    SplitPara          As String * 30 'added 5/19/05
    CyclePara          As String * 20 'added 5/19/05
    XDate    As Integer 'added 9/20/05
    DueDate  As Integer 'added 10/20/05
End Type

Type VAPaidOwedType
   AmtOwed   As Double
   AmtPaid   As Double
End Type

Type VACustPayListType
   CustAcct     As Long
   LastPayRec  As Long
   NumPayRec   As Long
End Type

Type VATaxPaymentRecType
    OperNum  As Integer
    payDate  As Integer
    CustAcct As Long
    CustName As String * 24
    CustAddr As String * 24
    AmtOwed  As Double
    TenderTY As String * 14
    CashAmt  As Double
    ChkAmt   As Double
    ChrgAmt  As Double
    DiscAmt  As Double
    AmtRecd  As Double
    Change   As Double
    Desc     As String * 20
    PaidOwed(1 To 10)  As VAPaidOwedType
    TotOwed  As Double
    AmtPaid  As Double
    TotPaid         As Double
    LastPayRec      As Long          'Pointer to first payment list record
    NumPayRec       As Integer       'Count of payment list records
    CustPin         As Long
    PrePayAmt As Double
    BillType As String * 1
End Type
'Type VAFLen2
'    V As String * 64
'End Type
Type VARealPayListType
  BillRec       As Long      'Pointer to bill trans rec this payment is for
  BillDate      As Integer 'added for 2.05
  DiscAmt       As Double  'added for 2.05
  DiscXDate     As Integer
  Principle1    As Double
  Interest1     As Double
  Collection    As Double
  LateList      As Double
  Penalty       As Double
  OptRev1       As Double
  OptRev2       As Double
  OptRev3       As Double
  TotPaid       As Double    'amount paid to this bill rec (partial payment)
  
  CustRec       As Long      'backup pointer to cust rec
  PrevListRec   As Long      'pointer to next paylist rec
  TaxYear       As Integer
  Description   As String * 30
  TotOwed       As Double
  PrePayAmt     As Double
End Type
    
Type VAPersPayListType
  BillRec       As Long      'Pointer to bill trans rec this payment is for
  BillDate      As Integer 'added for 2.05
  DiscAmt       As Double  'added for 2.05
  DiscXDate     As Integer
  Personal      As Double
  MachTools     As Double
  MerchCap      As Double
  FarmEquip     As Double
  MobHomes      As Double
  Interest      As Double
  Penalty       As Double
  Opt1          As Double
  Opt2          As Double
  Opt3          As Double
  TotPaid       As Double    'amount paid to this bill rec (partial payment)
  
  CustRec       As Long      'backup pointer to cust rec
  PrevListRec   As Long      'pointer to next paylist rec
  TaxYear       As Integer
  Description   As String * 30
  TotOwed       As Double
  PrePayAmt     As Double
End Type

Type VAPenaltyHandlingType
  PenIdx   As Integer 'of the 6th, 7th or 8th row choose one
  PenDesc  As String * 15 'user defined penalty name
  PenPct   As Double 'penalty percentage
  PenFlat  As Double 'penalty flat rate
  UsePct   As String * 1 'apply percentage only
  UseFlat  As String * 1 'apply flat rate only
  UseBoth  As String * 1 'apply either flat or pct in conjunction with
  'UseHigh or UseLow
  UseHigh  As String * 1 'if using UseBoth then take the higher of either
  'PenPct or PenFlat
  UseLow   As String * 1 'if using UseBoth the take the lower of either
  'PenPct or PenFlat
  AppToRev1 As String * 1 'apply penalty to this revenue
  Rev1Name  As String * 15 'description of Rev1
  AppToRev2 As String * 1
  Rev2Name  As String * 15
  AppToRev3 As String * 1
  Rev3Name  As String * 15
  AppToRev4 As String * 1
  Rev4Name  As String * 15
  AppToRev5 As String * 1
  Rev5Name  As String * 15
  AppToRev6 As String * 1
  Rev6Name  As String * 15
  AppToRev7 As String * 1
  Rev7Name  As String * 15
  AppToRev8 As String * 1
  Rev8Name  As String * 15
End Type
Type VATempTaxBillAddOn
  CustRec As Long
  CustName As String * 50
  Type As String * 50
  OldAmt As Double
  NewAmt As Double
End Type
Type VAOptRevRateTablesType
  OptRevNum As Integer
  Desc As String * 20
  Type As String * 1
  StepType As String * 1 'pct or flat rate
  FromAmt(1 To 10) As Double
  ToAmt(1 To 10) As Double
  TaxFAmt(1 To 10) As Double
  TaxPAmt(1 To 10) As Double
  FlatAmt As Double
  Deleted As Boolean
  RevType As String * 1
  Comment As String * 35
  Cushion As String * 64
End Type

Type VAPenaltyRateTablesType
  Desc As String * 20
  RateType(1 To 10) As String * 1
  StepType As String * 1 'pct or flat rate
  FromAmt(1 To 10) As Double
  ToAmt(1 To 10) As Double
  TaxFAmt(1 To 10) As Double
  TaxPAmt(1 To 10) As Double
  FlatAmt As Double
  Deleted As Boolean
  Comment As String * 35
  BillType As String * 1
  Cushion As String * 63
End Type

Type VAPPTaxBillInfoType
    TaxYear  As Integer
    BillNum  As Long
    PERSRATE As Double
    MHRate As Double
    MCRate As Double
    FERate As Double
    MTRate As Double
    ProRate As String * 1
    LATEPCT  As Double
    PRNORDER As String * 20
    DueDate As Integer
    CountyPara As String * 20 'added 5/19/05
    TwnShpPara As String * 30 'added 5/19/05
    SplitPara As String * 30 'added 5/19/05
    CyclePara As String * 20 'added 5/19/05
    XDate As Integer 'added 9/20/05End Type
End Type
Type VAPPTaxBillType
     CustRec            As Long                 'Acct #
     CustName           As String * 40
     CustAdd1           As String * 35
     CustAdd2           As String * 35
     CustAdd3           As String * 35
     CustZip            As String * 10
     RDesc1             As String * 30
     RDesc2             As String * 30
     RealPin            As String * 16
     PersValue          As Double
     MHVALUE            As Double
     MCVALUE            As Double
     FEValue            As Double
     MTVALUE            As Double
     ExptValue          As Double
     PersTaxDue         As Double
     MHTaxDue           As Double
     MCTaxDue           As Double
     FETaxDue           As Double
     MTTaxDue           As Double
     LateTaxDue         As Double
     TotalBillDue       As Double
     BillNumber         As Long         'Recpt #
     TaxYear            As Integer
     BillPrinted        As Integer      '-1 = printed
     PersPropRecord     As Long
     PriorYrBalance     As Double
     PersTaxRate        As Double
     MTTaxRate          As Double
     MCTaxRate          As Double
     FETaxRate          As Double
     MHTaxRate          As Double
     CustPin            As Long         'Same as Record #
     ChillHowieFudge    As Single
     PPTRAValue         As Double
     PPTRADiscnt        As Double
     InternalPin        As Long 'added 5/12/05
     OptRevTax1         As Double 'added 5/12/05
     OptRevTax2         As Double 'added 5/12/05
     OptRevTax3         As Double 'added 5/12/05
     OverPayAmt         As Double 'added 5/24/05
     RDesc3             As String * 30
     PersPin            As String * 20
     ProRate            As String * 1 'new for VA 2.05
     PersTaxNet         As Double 'new for VA 2.05
     MultiYrVal         As Integer 'new for VA 2.05
     DueDate            As Integer
     OptRevDesc1        As String * 20
     OptRevDesc2        As String * 20
     OptRevDesc3        As String * 20
     PostDate           As Integer
     TransRec           As Long
     Padding            As String * 92
End Type

Type TaxPPTRARemovalType
  CustName As String * 50
  CustAcct As Long
  PPTRADisc As Double
  PPTRAValue As Double
  TaxAmount As Double
  TransRec As Long
  BillNum As Long
  RmvlFile As String * 50
  BillDate As Integer
  BelongTo As Long
  TxBillPostRec As Integer
End Type

Type DMVHeader
  RecordType As String * 2              'Must be H
  Batch      As String * 7
  Jury       As String * 5              'AAND for AshLand  LURG for Lunenburg Cty
  TaxYear    As String * 5              'YYYY Format
  SubDate    As String * 9              'MMDDYYYY Format
  TotalVeh   As String * 8
  TotalAmt   As String * 13
  Filler     As String * 301
End Type

Type DMVRecord
  RecordType As String * 2              'Must be D
  LineNumber As String * 21
  SSN1       As String * 10
  LastName1  As String * 36
  FirstName1 As String * 21
  Init1      As String * 2
  SSN2       As String * 10
  LastName2  As String * 36
  FirstName2 As String * 21
  Init2      As String * 2
  Addr1      As String * 33
  Addr2      As String * 33
  City       As String * 18
  State      As String * 3
  Zip        As String * 10
  Vin        As String * 19
  VehValue   As String * 8      'Whole Dollars
  PPTaxPd    As String * 9      'Assume 2 Decimals
  PPTaxReimb As String * 7      'Assume 2 Decimals
  PPTaxStart As String * 7      'YYYYMM Format
  PPTaxEnd   As String * 7      'YYYYMM Format
  Jury       As String * 5      'AAND For Ashland
  SubDate    As String * 9      'YYYYMMDD Format
  Filler     As String * 21
End Type

Type DMVInformationType
 PerRate As Single
 Batch   As Long
 JCode   As String * 4
End Type
Public Sub OpenVATaxSetUpFile(TaxSetUpHandle As Integer)
  Dim TaxSetUpLen As Integer
  Dim TaxSetUp As VATaxMasterType
  TaxSetUpLen = Len(TaxSetUp)
  TaxSetUpHandle = FreeFile
  Open TaxSetupName For Random Shared As TaxSetUpHandle Len = TaxSetUpLen
End Sub

Public Sub OpenDMVInfoFile(DMVHandle As Integer, NumOfDMVFiles As Long)
  Dim DMVLen As Integer
  Dim DMVRec As DMVInformationType
  DMVLen = Len(DMVRec)
  DMVHandle = FreeFile
  Open DMVInfoFile For Random Shared As DMVHandle Len = DMVLen
  NumOfDMVFiles = LOF(DMVHandle) / DMVLen
End Sub
Public Sub OpenPPTRARmvlFile(PPTRARmvlHandle As Integer, NumOfPPTRARmvlFiles As Long)
  Dim PPTRARmvlLen As Integer
  Dim PPTRARmvlRec As TaxPPTRARemovalType
  PPTRARmvlLen = Len(PPTRARmvlRec)
  PPTRARmvlHandle = FreeFile
  Open PPTRARemovalFile For Random Shared As PPTRARmvlHandle Len = PPTRARmvlLen
  NumOfPPTRARmvlFiles = LOF(PPTRARmvlHandle) / PPTRARmvlLen
End Sub
Public Sub OpenVARPenRecFile(RPenRecHandle As Integer, NumOfRPenRecFiles As Long)
  Dim RPenRecLen As Integer
  Dim RPenRec As VAPenaltyRecType
  RPenRecLen = Len(RPenRec)
  RPenRecHandle = FreeFile
  Open VATaxRPenFile For Random Shared As RPenRecHandle Len = RPenRecLen
  NumOfRPenRecFiles = LOF(RPenRecHandle) / RPenRecLen
End Sub
Public Sub OpenVAPPenRecFile(PPenRecHandle As Integer, NumOfPPenRecFiles As Long)
  Dim PPenRecLen As Integer
  Dim PPenRec As VAPenaltyRecType
  PPenRecLen = Len(PPenRec)
  PPenRecHandle = FreeFile
  Open VATaxPPenFile For Random Shared As PPenRecHandle Len = PPenRecLen
  NumOfPPenRecFiles = LOF(PPenRecHandle) / PPenRecLen
End Sub
Public Sub OpenVATaxBillPersAddOn(AddOnHandle As Integer)
  Dim AddOnLen As Integer
  Dim AddOnRec As VATempTaxBillAddOn
  AddOnLen = Len(AddOnRec)
  AddOnHandle = FreeFile
  Open PersVATempTaxBillAddOn For Random Shared As AddOnHandle Len = AddOnLen
End Sub
Public Sub OpenVAPersTaxBillOverPayFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As VATaxTransactionType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open PersVATaxBillOPFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub
Public Sub OpenVASocSecIdxFile(SSHandle As Integer, NumOfSSFiles As Long)
  Dim SSRecLen As Integer
  Dim SSRec As SocSecIdxType
  SSRecLen = Len(SSRec)
  SSHandle = FreeFile
  Open VASocSecIdxFile For Random Shared As SSHandle Len = SSRecLen
  NumOfSSFiles = LOF(SSHandle) / SSRecLen
End Sub
Public Sub OpenVARealOptSearchFile(ROSHandle As Integer, NumOfROSFiles As Long)
  Dim ROSRecLen As Integer
  Dim ROSRec As OptRealIdxType
  ROSRecLen = Len(ROSRec)
  ROSHandle = FreeFile
  Open VARealOptSearch For Random Shared As ROSHandle Len = ROSRecLen
  NumOfROSFiles = LOF(ROSHandle) / ROSRecLen
End Sub
Public Sub OpenVACustOptSearchFile(COSHandle As Integer, NumOfCOSFiles As Long)
  Dim COSRecLen As Integer
  Dim COSRec As OptCustIdxType
  COSRecLen = Len(COSRec)
  COSHandle = FreeFile
  Open VACustOptSearch For Random Shared As COSHandle Len = COSRecLen
  NumOfCOSFiles = LOF(COSHandle) / COSRecLen
End Sub
Public Sub OpenVAAdvColRecFile(AdvColRecHandle As Integer, NumOfAdvColRecFiles As Long)
  Dim AdvColRecLen As Integer
  Dim AdvColRec As VAInterestRecType
  AdvColRecLen = Len(AdvColRec)
  AdvColRecHandle = FreeFile
  Open VATaxAdvFile For Random Shared As AdvColRecHandle Len = AdvColRecLen
  NumOfAdvColRecFiles = LOF(AdvColRecHandle) / AdvColRecLen
End Sub
Public Sub OpenVATaxManualBillList(TaxManBillListHandle As Integer, TaxManBillListCnt As Long)
  Dim TaxManBillListLen As Integer
  Dim TaxManBillListRec As ManualTaxListType
  TaxManBillListLen = Len(TaxManBillListRec)
  TaxManBillListHandle = FreeFile
  Open VATaxManualBillList For Random Shared As TaxManBillListHandle Len = TaxManBillListLen
  TaxManBillListCnt = LOF(TaxManBillListHandle) / Len(TaxManBillListRec)
End Sub
Public Sub OpenVATaxManualBillFile(VATaxManualBillHandle As Integer, VATaxManualBillCnt As Integer)
  Dim VATaxManualBillLen As Integer
  Dim VATaxManualBillRec As VATaxMTransactionType
  VATaxManualBillLen = Len(VATaxManualBillRec)
  VATaxManualBillHandle = FreeFile
  Open VATaxManualBill For Random Shared As VATaxManualBillHandle Len = VATaxManualBillLen
  VATaxManualBillCnt = LOF(VATaxManualBillHandle) / Len(VATaxManualBillRec)
End Sub
Public Sub OpenVARealHistFile(RealHistHandle As Integer, RealHistCnt As Long)
  Dim RealHistLen As Integer
  Dim RealHistRec As RealHistoryType
  RealHistLen = Len(RealHistRec)
  RealHistHandle = FreeFile
  Open VARealHistFile For Random Shared As RealHistHandle Len = RealHistLen
  RealHistCnt = LOF(RealHistHandle) / Len(RealHistRec)
End Sub
Public Sub OpenVATaxBillAddOn(AddOnHandle As Integer)
  Dim AddOnLen As Integer
  Dim AddOnRec As VATempTaxBillAddOn
  AddOnLen = Len(AddOnRec)
  AddOnHandle = FreeFile
  Open VATempTaxBillAddOn For Random Shared As AddOnHandle Len = AddOnLen
End Sub
Public Sub OpenVARealBillInfoFile(BillInfoHandle As Integer)
  Dim BillInfoLen As Integer
  Dim BillInfoRec As VARETaxBillInfoType
  BillInfoLen = Len(BillInfoRec)
  BillInfoHandle = FreeFile
  Open VARealTaxBillInfoFile For Random Shared As BillInfoHandle Len = BillInfoLen
End Sub
Public Sub OpenRealTaxBillFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As VARETaxBillType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open RealTaxBillFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub
Public Sub OpenPersTaxBillFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As VAPPTaxBillType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open PersTaxBillFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub
Public Sub OpenVARealTempBillRecs(TempBillHandle As Integer, TempCnt As Integer)
  Dim TempBillLen As Integer
  Dim TempBillRec As TempPayList
  TempBillLen = Len(TempBillRec)
  TempBillHandle = FreeFile
  Open VATempRealBillRecs For Random Shared As TempBillHandle Len = TempBillLen
  TempCnt = LOF(TempBillHandle) / Len(TempBillRec)
End Sub
Public Sub OpenVAPersTempBillRecs(TempBillHandle As Integer, TempCnt As Integer)
  Dim TempBillLen As Integer
  Dim TempBillRec As TempPayList
  TempBillLen = Len(TempBillRec)
  TempBillHandle = FreeFile
  Open VATempPersBillRecs For Random Shared As TempBillHandle Len = TempBillLen
  TempCnt = LOF(TempBillHandle) / Len(TempBillRec)
End Sub
Public Sub OpenVATaxPenFile(TaxPenHandle As Integer)
  Dim TaxPenLen As Integer
  Dim TaxPenRec As VAPenaltyHandlingType
  TaxPenLen = Len(TaxPenRec)
  TaxPenHandle = FreeFile
  Open VATaxPenHandling For Random Shared As TaxPenHandle Len = TaxPenLen
End Sub
Public Sub OpenVAPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As VATaxPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open VATaxPayFileName For Random Shared As PayHandle Len = PayRecLen
End Sub
Public Sub OpenVAPersPinFile(PersPinHandle As Integer, NumOfPersPins As Long)
  Dim PersPinLen As Integer
  Dim PersPinRec As VAPINSearchType
  PersPinLen = Len(PersPinRec)
  PersPinHandle = FreeFile
  Open VATaxPersPINFile For Random Shared As PersPinHandle Len = PersPinLen
  NumOfPersPins = LOF(PersPinHandle) / Len(PersPinRec)
End Sub
Public Sub OpenVARealPinFile(RealPinHandle As Integer, NumOfRealPins As Long)
  Dim RealPinLen As Integer
  Dim RealPinRec As VAPINSearchType
  RealPinLen = Len(RealPinRec)
  RealPinHandle = FreeFile
  Open VATaxRealPINFile For Random Shared As RealPinHandle Len = RealPinLen
  NumOfRealPins = LOF(RealPinHandle) / Len(RealPinRec)
End Sub
Public Sub OpenVARealPropFile(RealPropHandle As Integer, NumOfRealProp As Long)
  Dim RealPropLen As Integer
  Dim RealPropRec As VAPropertyRecType
  RealPropLen = Len(RealPropRec)
  RealPropHandle = FreeFile
  Open VATaxPropFile For Random Shared As RealPropHandle Len = RealPropLen
  NumOfRealProp = LOF(RealPropHandle) / Len(RealPropRec)
End Sub
Public Sub OpenVACustPinFile(CustPinHandle As Integer, NumOfCustPins As Long)
  Dim CustPinLen As Integer
  Dim CustPinRec As VAPINRecType
  CustPinLen = Len(CustPinRec)
  CustPinHandle = FreeFile
  Open VACustPinFile For Random Shared As CustPinHandle Len = CustPinLen
  NumOfCustPins = LOF(CustPinHandle) / Len(CustPinRec)
End Sub
Public Sub OpenVAPersPropFile(PersPropHandle As Integer, NumOfPersProp As Long)
  Dim PersPropLen As Integer
  Dim PersPropRec As VAPersonalRecType
  PersPropLen = Len(PersPropRec)
  PersPropHandle = FreeFile
  Open VATaxPersFile For Random Shared As PersPropHandle Len = PersPropLen
  NumOfPersProp = LOF(PersPropHandle) / Len(PersPropRec)
End Sub
Public Sub OpenVASrchNameIdxFile(SrchNameIdxHandle As Integer, NumOfNameIdxRec As Long)
  Dim SrchNameIdxLen As Integer
  Dim SrchNameIdxRec As SrchNameIdxType
  SrchNameIdxLen = Len(SrchNameIdxRec)
  SrchNameIdxHandle = FreeFile
  Open VASrchNameIdxFile For Random Shared As SrchNameIdxHandle Len = SrchNameIdxLen
  NumOfNameIdxRec = LOF(SrchNameIdxHandle) / Len(SrchNameIdxRec)
End Sub
Public Sub OpenVANameIdxFile(NameIdxHandle As Integer, NumOfNameIdxRec As Long)
  Dim NameIdxLen As Integer
  Dim NameIdxRec As TXCustNameIdxType
  NameIdxLen = Len(NameIdxRec)
  NameIdxHandle = FreeFile
  Open VACustNameIdxFile For Random Shared As NameIdxHandle Len = NameIdxLen
  NumOfNameIdxRec = LOF(NameIdxHandle) / Len(NameIdxRec)
End Sub
Public Sub OpenVARealTaxBillOverPayFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As VATaxTransactionType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open RealVATaxBillOPFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub
Public Sub OpenVATaxPropFile(TaxPropHandle As Integer, NumOfTaxProps As Long)
  Dim TaxPropLen As Integer
  Dim TaxPropRec As VAPropertyRecType
  TaxPropLen = Len(TaxPropRec)
  TaxPropHandle = FreeFile
  Open TaxPropName For Random Shared As TaxPropHandle Len = Len(TaxPropRec)
  NumOfTaxProps = LOF(TaxPropHandle) / Len(TaxPropRec)
End Sub
      
Public Sub OpenVATaxPersFile(PersTaxHandle As Integer, NumOfPersRecs As Long)
  Dim PersTaxLen As Integer
  Dim PersTaxRec As VAPersonalRecType
  PersTaxLen = Len(PersTaxRec)
  PersTaxHandle = FreeFile
  Open PerTaxName For Random Shared As PersTaxHandle Len = PersTaxLen
  NumOfPersRecs = LOF(PersTaxHandle) / Len(PersTaxRec)
End Sub
      
Public Sub OpenVATaxCustFile(TaxCustHandle As Integer, NumOfTaxCustRec As Long)
  Dim TaxCustLen As Integer
  Dim TaxCustRec As VATaxCustType
  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open VATaxCustFile For Random Shared As TaxCustHandle Len = TaxCustLen
  NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
End Sub
Public Sub OpenVAPInterestRecFile(InterestRecHandle As Integer, NumOfIntRecFiles As Long)
  Dim InterestRecLen As Integer
  Dim InterestRec As VAInterestRecType
  InterestRecLen = Len(InterestRec)
  InterestRecHandle = FreeFile
  Open VATaxPIntFile For Random Shared As InterestRecHandle Len = InterestRecLen
  NumOfIntRecFiles = LOF(InterestRecHandle) / InterestRecLen
End Sub
Public Sub OpenVARInterestRecFile(InterestRecHandle As Integer, NumOfIntRecFiles As Long)
  Dim InterestRecLen As Integer
  Dim InterestRec As VAInterestRecType
  InterestRecLen = Len(InterestRec)
  InterestRecHandle = FreeFile
  Open VATaxRIntFile For Random Shared As InterestRecHandle Len = InterestRecLen
  NumOfIntRecFiles = LOF(InterestRecHandle) / InterestRecLen
End Sub
Public Sub OpenVATaxTransFile(TaxTransHandle As Integer, NumOfTaxTransRecs As Long)
  Dim TaxTransLen As Integer
  Dim TaxTransRate As VATaxTransactionType
  TaxTransLen = Len(TaxTransRate)
  TaxTransHandle = FreeFile
  Open VATaxTransFile For Random Shared As TaxTransHandle Len = TaxTransLen
  NumOfTaxTransRecs = LOF(TaxTransHandle) / Len(TaxTransRate)
End Sub
Public Sub OpenRTaxGLInterBill(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As VATaxRAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxRGLInterBill For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub OpenPTaxGLInterBill(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As VATaxPAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxPGLInterBill For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub OpenRTaxGLInterPay(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As VATaxRAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxRGLInterPay For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub OpenPTaxGLInterPay(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As VATaxPAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxPGLInterPay For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub OpenPersPayListFile(PayListHandle As Integer, Oper As Integer)
  Dim PayListRec As VAPersPayListType
  Dim PayListRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  Operator$ = QPTrim$(Operator$)
  PayListRecLen = Len(PayListRec)
  PayListHandle = FreeFile
  Open "CMXPLOP" + Operator$ + ".DAT" For Random Shared As PayListHandle Len = PayListRecLen
End Sub
Public Sub OpenTempPersPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As VATaxPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  Operator$ = QPTrim$(Operator)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open "CMXPCPR" + Operator$ + ".DAT" For Random Shared As PayHandle Len = PayRecLen
End Sub
Public Sub OpenTempRealPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As VATaxPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  Operator$ = QPTrim$(Operator)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open "CMXRCPR" + Operator$ + ".DAT" For Random Shared As PayHandle Len = PayRecLen
End Sub
Public Sub OpenRealPayListFile(PayListHandle As Integer, Oper As Integer)
  Dim PayListRec As VARealPayListType
  Dim PayListRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  Operator$ = QPTrim$(Operator$)
  PayListRecLen = Len(PayListRec)
  PayListHandle = FreeFile
  Open "CMXRLOP" + Operator$ + ".DAT" For Random Shared As PayListHandle Len = PayListRecLen
End Sub

Public Function VACheckTaxYear(BillType$, ByRef ThisTYear As Integer) As Boolean
  Dim x As Long
  Dim TransRec As VATaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim TaxMasterRec As VATaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear As Integer
  Dim ThisDate$
  
  VACheckTaxYear = True
  ThisDate = Date2Num(Date)
  OpenVATaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  OpenVATaxTransFile TRHandle, NumOfTRRecs
  If BillType$ = "R" Then
    TaxYear = TaxMasterRec.RTaxYear
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TransRec
      If TransRec.TaxYear > TaxYear And ThisDate <= TransRec.DiscXDate And TransRec.BillType = "R" Then
        VACheckTaxYear = False
        ThisTYear = TransRec.TaxYear
        Exit For
      End If
    Next x
  ElseIf BillType$ = "P" Then
    TaxYear = TaxMasterRec.PTaxYear
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TransRec
      If TransRec.TaxYear > TaxYear And ThisDate <= TransRec.DiscXDate And TransRec.BillType = "P" Then
        VACheckTaxYear = False
        ThisTYear = TransRec.TaxYear
        Exit For
      End If
    Next x
  End If
  
  Close TRHandle
  
End Function

Public Function VARevsAndGLsOKP(frm As Form, TaxYear As Integer, ThisType As String) As Boolean
  Dim TMHandle As Integer, RRHandle As Integer, PPHandle As Integer, x As Integer
  Dim ThisYear As Integer, OptRev1 As Integer, OptRev2 As Integer, OptRev3 As Integer
  Dim UseIntPrinc As Boolean, UseIntOpt1 As Boolean, UseIntOpt2 As Boolean
  Dim UseIntOpt3 As Boolean, One As Integer, AHandle As Integer
  Dim RevRec As VATaxRAcctsType
  Static PRevRec As VATaxPAcctsType
  Dim TaxMasterRec As VATaxMasterType

  OpenVATaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If TaxMasterRec.AcctgMethod = "N" Then
    VARevsAndGLsOKP = True
    Exit Function
  End If
  
  One = 1
  AHandle = FreeFile
  Select Case frm.Name
    Case "frmVATaxPrebilling"
      If ThisType = "R" Then
        Open "revrglbill.dat" For Output As AHandle
      ElseIf ThisType = "P" Then
        Open "revpglbill.dat" For Output As AHandle
      End If
    Case "frmCMPaySource"
      If ThisType = "R" Then
        Open "revrglpay.dat" For Output As AHandle
      ElseIf ThisType = "P" Then
        Open "revpglpay.dat" For Output As AHandle
      End If
    Case "frmVATaxCalcAdCol"
      Open "revgladv.dat" For Output As AHandle
    Case "frmVATaxCalcInterest"
      Open "revglint.dat" For Output As AHandle
    Case "frmVATaxPManualBillEntry" '12/16/05
      Open "revglman.dat" For Output As AHandle
    Case "frmVATaxManualBillEntry"
      Open "revglman.dat" For Output As AHandle
  End Select
  Print #AHandle, One
  Close AHandle
  
  VARevsAndGLsOKP = True
  
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
  
  If Exist("revrglbill.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXRGLBAC.DAT") Then
      x = 1
      GoTo NoFileBill
    End If
    OpenRTaxGLInterBill RRHandle
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
      VARevsAndGLsOKP = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revrglbill.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxBillGLSetUp.GThisYear = ThisYear
        frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
        frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
        frmVATaxBillGLSetUp.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        VARevsAndGLsOKP = True
        KillFile "revrglbill.dat"
        TXLog ("ERROR: User warned that real billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the billing process anyway.")
      End If
    End If
  End If
  
  If Exist("revpglbill.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXPGLBAC.DAT") Then
      x = 1
      GoTo NoFileBillP
    End If
    OpenPTaxGLInterBill PPHandle
    Get PPHandle, 1, PRevRec
    Close PPHandle
    For x = 1 To 51
      If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(PRevRec.TaxAcct(x).PersCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PersDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FECRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FEDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
      End If
    Next x
NoFileBillP:
    If x < 52 Then
      VARevsAndGLsOKP = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revpglbill.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxPBillGLSetUp.GThisYear = ThisYear
        frmVATaxPBillGLSetUp.fpListYear.SearchText = frmVATaxPBillGLSetUp.GThisYear
        frmVATaxPBillGLSetUp.fpListYear.ListIndex = frmVATaxPBillGLSetUp.fpListYear.SearchIndex
        frmVATaxPBillGLSetUp.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        VARevsAndGLsOKP = True
        KillFile "revpglbill.dat"
        TXLog ("ERROR: User warned that personal billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the billing process anyway.")
      End If
    End If
  End If
  
  If Exist("revrglpay.dat") Then
    If Not Exist("TAXRGLACT.DAT") Then
      x = 1
      GoTo NoFilePay
    End If
    OpenRTaxGLInterPay RRHandle
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
      VARevsAndGLsOKP = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") payment requirements. This needs to be fixed before continuing the payment process. Press F5 if you would like to jump to the payment General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the payment process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revrglpay.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxPayGLSetup.GThisYear = ThisYear
        frmVATaxPayGLSetup.fpListYear.SearchText = frmVATaxPayGLSetup.GThisYear
        frmVATaxPayGLSetup.fpListYear.ListIndex = frmVATaxPayGLSetup.fpListYear.SearchIndex
        frmVATaxPayGLSetup.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        KillFile "revrglpay.dat"
        VARevsAndGLsOKP = True
        TXLog ("ERROR: User warned that real pay revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the payment process anyway.")
      End If
    End If
  End If
  
  If Exist("revpglpay.dat") Then
    If Not Exist("TAXPGLACT.DAT") Then
      x = 1
      GoTo NoFilePayP
    End If
    OpenPTaxGLInterPay PPHandle
    Get PPHandle, 1, PRevRec
    Close PPHandle
    For x = 1 To 51
      If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(PRevRec.TaxAcct(x).PersCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PersDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FECRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FEDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
      End If
    Next x
NoFilePayP:
    If x < 52 Then
      VARevsAndGLsOKP = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") payment requirements. This needs to be fixed before continuing the payment process. Press F5 if you would like to jump to the payment General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the payment process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revpglpay.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxPPayGLSetUp.GThisYear = ThisYear
        frmVATaxPPayGLSetUp.fpListYear.SearchText = frmVATaxPPayGLSetUp.GThisYear
        frmVATaxPPayGLSetUp.fpListYear.ListIndex = frmVATaxPPayGLSetUp.fpListYear.SearchIndex
        frmVATaxPPayGLSetUp.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        KillFile "revpglpay.dat"
        VARevsAndGLsOKP = True
        TXLog ("ERROR: User warned that personal pay revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the payment process anyway.")
      End If
    End If
  End If
  
  If Exist("revgladv.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXRGLBAC.DAT") Then
      x = 1
      GoTo NoFileAdv
    End If
    OpenRTaxGLInterBill RRHandle
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
      VARevsAndGLsOKP = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") advertising charge requirements. This needs to be fixed before continuing the advertising charges process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the advertising charges process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revgladv.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxBillGLSetUp.GThisYear = ThisYear
        frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
        frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
        frmVATaxBillGLSetUp.Show
        DoEvents
'        Unload frm
      Else
        Unload frmVATaxMsgWOpts
        VARevsAndGLsOKP = True
        KillFile "revgladv.dat"
        TXLog ("ERROR: User warned that advertising charges revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the advertising charges process anyway.")
      End If
    End If
  End If
  
  If Exist("revglint.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If ThisType = "R" Then
      If Not Exist("TAXRGLBAC.DAT") Then
        x = 1
        GoTo NoFileIntR
      End If
      OpenRTaxGLInterBill RRHandle
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
NoFileIntR:
      If x < 52 Then
        VARevsAndGLsOKP = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") real interest calculations requirements. This needs to be fixed before continuing the interest calculations process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the interest calculations process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglint.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxBillGLSetUp.GThisYear = ThisYear
          frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
          frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
          frmVATaxBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOKP = True
          KillFile "revglint.dat"
          TXLog ("ERROR: User warned that real interest calculations revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the interest charges process anyway.")
        End If
      End If
    ElseIf ThisType = "P" Then
      If Not Exist("TAXPGLBAC.DAT") Then
        x = 1
        GoTo NoFileIntP
      End If
      OpenPTaxGLInterBill RRHandle
      Get RRHandle, 1, PRevRec
      Close RRHandle
      For x = 1 To 51
        If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
          If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
            Exit For
          End If
        End If
      Next x
NoFileIntP:
      If x < 52 Then
        VARevsAndGLsOKP = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") personal interest calculations requirements. This needs to be fixed before continuing the interest calculations process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the interest calculations process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglint.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxPBillGLSetUp.GThisYear = ThisYear
          frmVATaxPBillGLSetUp.fpListYear.SearchText = frmVATaxPBillGLSetUp.GThisYear
          frmVATaxPBillGLSetUp.fpListYear.ListIndex = frmVATaxPBillGLSetUp.fpListYear.SearchIndex
          frmVATaxPBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOKP = True
          KillFile "revglint.dat"
          TXLog ("ERROR: User warned that personal interest calculations revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the interest charges process anyway.")
        End If
      End If
    End If
  Else
    KillFile "revglint.dat"
  End If
  
  If Exist("revglman.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If ThisType = "R" Then
      If Not Exist("TAXRGLBAC.DAT") Then
        x = 1
        GoTo NoFileManR
      End If
      OpenRTaxGLInterBill RRHandle
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
          If QPTrim$(RevRec.TaxAcct(x).IntCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).IntDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).PenCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).PenDBAcct) = "" Then
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
NoFileManR:
      If x < 52 Then
        VARevsAndGLsOKP = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") real billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglman.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxBillGLSetUp.GThisYear = ThisYear
          frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
          frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
          frmVATaxBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOKP = True
          KillFile "revglman.dat"
          TXLog ("ERROR: User warned that real manual billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the manual billing process anyway.")
        End If
      End If
    ElseIf ThisType = "P" Then
      If Not Exist("TAXPGLBAC.DAT") Then
        x = 1
        GoTo NoFileManP
      End If
      OpenPTaxGLInterBill RRHandle
      Get RRHandle, 1, PRevRec
      Close RRHandle
      For x = 1 To 51
        If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
          If QPTrim$(PRevRec.TaxAcct(x).PersCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).PersDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MTCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MTDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MCCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MCDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).FECRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).FEDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MHCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MHDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).PenCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).PenDBAcct) = "" Then
            Exit For
          End If
          If OptRev1 = 1 Then
            If QPTrim$(PRevRec.TaxAcct(x).Opt1CRAcct) = "" Then
              Exit For
            End If
            If QPTrim$(PRevRec.TaxAcct(x).Opt1DBAcct) = "" Then
              Exit For
            End If
          End If
          If OptRev2 = 1 Then
            If QPTrim$(PRevRec.TaxAcct(x).Opt2CRAcct) = "" Then
              Exit For
            End If
            If QPTrim$(PRevRec.TaxAcct(x).Opt2DBAcct) = "" Then
              Exit For
            End If
          End If
          If OptRev3 = 1 Then
            If QPTrim$(PRevRec.TaxAcct(x).Opt3CRAcct) = "" Then
              Exit For
            End If
            If QPTrim$(PRevRec.TaxAcct(x).Opt3DBAcct) = "" Then
              Exit For
            End If
          End If
        End If
      Next x
NoFileManP:
      If x < 52 Then
        VARevsAndGLsOKP = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") personal billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglman.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxPBillGLSetUp.GThisYear = ThisYear
          frmVATaxPBillGLSetUp.fpListYear.SearchText = frmVATaxPBillGLSetUp.GThisYear
          frmVATaxPBillGLSetUp.fpListYear.ListIndex = frmVATaxPBillGLSetUp.fpListYear.SearchIndex
          frmVATaxPBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOKP = True
          KillFile "revglman.dat"
          TXLog ("ERROR: User warned that personal manual billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the manual billing process anyway.")
        End If
      End If
    End If
  End If
  
End Function

Public Function VAGetCustPersBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    VAGetCustPersBalance = 0
    Exit Function
  End If
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenVATaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.BillType <> "P" Then GoTo MoveAlong
      If TaxYear < 0 Then GoTo AllYears
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
AllYears:
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 5    'penalty
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
      Case 30    'PPTRA removal transaction
        GTOwed# = OldRound(GTOwed# + TaxTran.PPTRARmvl)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    VAGetCustPersBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    VAGetCustPersBalance# = 0
  End If

  Close THandle

End Function

Public Function VAGetCustRealBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    VAGetCustRealBalance = 0
    Exit Function
  End If
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenVATaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.BillType <> "R" Then GoTo MoveAlong
      If TaxYear < 0 Then GoTo AllYears
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
AllYears:
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 5    'penalty
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
      Case 30    'PPTRA removal transaction...never happens w/Real
        GTOwed# = OldRound(GTOwed# + TaxTran.PPTRARmvl)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    VAGetCustRealBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    VAGetCustRealBalance# = 0
  End If

  Close THandle

End Function


Public Sub VATaxMsg(Top As Integer, Message As String)
  frmVATaxMsg.Label1.Caption = Message
  frmVATaxMsg.Label1.Top = Top
  frmVATaxMsg.Show vbModal
End Sub
Public Sub VASavemsg(Top As Integer, Message As String)
  frmVATaxSave.Label1.Caption = Message
  frmVATaxSave.Label1.Top = Top
  frmVATaxSave.Show vbModal
End Sub

Public Function VATaxMsgWOpts(Top As Integer, Message As String, CmdF10 As String, CmdESC As String) As String
  frmVATaxMsgWOpts.Label1.Caption = Message
  frmVATaxMsgWOpts.Label1.Top = Top
  frmVATaxMsgWOpts.cmdCont.Text = CmdF10
  frmVATaxMsgWOpts.CmdExit.Text = CmdESC
  VATaxMsgWOpts = frmVATaxMsgWOpts.fptxtChoice.Text
End Function
Public Function Check4CustInPayBatch(CustRec As Long, ByRef OpNum$) As Boolean
  Dim OHandle As Integer
  Dim OperRec As CitiPassType
  Dim NumOperRecs As Integer
  Dim x As Integer, y As Integer
  Dim Operator$
  Dim TaxPaymentRec As VATaxPaymentRecType
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
    If Exist("TAXRCPR" + Operator$ + ".DAT") Then
      OpenTempRealPayFile PHandle, OperNum
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
    If Exist("TAXPCPR" + Operator$ + ".DAT") Then
      OpenTempPersPayFile PHandle, OperNum
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
Public Function VABegBalCheck(CustNum As Long, ByVal ONum$, ByRef ThisRec As Integer, ThisBillType$) As Integer
  Dim OHandle As Integer
  Dim OperRec As CitiPassType 'CMOperRecType
  Dim NumOperRecs As Integer
  Dim x As Integer
  Dim Operator$
  Dim y As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As VATaxPaymentRecType
  Dim NumOfPayRecs As Integer
  
  VABegBalCheck = 1
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
    If Exist("TAXRCPR" + Operator$ + ".DAT") Or Exist("TAXPCPR" + Operator$ + ".DAT") Then
      If ThisBillType = "R" Then
        OpenTempRealPayFile PayHandle, OpIdx(x) 'look thru all operator files
      ElseIf ThisBillType = "P" Then
        OpenTempPersPayFile PayHandle, OpIdx(x) 'look thru all operator files
      End If
      NumOfPayRecs = LOF(PayHandle) / Len(EditPayRec)
      For y = 1 To NumOfPayRecs 'if you find this customer already
      'has
        Get PayHandle, y, EditPayRec
        If CustNum = EditPayRec.CustAcct Then
          If EditPayRec.LastPayRec = 0 Then GoTo SkipDeleted
          If QPTrim$(Operator$) = QPTrim$(Str(ONum)) Then
            frmVATaxMsgWOpts.Label1.Caption = "An unposted transaction is in progress for this customer. Do you want to edit this transaction?"
            frmVATaxMsgWOpts.Label1.Top = 900
            frmVATaxMsgWOpts.cmdCont.Text = "F10 Edit"
            frmVATaxMsgWOpts.CmdExit.Text = "ESC No"
            frmVATaxMsgWOpts.Show vbModal
            If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
              Unload frmVATaxMsgWOpts
              TXLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + CStr(CustNum) + " on " + MakeRegDate(EditPayRec.payDate) + " and opted to continue with the payment edit.")
              VABegBalCheck = 2
              ONum = "Operator"
              ThisRec = y
              Close PayHandle
            Else
              Unload frmVATaxMsgWOpts
              TXLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + CStr(CustNum) + " on " + MakeRegDate(EditPayRec.payDate) + " and opted to exit the payment edit.")
              VABegBalCheck = 4
            End If
            x = NumOperRecs
            Exit For
          Else
            frmVATaxMsg.Label1.Caption = "An unposted transaction is in progress by operator number " + Operator$ + " on " + MakeRegDate(EditPayRec.payDate) + ". Edit attempt is aborted."
            frmVATaxMsg.Label1.Top = 800
            frmVATaxMsg.Show vbModal
            VABegBalCheck = 4
            TXLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + CStr(CustNum) + " by operator #" + QPTrim$(Operator$) + " on " + MakeRegDate(EditPayRec.payDate) + " and edit attempt was aborted.")
            Exit For
          End If
        End If
SkipDeleted:
      Next y
    End If
  Next x
  Close PayHandle
End Function
Public Function VATaxMsgW3Opts(Top As Integer, Message As String, CmdF5 As String, CmdF10 As String, CmdESC As String) As String
  frmVATaxMsgW3Opts.Label1.Caption = Message
  frmVATaxMsgW3Opts.Label1.Top = Top
  frmVATaxMsgW3Opts.cmdCont.Text = CmdF10 'continue
  frmVATaxMsgW3Opts.CmdExit.Text = CmdESC 'abort
  frmVATaxMsgW3Opts.cmdOption.Text = CmdF5 'option
  frmVATaxMsgW3Opts.Show vbModal
 VATaxMsgW3Opts = frmVATaxMsgW3Opts.fptxtChoice.Text
End Function
Public Function VAGetOverPayBalance(RecNo&, TransType$) As Double
  Dim TaxTran As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  VAGetOverPayBalance = 0
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenVATaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.BillType <> TransType And TransType <> "N" Then GoTo NextLoop
      If TaxTran.Revenue.PrePaidBal <> 0 Then
        VAGetOverPayBalance = TaxTran.Revenue.PrePaidBal
        Exit Do
      End If
NextLoop:
      PrevTranRec& = TaxTran.LastTrans
    Loop
  End If

  Close THandle

End Function
Public Function VAGetCustBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    VAGetCustBalance = 0
    Exit Function
  End If
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenVATaxTransFile THandle, NumOfTRecs
  Dim xx As String
  xx = TaxCustRec.CustName
  
  
  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
     
      
'      If TaxTran.Amount = 0.58 Then Stop
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
'      TaxTran.BelongTo = TaxTran.BelongTo
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 5    'penalty
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
      Case 30    'PPTRA removal transaction
        GTOwed# = OldRound(GTOwed# + TaxTran.PPTRARmvl)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    VAGetCustBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    VAGetCustBalance# = 0
  End If

  Close THandle

End Function
Public Function FindVACustInBatchFile(CustNum As Long, BillType$) As String
  Dim TaxRInt As Boolean
  Dim TaxPInt As Boolean
  Dim TaxRPen As Boolean
  Dim TaxPPen As Boolean
  Dim TaxAdv As Boolean
  Dim TaxRBill As Boolean
  Dim TaxPBill As Boolean
  Dim IRHandle As Integer
  Dim IntRRec As VAInterestRecType
  Dim NumOfIRRecs As Long
  Dim IPHandle As Integer
  Dim IntPRec As VAInterestRecType
  Dim NumOfIPRecs As Long
  Dim x As Long
  Dim RPenRec As VAPenaltyRecType
  Dim RPenHandle As Integer
  Dim NumOfRPenRecs As Long
  Dim PPenRec As VAPenaltyRecType
  Dim PPenHandle As Integer
  Dim NumOfPPenRecs As Long
  Dim AdvRec As VAInterestRecType
  Dim AHandle As Integer
  Dim NumOfARecs As Long
  Dim PBillRec As VAPPTaxBillType
  Dim PBillHandle As Integer
  Dim NumOfPBillRecs As Long
  Dim RBillRec As VARETaxBillType
  Dim RBillHandle As Integer
  Dim NumOfRBillRecs As Long
  
  TaxRInt = False
  TaxPInt = False
  TaxRPen = False
  TaxPPen = False
  TaxAdv = False
  TaxRBill = False
  TaxPBill = False
  
  If Exist(VATaxRIntFile) Then TaxRInt = True
  If Exist(VATaxPIntFile) Then TaxPInt = True
  If Exist(VATaxRPenFile) Then TaxRPen = True
  If Exist(VATaxPPenFile) Then TaxPPen = True
  If Exist(VATaxAdvFile) Then TaxAdv = True
  If Exist(RealTaxBillFile) Then TaxRBill = True
  If Exist(PersTaxBillFile) Then TaxPBill = True

  If BillType = "R" Then
    If TaxRInt = True Then
      OpenVARInterestRecFile IRHandle, NumOfIRRecs
      For x = 1 To NumOfIRRecs
        Get IRHandle, x, IntRRec
        If IntRRec.DelFlag = True Then GoTo SkipIR
        If IntRRec.CustRec = CustNum Then
          FindVACustInBatchFile = "1"
          Exit For
        End If
SkipIR:
      Next x
      Close IRHandle
    End If

    If TaxRPen = True Then
      OpenVARPenRecFile RPenHandle, NumOfRPenRecs
      For x = 1 To NumOfRPenRecs
        Get RPenHandle, x, RPenRec
        If RPenRec.DelFlag = True Then GoTo SkipRPen
        If RPenRec.CustRec = CustNum Then
          FindVACustInBatchFile = FindVACustInBatchFile + "2"
          Exit For
        End If
SkipRPen:
      Next x
      Close RPenHandle
    End If
    
    If TaxAdv = True Then
      OpenVAAdvColRecFile AHandle, NumOfARecs
      For x = 1 To NumOfARecs
        Get AHandle, x, AdvRec
        If AdvRec.DelFlag = True Then GoTo SkipAdv
        If AdvRec.CustRec = CustNum Then
          FindVACustInBatchFile = FindVACustInBatchFile + "3"
          Exit For
        End If
SkipAdv:
      Next x
    End If
    
    If TaxRBill = True Then
      OpenRealTaxBillFile RBillHandle, NumOfRBillRecs
      For x = 1 To NumOfRBillRecs
        Get RBillHandle, x, RBillRec
        If RBillRec.CustRec = CustNum Then
          If RBillRec.TotalBillDue > 0 Then
            FindVACustInBatchFile = FindVACustInBatchFile + "4"
          End If
          Exit For
        End If
      Next x
    End If
  End If

  If BillType = "P" Then
    If TaxPInt = True Then
      OpenVAPInterestRecFile IPHandle, NumOfIPRecs
      For x = 1 To NumOfIPRecs
        Get IPHandle, x, IntPRec
        If IntPRec.DelFlag = True Then GoTo SkipIP
        If IntPRec.CustRec = CustNum Then
          FindVACustInBatchFile = FindVACustInBatchFile + "5"
          Exit For
        End If
SkipIP:
      Next x
      Close IPHandle
    End If
    
    If TaxPPen = True Then
      OpenVAPPenRecFile PPenHandle, NumOfPPenRecs
      For x = 1 To NumOfPPenRecs
        Get PPenHandle, x, PPenRec
        If PPenRec.DelFlag = True Then GoTo SkipPPen
        If PPenRec.CustRec = CustNum Then
          FindVACustInBatchFile = FindVACustInBatchFile + "6"
          Exit For
        End If
SkipPPen:
      Next x
      Close PPenHandle
    End If
    
    If TaxPBill = True Then
      OpenPersTaxBillFile PBillHandle, NumOfPBillRecs
      For x = 1 To NumOfPBillRecs
        Get PBillHandle, x, PBillRec
        If PBillRec.CustRec = CustNum Then
          If PBillRec.TotalBillDue > 0 Then
            FindVACustInBatchFile = FindVACustInBatchFile + "7"
          End If
          Exit For
        End If
      Next x
    End If
  End If
  
  If FindVACustInBatchFile = "" Then FindVACustInBatchFile = "0"
End Function
Public Function VARevsAndGLsOK(frm As Form, TaxYear As Integer, ThisType As String) As Boolean
  Dim TMHandle As Integer, RRHandle As Integer, PPHandle As Integer, x As Integer
  Dim ThisYear As Integer, OptRev1 As Integer, OptRev2 As Integer, OptRev3 As Integer
  Dim UseIntPrinc As Boolean, UseIntOpt1 As Boolean, UseIntOpt2 As Boolean
  Dim UseIntOpt3 As Boolean, One As Integer, AHandle As Integer
  Dim RevRec As VATaxRAcctsType
  Static PRevRec As VATaxPAcctsType
  Static TaxMasterRec As VATaxMasterType

  OpenVATaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If TaxMasterRec.AcctgMethod = "N" Then
    VARevsAndGLsOK = True
    Exit Function
  End If
  
  One = 1
  AHandle = FreeFile
  Select Case frm.Name
    Case "frmCMPaySource"
      If ThisType = "R" Then
        Open "revrglpay.dat" For Output As AHandle
      ElseIf ThisType = "P" Then
        Open "revpglpay.dat" For Output As AHandle
      End If
  End Select
  Print #AHandle, One
  Close AHandle
  
  VARevsAndGLsOK = True
  
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
  
  If Exist("revrglbill.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXRGLBAC.DAT") Then
      x = 1
      GoTo NoFileBill
    End If
    OpenRTaxGLInterBill RRHandle
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
      VARevsAndGLsOK = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revrglbill.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxBillGLSetUp.GThisYear = ThisYear
        frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
        frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
        frmVATaxBillGLSetUp.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        VARevsAndGLsOK = True
        KillFile "revrglbill.dat"
        TXLog ("ERROR: User warned that real billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the billing process anyway.")
      End If
    End If
  End If
  
  If Exist("revpglbill.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXPGLBAC.DAT") Then
      x = 1
      GoTo NoFileBillP
    End If
    OpenPTaxGLInterBill PPHandle
    Get PPHandle, 1, PRevRec
    Close PPHandle
    For x = 1 To 51
      If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(PRevRec.TaxAcct(x).PersCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PersDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FECRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FEDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
      End If
    Next x
NoFileBillP:
    If x < 52 Then
      VARevsAndGLsOK = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revpglbill.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxPBillGLSetUp.GThisYear = ThisYear
        frmVATaxPBillGLSetUp.fpListYear.SearchText = frmVATaxPBillGLSetUp.GThisYear
        frmVATaxPBillGLSetUp.fpListYear.ListIndex = frmVATaxPBillGLSetUp.fpListYear.SearchIndex
        frmVATaxPBillGLSetUp.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        VARevsAndGLsOK = True
        KillFile "revpglbill.dat"
        TXLog ("ERROR: User warned that personal billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the billing process anyway.")
      End If
    End If
  End If
  
  If Exist("revrglpay.dat") Then
    If Not Exist("TAXRGLACT.DAT") Then
      x = 1
      GoTo NoFilePay
    End If
    OpenRTaxGLInterPay RRHandle
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
      VARevsAndGLsOK = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") payment requirements. This needs to be fixed before continuing the payment process. Press F5 if you would like to jump to the payment General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the payment process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revrglpay.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxPayGLSetup.GThisYear = ThisYear
        frmVATaxPayGLSetup.fpListYear.SearchText = frmVATaxPayGLSetup.GThisYear
        frmVATaxPayGLSetup.fpListYear.ListIndex = frmVATaxPayGLSetup.fpListYear.SearchIndex
        frmVATaxPayGLSetup.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        KillFile "revrglpay.dat"
        VARevsAndGLsOK = True
        TXLog ("ERROR: User warned that real pay revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the payment process anyway.")
      End If
    End If
  End If
  
  If Exist("revpglpay.dat") Then
    If Not Exist("TAXPGLACT.DAT") Then
      x = 1
      GoTo NoFilePayP
    End If
    OpenPTaxGLInterPay PPHandle
    Get PPHandle, 1, PRevRec
    Close PPHandle
    For x = 1 To 51
      If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
        If QPTrim$(PRevRec.TaxAcct(x).PersCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PersDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MTDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MCDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FECRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).FEDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).MHDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenCRAcct) = "" Then
          Exit For
        End If
        If QPTrim$(PRevRec.TaxAcct(x).PenDBAcct) = "" Then
          Exit For
        End If
        If OptRev1 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt1CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt1DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev2 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt2CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt2DBAcct) = "" Then
            Exit For
          End If
        End If
        If OptRev3 = 1 Then
          If QPTrim$(PRevRec.TaxAcct(x).Opt3CRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).Opt3DBAcct) = "" Then
            Exit For
          End If
        End If
      End If
    Next x
NoFilePayP:
    If x < 52 Then
      VARevsAndGLsOK = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") payment requirements. This needs to be fixed before continuing the payment process. Press F5 if you would like to jump to the payment General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the payment process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revpglpay.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxPPayGLSetUp.GThisYear = ThisYear
        frmVATaxPPayGLSetUp.fpListYear.SearchText = frmVATaxPPayGLSetUp.GThisYear
        frmVATaxPPayGLSetUp.fpListYear.ListIndex = frmVATaxPPayGLSetUp.fpListYear.SearchIndex
        frmVATaxPPayGLSetUp.Show
        DoEvents
      Else
        Unload frmVATaxMsgWOpts
        KillFile "revpglpay.dat"
        VARevsAndGLsOK = True
        TXLog ("ERROR: User warned that personal pay revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the payment process anyway.")
      End If
    End If
  End If
  
  If Exist("revgladv.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If Not Exist("TAXRGLBAC.DAT") Then
      x = 1
      GoTo NoFileAdv
    End If
    OpenRTaxGLInterBill RRHandle
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
      VARevsAndGLsOK = False
      frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") advertising charge requirements. This needs to be fixed before continuing the advertising charges process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the advertising charges process."
      frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
      frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
      frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
      frmVATaxMsgW3Opts.Show vbModal
      If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "revgladv.dat"
        Exit Function
      ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
        Unload frmVATaxMsgWOpts
        frmVATaxBillGLSetUp.GThisYear = ThisYear
        frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
        frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
        frmVATaxBillGLSetUp.Show
        DoEvents
'        Unload frm
      Else
        Unload frmVATaxMsgWOpts
        VARevsAndGLsOK = True
        KillFile "revgladv.dat"
        TXLog ("ERROR: User warned that advertising charges revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the advertising charges process anyway.")
      End If
    End If
  End If
  
  If Exist("revglint.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If ThisType = "R" Then
      If Not Exist("TAXRGLBAC.DAT") Then
        x = 1
        GoTo NoFileIntR
      End If
      OpenRTaxGLInterBill RRHandle
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
NoFileIntR:
      If x < 52 Then
        VARevsAndGLsOK = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") real interest calculations requirements. This needs to be fixed before continuing the interest calculations process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the interest calculations process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglint.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxBillGLSetUp.GThisYear = ThisYear
          frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
          frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
          frmVATaxBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOK = True
          KillFile "revglint.dat"
          TXLog ("ERROR: User warned that real interest calculations revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the interest charges process anyway.")
        End If
      End If
    ElseIf ThisType = "P" Then
      If Not Exist("TAXPGLBAC.DAT") Then
        x = 1
        GoTo NoFileIntP
      End If
      OpenPTaxGLInterBill RRHandle
      Get RRHandle, 1, PRevRec
      Close RRHandle
      For x = 1 To 51
        If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
          If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
            Exit For
          End If
        End If
      Next x
NoFileIntP:
      If x < 52 Then
        VARevsAndGLsOK = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") personal interest calculations requirements. This needs to be fixed before continuing the interest calculations process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the interest calculations process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglint.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxPBillGLSetUp.GThisYear = ThisYear
          frmVATaxPBillGLSetUp.fpListYear.SearchText = frmVATaxPBillGLSetUp.GThisYear
          frmVATaxPBillGLSetUp.fpListYear.ListIndex = frmVATaxPBillGLSetUp.fpListYear.SearchIndex
          frmVATaxPBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOK = True
          KillFile "revglint.dat"
          TXLog ("ERROR: User warned that personal interest calculations revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the interest charges process anyway.")
        End If
      End If
    End If
  Else
    KillFile "revglint.dat"
  End If
  
  If Exist("revglman.dat") And TaxMasterRec.AcctgMethod <> "C" Then
    'cash doesn't have anything to do with billing
    If ThisType = "R" Then
      If Not Exist("TAXRGLBAC.DAT") Then
        x = 1
        GoTo NoFileManR
      End If
      OpenRTaxGLInterBill RRHandle
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
          If QPTrim$(RevRec.TaxAcct(x).IntCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).IntDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).PenCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(RevRec.TaxAcct(x).PenDBAcct) = "" Then
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
NoFileManR:
      If x < 52 Then
        VARevsAndGLsOK = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") real billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglman.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxBillGLSetUp.GThisYear = ThisYear
          frmVATaxBillGLSetUp.fpListYear.SearchText = frmVATaxBillGLSetUp.GThisYear
          frmVATaxBillGLSetUp.fpListYear.ListIndex = frmVATaxBillGLSetUp.fpListYear.SearchIndex
          frmVATaxBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOK = True
          KillFile "revglman.dat"
          TXLog ("ERROR: User warned that real manual billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the manual billing process anyway.")
        End If
      End If
    ElseIf ThisType = "P" Then
      If Not Exist("TAXPGLBAC.DAT") Then
        x = 1
        GoTo NoFileManP
      End If
      OpenPTaxGLInterBill RRHandle
      Get RRHandle, 1, PRevRec
      Close RRHandle
      For x = 1 To 51
        If PRevRec.TaxAcct(x).TaxYear = ThisYear Then
          If QPTrim$(PRevRec.TaxAcct(x).PersCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).PersDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MTCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MTDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MCCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MCDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).FECRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).FEDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MHCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).MHDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).IntCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).IntDBAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).PenCRAcct) = "" Then
            Exit For
          End If
          If QPTrim$(PRevRec.TaxAcct(x).PenDBAcct) = "" Then
            Exit For
          End If
          If OptRev1 = 1 Then
            If QPTrim$(PRevRec.TaxAcct(x).Opt1CRAcct) = "" Then
              Exit For
            End If
            If QPTrim$(PRevRec.TaxAcct(x).Opt1DBAcct) = "" Then
              Exit For
            End If
          End If
          If OptRev2 = 1 Then
            If QPTrim$(PRevRec.TaxAcct(x).Opt2CRAcct) = "" Then
              Exit For
            End If
            If QPTrim$(PRevRec.TaxAcct(x).Opt2DBAcct) = "" Then
              Exit For
            End If
          End If
          If OptRev3 = 1 Then
            If QPTrim$(PRevRec.TaxAcct(x).Opt3CRAcct) = "" Then
              Exit For
            End If
            If QPTrim$(PRevRec.TaxAcct(x).Opt3DBAcct) = "" Then
              Exit For
            End If
          End If
        End If
      Next x
NoFileManP:
      If x < 52 Then
        VARevsAndGLsOK = False
        frmVATaxMsgW3Opts.Label1.Caption = "Not all required General Ledger account numbers have been set up for this tax year's (" + CStr(ThisYear) + ") personal billing requirements. This needs to be fixed before continuing the billing process. Press F5 if you would like to jump to the billing General Ledger set up screen now. Press ESC to return to the menu. Otherwise, press F10 to continue the billing process."
        frmVATaxMsgW3Opts.cmdCont.Text = "F10 Continue"
        frmVATaxMsgW3Opts.CmdExit.Text = "ESC Exit"
        frmVATaxMsgW3Opts.cmdOption.Text = "F5 Jump"
        frmVATaxMsgW3Opts.Show vbModal
        If frmVATaxMsgW3Opts.fptxtChoice.Text = "abort" Then
          Unload frmVATaxMsgWOpts
          KillFile "revglman.dat"
          Exit Function
        ElseIf frmVATaxMsgW3Opts.fptxtChoice.Text = "option" Then
          Unload frmVATaxMsgWOpts
          frmVATaxPBillGLSetUp.GThisYear = ThisYear
          frmVATaxPBillGLSetUp.fpListYear.SearchText = frmVATaxPBillGLSetUp.GThisYear
          frmVATaxPBillGLSetUp.fpListYear.ListIndex = frmVATaxPBillGLSetUp.fpListYear.SearchIndex
          frmVATaxPBillGLSetUp.Show
          DoEvents
        Else
          Unload frmVATaxMsgWOpts
          VARevsAndGLsOK = True
          KillFile "revglman.dat"
          TXLog ("ERROR: User warned that personal manual billing revenue GL account numbers have not been set up for the current year (" + CStr(ThisYear) + ") and elected to continue the manual billing process anyway.")
        End If
      End If
    End If
  End If
  
End Function

