Attribute VB_Name = "NCTAX_STRUCT"
Option Explicit

Type NCPersonalRecType
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
'end of form
   CustPin        As Long
   NextRec        As Long
   LastYrPrinted  As Integer
   Deleted        As Integer
   VehTaxYear     As Integer
   DMVSubmitted   As String * 1
   InternalPin    As Long
   OptSearch      As String * 20 'added 8/16/06
   Blank          As String * 97
End Type


Type NCPropertyRecType
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
    Blank  As String * 55
End Type

Type NCTaxMasterType      'Master Default Information in Setup
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
  TaxYear As Integer
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
  CurrYrInt As Double '1/26/05
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
  OptRev1 As String * 35
  OptRev2 As String * 35
  OptRev3 As String * 35
  DiscXDate As Integer      'discount amount to calc on payment screen
  DisPct As Double
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
  OptSrchPers As String * 15 'added 8/16/06
  Padding As String * 98 '99
  AutoFillSrvAdd As String * 1
End Type

Type NCOptRealIdxType
  OptDesc As String * 20
  RealRec As Long
  RealPin As String * 20
End Type

Type NCTaxBillType
     CustRec            As Long                 'Acct #
     CustName           As String * 40
     CustAdd1           As String * 35
     CustAdd2           As String * 35
     CustAdd3           As String * 35
     CustZip            As String * 10
     RDesc1             As String * 30
     RDesc2             As String * 30
     RealPin            As String * 20
     PersPin            As String * 20
     RealValue          As Double
     PersValue          As Double
     ExptValue          As Double
     RealTaxDue         As Double
     PersTaxDue         As Double
     LateTaxDue         As Double
     TotalBillDue       As Double
     BillNumber         As Long                          'Recpt #
     TaxYear            As Integer
     BillPrinted        As Integer            '-1 = printed
     RealPropRecord     As Long
     PersPropRecord     As Long
     PriorYrBalance     As Double
     RealTaxRate        As Double
     PersTaxRate        As Double
     CustPin            As Long         'additional Protection for relinking
     TownShip           As String * 20
     MORTCODE           As String * 8
     LotOrAcre          As String * 1
     LASize             As String * 9
     MortRec            As Integer
     CarShore           As Double
     RDesc3             As String * 30
     InternalPin        As Long 'added 5/12/05
     OptRevTax1         As Double 'added 5/12/05
     OptRevTax2         As Double 'added 5/12/05
     OptRevTax3         As Double 'added 5/12/05
     OverPayAmt         As Double 'added 5/24/05
     SetDscvry2No       As String * 1 'added 12/5/06
     Padding            As String * 104
End Type

Type NCOptRevRateTablesType
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
  Cushion As String * 100
End Type

Type NCWinTAXGLAcctRecType
  TaxYear       As Integer        'protected
  TaxDBAcct     As String * 14
  TaxCRAcct     As String * 14
  IntDBAcct     As String * 14
  IntCRAcct     As String * 14
  AdvDBAcct     As String * 14
  AdvCRAcct     As String * 14
  Fill1         As String * 1     'protected
  LtLstDBAcct   As String * 14
  LtLstCRAcct   As String * 14
  Opt1DBAcct    As String * 14
  Opt1CRAcct    As String * 14
  Opt2DBAcct    As String * 14
  Opt2CRAcct    As String * 14
  Opt3DBAcct    As String * 14
  Opt3CRAcct    As String * 14
End Type

Type NCTaxAcctsType
  TaxAcct(1 To 51) As NCWinTAXGLAcctRecType
End Type

Type NCWinRevSourceType
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

Type NCTaxTransactionType
  TransDate    As Integer          'Transaction Date
  TaxYear      As Integer          'Must Contain Full 4 digit Tax Year Here
  TranType     As Integer          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
                                   '10=DwnAdjPay 11=UpAdjPay
                                   '22=PrePayment 23=Refund Prepayment added 3-25-03
  BillType     As String * 1       'R=Real P=Personal Property C=Combined (NC/
  Amount       As Double           'Total Transaction Amount
  Revenue      As NCWinRevSourceType    'See Revenue Source Type File above
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
  CntyPara     As String * 20
  CyclPara     As String * 20
  TShpPara     As String * 25
End Type

Type NCMortCodeRecType
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

Type NCTxBill1DefaultsType
    TxtHead1 As String * 50
    TxtHead2 As String * 50
    txtOpt1 As String * 40
    TxtOpt2 As String * 40
    TxtOpt3 As String * 40
    TxtOpt4 As String * 40
    txtPgph0 As String * 125
    txtPgph1 As String * 125
    txtPgph2 As String * 125
    txtPgph3 As String * 125
    txtPgph4 As String * 125
    txtPgph5 As String * 125
    txtPgph6 As String * 125
    txtPgph7 As String * 125
    TxtOpt5 As String * 75
    txtHead4 As String * 40
    txtHead5 As String * 40
    txtHead6 As String * 40
    TxtOpt6 As String * 45
    TxtOpt7 As String * 75
    dologo As Integer  '0 for no 1 for yes
    UseBarCode As Boolean
End Type

Type NCTAXLateLetterType
  Head1    As String * 40
  Head2    As String * 40
  Head3    As String * 40
  Head4    As String * 40
  Head5    As String * 40
  Body(1 To 20) As String * 75
End Type


