Attribute VB_Name = "modGLFileType"
Option Explicit

Type GLSetupRecType                 'still under const.
   UserName    As String * 30
   TotAcctLen  As Integer
   FundLen     As Integer
   AcctLen     As Integer
   DetLen      As Integer
   CashAcct    As String * 14
   APAcct      As String * 14
   EncAcct     As String * 14
   FBAcct      As String * 14
   FYBeg       As Integer
   FYEnd       As Integer
   NYBeg       As Integer
   NYEnd       As Integer
   CDCash      As String * 14  'new
   CDDue       As String * 14
   CDActive    As String * 1
   CRCashAcct  As String * 14
   DeptCode    As String * 1
   LPDate      As Integer
   HPDate      As Integer
   CDCashAcct  As String * 14
   CDsbCash    As String * 14
End Type

Type GLFundIndexType                'Fund Index : 16 bytes
   FundNum     As String * 4        'Fund Number
   RecNum      As Integer           'Pointer to record
End Type

Type GLFundRecType                  'Fund Record Type: 64 bytes
   DELETED     As Integer           'Deleted Flag
   FundNum     As String * 4        'Fund Code
   title       As String * 30       'Fund Title
   Res         As String * 28       'Reserve for future needs
End Type

Type GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     As String * 14       'Formatted account Number string
   RecNum      As Integer           'Pointer to record
End Type

Type GLAcctRecType                  'Account Record Type: ? bytes
   DELETED     As Integer           'Active Account Flag
   Num         As String * 14       'Formatted Account Number
   title       As String * 30       'Account Description
   Typ         As String * 1        'Account Type
   FrstTran    As Long              'Pointer to First Trans
   LastTran    As Long              'Pointer to Last Trans
   PYAct       As Double            'Prior Year Actual
   BegBal      As Double            'Beginning Balance
   Bgt         As Double            'Budget Amount
   Bal         As Double            'Running Balance
   Encumb      As Double            'Encumbered Amount
   MTD         As Double            'Month to Date Bal (calc as needed)
   YTD         As Double            'Year to Date Bal (calc as needed)
   NYEst       As Double            'Bgt Estimate
   NYReq       As Double            'New Year Bgt Requested
   NYRec       As Double            'New Year Bgt Recommended
   NYApp       As Double            'New Year Bgt Approved
   FrstBTran   As Integer           'Pointer to First Budget Trans
   LastBTran   As Integer           'Pointer to Last Budget Trans
   FrstPTran   As Integer           'Pointer to First Budget Trans
   LastPTran   As Integer           'Pointer to Last Budget Trans
   'Res         AS STRING * 25       'Reserved for future needs
   Work        As Double            'Temp added 08/17/96 for closeout
   Res         As String * 17
   Marked      As Integer           '
End Type

Type GLDeptIndexType                'Dept Index
   DeptNum     As String * 8        'Dept Number
   RecNum      As Integer           'Pointer to record
End Type

Type GLDeptRecType                  'Dept Record Type
   DELETED     As Integer           'Deleted Flag
   DeptNum     As String * 8        'Fund Code
   title       As String * 30       'Fund Title
   Res         As String * 20       'Reserve for future needs
End Type

Type GLTransRecType                 'Transaction Record: 96 bytes
   AcctRec     As Integer           'Pointer to Acct Record
   AcctNum     As String * 14       'Formatted Acct Number string
   TRDATE      As Integer           'Date2Num function
   DESC        As String * 20       'Transaction Description
   Ref         As String * 8        'Document Reference
   DrAmt       As Double            'Debit Amount
   CrAmt       As Double            'Credit Amount
   Src         As String * 8        'Module Source Code
   NextTran    As Long              'Pointer to Next Trans
   Res         As String * 20       'Reserved for future needs
   Marked      As Integer
End Type

Type TrEditRecType                  'Experimental GJ edit record:
   DELETED     As Integer           'Deleted transaction flag
   Posted      As Integer           'Posted flag
   AcctRec     As Integer           'Pointer to acct rec
   AcctNum     As String * 14       'Formatted Acct number string
   AcctName    As String * 30       'Account Title
   TRDATE      As Integer           'Date2Num
   DrAmt       As Double            'Transaction Debit Amount
   CrAmt       As Double            'Transaction Credit Amount
   EType       As String * 1        'Entry Type (Debit/Credit)
   DESC        As String * 20       'Transaction Description
   Ref         As String * 8        'Document Reference #
   Src         As String * 8        'Module Source Code
   'Res         AS STRING *          'Reserve for future needs
End Type

Type TrSortType                     'Used for sorting trans in history rpt
   TRDATE     As Integer            'Transaction Date
   Record     As Long               'Pointer to transaction record
End Type

Type TrSortType1                    'Used for sorting trans in history rpt
   TRDATE     As String * 12             'Transaction Date
   Record     As Long               'Pointer to transaction record
End Type

Type TrSortType2                    'Used for sorting trans in history rpt
   TRDATE     As Integer            'Transaction Date
   Record     As Long               'Pointer to transaction record
   Fill       As String * 2
End Type
'Check Rec File
Type OSChkRecType
   ChkNum   As Single        '4 AS chknum$
   ChkDate  As String * 8    '8 AS chkdate$
   DESC     As String * 30   '30 AS chkdesc$
   Amt      As Single        '4 AS chkamt$
   Src      As Integer       '2 AS CHKSOURCE$
   Cleared  As Integer       'added by JB
   BankCode As Integer
   Filler   As String * 12   '16 AS nul$
End Type

Type IFRecType
   AcctNum As String * 9      '9 AS tranacct$
   TRDATE As String * 8       '8 AS trandate$
   DESC As String * 20        '20 AS trandesc$
   CrAmt As Double            '8 AS cramt$
   DrAmt As Double            '8 AS dramt$
   Ref As String * 8          '8 AS detail$
   Src As String * 8          '8 AS source$
   Filler As String * 14      '4 AS nexttr$
   Posted As Integer
End Type

Type GLFBAdjRecType
   AcctNum  As String * 16
   AdjAmt As Double
End Type

Type GLFundCloseRecType
   FundNum As String * 3
End Type

Type GLBankRecType   '128 bytes
   DELETED As Integer
   BankNum As Integer
   BankName As String * 25
   BankAcct As String * 25
   GLAcct As String * 25
   Pad As String * 49
End Type

Type GLSumSetupRecType                 'still under const.
   Beg1  As String * 6
   End1  As String * 6
   Desc1 As String * 30
   Beg2  As String * 6
   End2  As String * 6
   Desc2 As String * 30
   Beg3  As String * 6
   End3  As String * 6
   Desc3 As String * 30
   Beg4  As String * 6
   End4  As String * 6
   Desc4 As String * 30
   Beg5  As String * 6
   End5  As String * 6
   Desc5 As String * 30
   Beg6  As String * 6
   End6  As String * 6
   Desc6 As String * 30
   Beg7  As String * 6
   End7  As String * 6
   Desc7 As String * 30
   Beg8  As String * 6
   End8  As String * 6
   Desc8 As String * 30
   Beg9  As String * 6
   End9  As String * 6
   Desc9 As String * 30
   Pad   As String * 75
End Type

Type CJDistType
    DACN      As String * 16
    DACNM     As String * 20
    'DACREC    As String * 2
    DACREC    As Integer
    DAMT      As Double
End Type


Type CJEditRecType
    DELFLAG   As Integer
    TRDATE    As Integer
    DESC      As String * 20
    DOCREF    As String * 8
    Amt       As Double
    BATCHNUM  As String * 8
    RECCODE   As String * 2 'AS INTEGER  BankNumber????
    Dist(1 To 36)  As CJDistType
End Type


'AP.BI

'--Vendor Index
Type VendorIdxRecType
   VendorCode As String * 10
   RecNum As Integer
End Type


Type VendorRecType
   VIN        As Long
   vnum       As String * 10
   VNAME      As String * 30
   Addr1      As String * 30
   Addr2      As String * 30
   City       As String * 22
   State      As String * 2
   Zip        As String * 10
   PayToName  As String * 30
   PayToAddr  As String * 30
   PaytoAddr2 As String * 30
   PayToCity  As String * 22
   PaytoState As String * 2
   PaytoZip   As String * 10
   phone      As String * 14
   VTERMS     As Integer
   pad2 As String * 5
   FedID      As String * 12
   CoCode     As String * 3
   StCode     As String * 2
   YTDPay     As Double
   Get1099  As String * 1
   CurrBal    As Double
   FrstTran   As Long
   LastTran   As Long
   FrstPO     As Long
   LastPO     As Long
   DefDist    As Long
   DELFLAG    As Integer
   'Pad        AS STRING * 45
   Contact    As String * 30
   Fax        As String * 14
   Pad As String * 1
End Type

'--Distribution type for invoice edit
Type DistType
   DACN       As String * 16
   DACNM      As String * 20
   DACREC     As String * 2      'this is an integer rec number
   'DACREC     AS INTEGER
   DAMT       As Double
End Type
Type APPOType
    DELFLAG  As Integer
    VENDOR   As String * 10
    VENDNAME As String * 20
    PONUM    As String * 15
    POAMT   As Double
    PODATE   As Integer
    Dist(1 To 36) As DistType
End Type

'--Type for Invoice Edit
Type APInvType
    DELFLAG   As Integer
    VENDOR    As String * 10
    VENDNAME  As String * 20
    VRecNum   As String * 2
    PONUM     As String * 18
    INVNUM    As String * 25
    INVAMT    As Double
    PAYCODE   As String * 1
    INVDATE   As Integer
    DUEDATE   As Integer
    DISTDATE  As Integer
    Dist(1 To 36) As DistType
    POAPLRecNum  As String * 4        'this is a long integer pointer
    POFLAG    As Integer           'Flag to indicate an active PO
End Type

'--Type for Invoice Edit
Type APInv85Type
    DELFLAG  As Integer
    VENDOR   As String * 10
    VENDNAME As String * 20
    VRecNum  As String * 2
    INVNUM   As String * 25
    PONUM    As String * 18
    INVAMT   As Double
    INVDESC  As String * 33
    TAXYN    As String * 1
    PAYCODE  As String * 1
    INVDATE  As Integer
    DUEDATE  As Integer
    DISTDATE As Integer
    PSLFlag  As String * 1
    Get1099  As String * 1
    STAXAMT  As Double
    CTAXAMT  As Double
    GRANDTOT As Double
'**************************
    Dist(1 To 36) As DistType     '
    POAPLRecNum  As String * 4     'this is a long integer pointer
    POFLAG    As Integer           'Flag to indicate an active PO
'Added this for multi PO's
    PORecs(1 To 6)  As String * 4
End Type

Type APLedger81RecType             'version for Troy's paid supply list
   VIN          As Integer
   VendorCode   As String * 10
   VRecNum      As Integer
   TRDATE       As Integer
   GLDistDate   As Integer
   DUEDATE      As Integer
   TrCode       As Integer      '1=Invoice, 4=PO, 3=Check, -3=Void Check, -4=Cleared PO
   DOCNum       As String * 25
   PONUM        As String * 20
   PAYCODE      As Integer
   PrintCode    As Integer
   PDCheckNum   As Long
   PDCheckDate  As Integer
   Comment      As String * 31          'Reduced to 31 bytes to allow for dept to be stored (wrightsville beach)
   DeptNumb     As Long
   PSLFlag      As String * 1
   Get1099      As String * 1
   Amt          As Double
   FrstDist     As Long
   LastDist     As Long
   NextTrans    As Long
   TaxAmt       As Double
   Pad           As String * 2
End Type

'--A/P Ledger
Type APLedgerRecType  'version 8.0
   VIN        As Integer
   VendorCode As String * 10
   VRecNum    As Integer
   TRDATE     As Integer
   GLDistDate As Integer
   DUEDATE    As Integer
   TrCode     As Integer
   DOCNum     As String * 25
   PONUM      As String * 20
   PAYCODE    As Integer
   PrintCode  As Integer
   PDCheckNum   As Long
   PDCheckDate  As Integer
   MiscCode     As String * 23
   Amt          As Double
   FrstDist     As Long
   LastDist   As Long
   NextTrans  As Long
End Type

Type APLedgerRecType0
  VendorCode As Integer                 '2 AS vennum$
  TRDATE As String * 8             '8 AS INVDATE$
  DOCREF As String * 25           '25 AS invnum$
  PONUM As String * 10             '10 AS ponum$
  Amt As Double                    '8 AS amount$
  PAYCODE As Integer               '2 AS PAYCODE$
  FirstDist As Single              '4 AS fdist$
  LastDist As Single               '4 AS ldist$
  NextTrans As Single              '4 AS ndata$
  CoTaxCode As String * 3          '3 AS tctycode$
  StTaxCode As String * 3          '3 AS tstcode$
  CoTaxAmt As Single               '4 AS ctaxamt$
  StTaxAmt As Single               '4 AS staxamt$
  Fill As String * 15
End Type

'--A/P Ledger Accounting Distributions
Type APDistRecType
   APLedgerRec As Long
   DistAcctRec As Integer
   'DistAcctRec AS STRING * 2
   'Changed distacctnum from 16 to 15
   'and added diststat for flag for partial po's 3-12-02
   DistAcctNum As String * 15
   DistStat    As String * 1  'Blank = Not Used, T = Tagged, L = Liquidated
   DistAmt     As Double
   NextDist    As Long
End Type

'--Distributions Summary work array type
Type DistSumType
   DistAcctNum As String * 16
   AcctTitle   As String * 20
   DistAmt     As Double
End Type

'--Vendor Default Distribution
'need to revise

Type DefDistAcctsRecType
   DefAcct      As String * 16
   DefAcctName  As String * 20
   DefPct       As Single
End Type

Type VendorDefDistRecType
   VRecNum    As Integer
   DefDist(1 To 8)  As DefDistAcctsRecType
End Type

'*********************************

Type InvTaxDefType
  ACCTNO As String * 16
  TaxAmt As Double
End Type

Type APInvTaxRecType
  InvTax(1 To 2) As InvTaxDefType
  AUTODIST As String * 1
End Type

'Remember to change this name in all modules (changed here 08-24-01)
'Type GLInvTaxRecType
'    AcctNo  As String * 16
'    TaxAmt  As Double
'End Type

Type POItemsRecType
    STKNO    As String * 8
    DESC     As String * 40
    'change desca to desc uped to 40 from 20 and
    'eliminated second descb
    'DESCB    As String * 20
    QUAN     As Double
    PRICE    As Double
    EXT      As Double
    ACCTNO   As String * 14
    AcctRec  As Integer
End Type

Type POFORMRecType2
    DELETED  As Integer
    PONUM    As String * 8
    REQNUM   As String * 7              'dept number here
    PODATE   As Integer
    VNDRCODE As String * 10
    VNDRREC  As Integer
    VNDRINF1 As String * 30
    VNDRINF2 As String * 30
    VNDRINF3 As String * 30
    VNDRINF4 As String * 30
    VNDRINF5 As String * 30
    SHPLINE1 As String * 30
    SHPLINE2 As String * 30
    SHPLINE3 As String * 30
    SHPLINE4 As String * 30
    SHPLINE5 As String * 30
    FOB      As String * 20
    Shipvia  As String * 20
    Terms    As String * 20
    SHIPON   As String * 20
    Addinst1 As String * 30
    Addinst2 As String * 30
    Addinst3 As String * 30
    POAMT    As Double
    ITEMS(1 To 36) As POItemsRecType
    'Change from (1 to 12) to (1 to 36) for windows ver because could fit all 36 on spread
    'ITEMS(1 To 12) As POItemsRecType
    DUMMY    As String * 1
End Type

Type POControlRecType
    PONumber As String * 8
    Header1  As String * 35
    Header2  As String * 35
    Header3  As String * 35
    Header4  As String * 35
    Shipto1  As String * 30
    Shipto2  As String * 30
    Shipto3  As String * 30
    Shipto4  As String * 30
    Shipto5  As String * 30
    FOB      As String * 20
    Shipvia  As String * 20
    Terms    As String * 20
    Addinst1 As String * 20
    Addinst2 As String * 20
    Addinst3 As String * 20
    Pading   As String * 94

End Type

Type AcctPOChkType
    Acct    As String * 14
    Bgt     As Double
    Encumb  As Double
    Bal     As Double
    POTotal As Double
    NYApp   As Double
    NYEncmb As Double
    Fill    As String * 2
End Type


