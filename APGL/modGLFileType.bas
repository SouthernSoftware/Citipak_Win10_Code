Attribute VB_Name = "modGLFileType"
Option Explicit

Type GLSetupRecType                 'V205 added new fields noted below
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
   APChkCode   As Integer
   POStop      As Boolean          'new 7/22/02 for potab on invoice entry
 'Fields added for V205
   PSLFlag     As Integer   '1 for default to Yes, 0 for No
   DupInvFlag  As Integer   '1 to allow duplicates, 0 for No
   CRBank      As Integer   'banknum as default on entry
   CDBank      As Integer   'banknum for default on entry
   ChkBank     As Integer   'banknum for default on check printing
   Pad         As String * 20
   ChkVer      As String * 4   ' for "V205"
End Type

Type GLFundIndexType                'Fund Index : 16 bytes
   FundNum     As String * 4        'Fund Number
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLFundRecType                  'Fund Record Type: 64 bytes
   Deleted     As Integer           'Deleted Flag
   FundNum     As String * 4        'Fund Code
   Title       As String * 30       'Fund Title
   Res         As String * 28       'Reserve for future needs
End Type
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'ADDED Function 6-11-04 For GASB34
Type GLFNCTIndexType                'Function Index
   FnctNum     As String * 5        'Function Number
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLFNCTRecType                  'Function Record Type: 64 bytes
   Deleted     As Integer           'Deleted Flag
   FnctNum     As String * 5        'Function Code
   Title       As String * 30       'Function Title
   Res         As String * 27       'Reserve for future needs
End Type
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Type GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     As String * 14       'Formatted account Number string
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLAcctRecType                  'Account Record Type: ? bytes
   Deleted     As Integer           'Active Account Flag
   Num         As String * 14       'Formatted Account Number
   Title       As String * 30       'Account Description
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
'edit the res added function rec pointer 6/11/04
   FNCTRec     As Long
   Res         As String * 12
   ChkByte     As String * 1    'this is updated at GASB34 conversion with chr$(1)
   Marked      As Integer           '
End Type

Type GLDeptIndexType                'Dept Index
   DeptNum     As String * 8        'Dept Number
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLDeptRecType                  'Dept Record Type
   Deleted     As Integer           'Deleted Flag
   DeptNum     As String * 8        'Fund Code
   Title       As String * 30       'Fund Title
   Res         As String * 20       'Reserve for future needs
End Type
'' V204 and before
''Type GLTransRecType                 'Transaction Record: 96 bytes
''   AcctRec     As Integer           'Pointer to Acct Record
''   AcctNum     As String * 14       'Formatted Acct Number string
''   TRDATE      As Integer           'Date2Num function
''   Desc        As String * 20       'Transaction Description
''   Ref         As String * 8        'Document Reference
''   DrAmt       As Double            'Debit Amount
''   CrAmt       As Double            'Credit Amount
''   Src         As String * 8        'Module Source Code
''   NextTran    As Long              'Pointer to Next Trans
''   Res         As String * 20       'Reserved for future needs
''   Marked      As Integer
''End Type

Type GLTransRecType                 'Transaction Record: 128 bytes
   AcctRec     As Integer           'Pointer to Acct Record
   AcctNum     As String * 14       'Formatted Acct Number string
   TRDATE      As Integer           'Date2Num function
   Desc        As String * 20       'Transaction Description
  'add extra desc for V205
   LDesc       As String * 32       'Originally for invoice desc but use for other trans as well
   Ref         As String * 8        'Document Reference
   DrAmt       As Double            'Debit Amount
   CrAmt       As Double            'Credit Amount
   Src         As String * 8        'Module Source Code
   NextTran    As Long              'Pointer to Next Trans
   Res         As String * 19       'Reserved for future needs
   ChkByte     As String * 1        'chr$(1) for V205
   Marked      As Integer
End Type
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'old one   new one for v205 with extra desc
''Type TrEditRecType                  'Experimental GJ edit record:
''   Deleted     As Integer           'Deleted transaction flag
''   Posted      As Integer           'Posted flag
''   AcctRec     As Integer           'Pointer to acct rec
''   AcctNum     As String * 14       'Formatted Acct number string
''   AcctName    As String * 30       'Account Title
''   TRDATE      As Integer           'Date2Num
''   DrAmt       As Double            'Transaction Debit Amount
''   CrAmt       As Double            'Transaction Credit Amount
''   EType       As String * 1        'Entry Type (Debit/Credit)
''   Desc        As String * 19       'changed from 20 to 19 Transaction Description
''   LOCKED      As Byte              'added 6-28-02 flag to prevent dual editing
''   Ref         As String * 8        'Document Reference #
''   Src         As String * 8        'Module Source Code
''   'Res         AS STRING *          'Reserve for future needs
''End Type
Type TrEditRecType                  'New Edit Rec Type for V205
   Deleted     As Integer           'Deleted transaction flag
   Posted      As Integer           'Posted flag
   AcctRec     As Integer           'Pointer to acct rec
   AcctNum     As String * 14       'Formatted Acct number string
   AcctName    As String * 30       'Account Title
   TRDATE      As Integer           'Date2Num
   DrAmt       As Double            'Transaction Debit Amount
   CrAmt       As Double            'Transaction Credit Amount
   EType       As String * 1        'Entry Type (Debit/Credit)
   Desc        As String * 20       'back to 20
   LDesc       As String * 32       'New Extra desc for V205
   LOCKED      As Byte              'added 6-28-02 flag to prevent dual editing
   Ref         As String * 8        'Document Reference #
   Src         As String * 8        'Module Source Code
   'Res         AS STRING *          'Reserve for future needs
End Type

Type ChkSortType                    'Use for check listing
   CHKinfo     As String * 14
   Record      As Long
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
   fill       As String * 2
End Type
'Check Rec File
Type OSChkRecType
   ChkNum   As Single        '4 AS chknum$
   chkdate  As String * 8    '8 AS chkdate$
   Desc     As String * 30   '30 AS chkdesc$
   Amt      As Single        '4 AS chkamt$
   Src      As Integer       '2 AS CHKSOURCE$
   Cleared  As Integer       'added by JB
   Bankcode As Integer
   Filler   As String * 12   '16 AS nul$
End Type
'New Check Rec File  2002
'Changed Amt from single to double
'and Chknum to long
'and chkdate from string of 8 to integer
Type OSChekRecType
   ChkNum   As Long         '4 AS chknum$
   chkdate  As Integer       '8 AS chkdate$
   Desc     As String * 30   '30 AS chkdesc$
   Amt      As Double        '4 AS chkamt$
   Src      As Integer       '2 AS CHKSOURCE$
   Cleared  As Integer       'added by JB
   Bankcode As Integer
   VoidFlag As Integer       'added for V205  1 for void
   Filler   As String * 10   '16 AS nul$
End Type
Type OSChekSrtType
  ChkNum As Long
  RecNo  As Long
End Type

Type GLFBAdjRecType
   AcctNum  As String * 16
   AdjAmt As Double
End Type

Type GLFundCloseRecType
   FundNum As String * 3
End Type

Type GLBankRecType   '128 bytes
   Deleted As Integer
   BankNum As Integer
   BankName As String * 25
   BankAcct As String * 25
   GLAcct As String * 25
   Pad As String * 49
End Type

Type GLSumSetupRecType                 'still under const.
   Beg1  As String * 6
   End1  As String * 6
   DESC1 As String * 30
   Beg2  As String * 6
   End2  As String * 6
   DESC2 As String * 30
   Beg3  As String * 6
   End3  As String * 6
   DESC3 As String * 30
   Beg4  As String * 6
   End4  As String * 6
   DESC4 As String * 30
   Beg5  As String * 6
   End5  As String * 6
   DESC5 As String * 30
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
   '******
    DAMT      As Double
End Type

Type CJEditRecType
    DelFlag   As Integer
    TRDATE    As Integer
    Desc      As String * 20        'changed bact to 20 from 19
    LDesc     As String * 32        'new long desc for V205
    LOCKED    As Byte               'added 6-28-02 for flag to prevent dual editing
    DOCREF    As String * 8
    Amt       As Double
    BATCHNUM  As String * 8
    RECCODE   As Integer    'for banknum
    Dist(1 To 36)  As CJDistType
End Type


'AP.BI

'--Vendor Index
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'This One too?????
Type VendorIdxRecType
   VendorCode As String * 10
   RecNum As Integer
End Type

'New Type 12-2003 for DBA Field
Type VendorRecType
   VIN        As Long
   vnum       As String * 10
   VNAME      As String * 30
   ADDR1      As String * 30
   ADDR2      As String * 30
   City       As String * 22
   STATE      As String * 2
   Zip        As String * 10
   PaytoName  As String * 30
   PaytoAddr  As String * 30
   PaytoAddr2 As String * 30
   PayToCity  As String * 22
   PaytoState As String * 2
   PaytoZip   As String * 10
   Phone      As String * 14
   VTerms     As Integer
   pad2       As String * 5
   Fedid      As String * 12
   CoCode     As String * 3
   StCode     As String * 2
   YTDPay     As Double
   Get1099    As String * 1
   CurrBal    As Double
   FrstTran   As Long
   LastTran   As Long
   FrstPO     As Long
   LastPO     As Long
   DefDist    As Long
   DelFlag    As Integer
   Contact    As String * 30
   Fax        As String * 14
   DBA        As String * 30  'added 12-2003
   Memo       As String * 30  'added for V205
   ActiveFlag As Integer      'added for V205  0 FOR ACTIVE, 1 FOR INACTIVE
   Pad As String * 66         'while in the neighborhood add extra space
   ChkByte    As String * 1   'chr$(1) for V205
End Type

'--Distribution type for invoice edit
'Type DistType
'   DACN       As String * 16
'   DACNM      As String * 20
'   DACREC     As String * 2      'this is an integer rec number
'   'DACREC     AS INTEGER
'   DAMT       As Double
'End Type

'New Distribution for AP 3/22/02 for invoices also using partial POs
Type INVDistType
    DISTNUM  As Long
    DACREC   As Integer
    '*****
    DACODE   As String * 1
    DACN     As String * 15
    DACNM    As String * 20
    DAMT     As Double
End Type


'--Type for Invoice Edit
'Type APInv85Type
'    DelFlag  As Integer
'    Vendor   As String * 10
'    VendName As String * 20
'    VRecNum  As Integer
'    '******
'    InvNum   As String * 25
'    PONum    As String * 18
'    InvAmt   As Double
'    INVDESC  As String * 32      'Was 33
'    LOCKED   As Byte             '6-28-02 added locked recored flag to prevent dual editing
'    TAXYN    As String * 1
'    PAYCODE  As String * 1
'    InvDate  As Integer
'    DueDate  As Integer
'    DISTDATE As Integer
'    PSLFlag  As String * 1
'    Get1099  As String * 1
'    STAXAMT  As Double
'    CTAXAMT  As Double
'    GRANDTOT As Double
''Added this for multi PO's
'    'PORecs(1 To 6)  As String * 4
''fix it this way for new version - no need of multi pos
'    POAPLRecNum  As Long   'this is a long integer pointer
'    POFLAG    As Integer   'Flag to indicate an active PO
'    POLINES   As Integer    'Num of Dist Per PO
'    POUSED    As Integer    'Num of Dist from PO used on invoice
''**************************
'    Dist(1 To 36) As INVDistType     '
''***************************
'End Type
Type APInv85Type
    DelFlag  As Integer
    Vendor   As String * 10
    VendName As String * 20
    VRecNum  As Integer
    '******
    InvNum   As String * 25
    PONum    As String * 20
    MPONum   As String * 20
    InvAmt   As Double
    INVDESC  As String * 32      'Was 33
    LOCKED   As Byte             '6-28-02 added locked recored flag to prevent dual editing
    TAXYN    As String * 1
    PAYCODE  As String * 1
    InvDate  As Integer
    DueDate  As Integer
    DISTDATE As Integer
    PSLFlag  As String * 1
    Get1099  As String * 1
    STAXAMT  As Double
    CTAXAMT  As Double
    GRANDTOT As Double
'Added this for multi PO's
    'PORecs(1 To 6)  As String * 4
'fix it this way for new version - no need of multi pos
    POAPLRecNum  As Long   'this is a long integer pointer
    POFLAG    As Integer   'Flag to indicate an active PO
    POLINES   As Integer    'Num of Dist Per PO
    POUSED    As Integer    'Num of Dist from PO used on invoice
'**************************
    Dist(1 To 36) As INVDistType     '
'***************************
End Type

Type APLedger81RecType             'version for Troy's paid supply list
   VIN          As Integer
   VendorCode   As String * 10
   VRecNum      As Integer
   TRDATE       As Integer
   GLDistDate   As Integer
   DueDate      As Integer
   TRCode       As Integer      '1=Invoice, 4=PO, 3=Check, -3=Void Check, -4=Cleared PO
   DOCNum       As String * 25
   PONum        As String * 20
   MPONum       As String * 20   'added for V205
   PAYCODE      As Integer
   PrintCode    As Integer
   PDCheckNum   As Long
   PDCheckDate  As Integer
   Comment      As String * 32          'made 32 again for V205'Reduced to 31 bytes to allow for dept to be stored (wrightsville beach)
   DeptNumb     As Long
   PSLFlag      As String * 1
   Get1099      As String * 1
   Amt          As Double
   FrstDist     As Long
   LastDist     As Long
   NextTrans    As Long
   TaxAmt       As Double
   'Pad          As String * 2
'Change pad to Bankcode to allow correct acct update during check void 7-29-02
   Bankcode     As Integer
   Pad          As String * 30        'added for V205 to make total reclen 196
   ChkByte      As String * 1         'chr$(1) for V205
End Type


'--A/P Ledger Accounting Distributions
Type APDistRecType
   APLedgerRec As Long
   DistAcctRec As Integer
   '********
   'DistAcctRec AS STRING * 2
   'Changed distacctnum from 16 to 15
   'and added diststat for flag for partial po's 3-12-02
   DistAcctNum As String * 15
   DistStat    As String * 1  'Blank =Not Used, L =Liquidated T =Tagged on invoice
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
   '******
   DefDist(1 To 8)  As DefDistAcctsRecType
End Type

'*********************************
Type LuneySortType
   Vendor As String * 10
   TRRec  As Long
   fill   As String * 4
End Type

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
    Desc     As String * 40
    'change desca to desc uped to 40 from 20 and
    'eliminated second descb
    'DESCB    As String * 20
    QUAN     As Double
    PRICE    As Double
    EXT      As Double
    ACCTNO   As String * 14
    AcctRec  As Integer
    '******
End Type

Type POFORMRecType2
    Deleted  As Integer
    PONum    As String * 8
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
    POAmt    As Double
    ITEMS(1 To 36) As POItemsRecType
    'Change from (1 to 12) to (1 to 36) for windows ver because could fit all 36 on spread
    'ITEMS(1 To 12) As POItemsRecType
    LOCKED   As Byte         'changed from dummy 6-28-02 for flag to prevent dual editing
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
    fill    As String * 2
End Type

  Type TPayListType
    LedgerRecNum As Long
    VendorRecNum As Integer
  End Type
  
Type TPayNotListType
    VendorRecNum As Integer
    Amt          As Double
End Type

  Type CheckInfoType2
    Ledger1st     As Long
    LedgerLst     As Long
    StartChk      As Long
    LastChk       As Long
    ChkAmt        As Double
    chkdate       As Integer
    VendorRecNum  As Integer
  End Type

  Type oCheckInfoType
    Ledger1st     As Long
    LedgerLst     As Long
    ChkNum        As Long
    ChkAmt        As Double
    chkdate       As Integer
    VendorRecNum  As Integer
  End Type

  Type CheckInfoType3
    ListFirst     As Integer
    ListLast      As Integer
    StartChk      As Long
    LastChk       As Long
    ChkAmt        As Double
    chkdate       As Integer
    VendorRecNum  As Integer
    VoidFlag      As Integer
    Bankcode      As Integer 'To track o/s checks by bank
  End Type

  Type LedgerInfoType
    LedDate       As String * 17
    LedInvNum     As String * 25
    InvAmt        As String * 17
  End Type

  Type DistInfoType
    Fill1         As String * 19
    DistAcct      As String * 16
    DistAmt       As String * 23
    'InvAmt        AS STRING * 16
  End Type

  Type CheckRegType
    ChkNum     As String * 10
    chkdate    As String * 13
    VendName   As String * 30
    ChkAmt     As String * 14
  End Type

'  Type FLen
'    V As String * 42
'  End Type
'
'  Type FLen2
'    V As String * 64
'  End Type
'
'  Type FLen3
'    V As String * 32
'  End Type

Type LedgerInfoType2
  InvDate       As String * 12
  DueDate       As String * 12
  InvNum        As String * 27
  PONum         As String * 14
  Amt           As String * 12
  DistAcct      As String * 16
  DistAmt       As String * 12
End Type

Type AP1099RecType
    Deleted  As Integer
    RecID    As String * 15
    RecName  As String * 30
    'Add dba line 12-2003
    DBA      As String * 30
    RecADDR  As String * 30
    RecADDR2 As String * 30
    RecCSZ   As String * 30
    RecACCT  As String * 25
    NOTICE   As String * 1
    BOX1     As Double
    BOX2     As Double
    BOX3     As Double
    BOX4     As Double
    BOX5     As Double
    BOX6     As Double
    BOX7     As Double
    BOX8     As Double
    BOX9     As String * 1
    BOX10    As Double
  'Changed to match new 1099's
    BOX13    As Double
    BOX14    As Double
    BOX15    As String * 30  'DOUBLE
    BOX16    As Double
    BOX17    As String * 12
    BOX18    As Double
'new for v205
    Void     As Integer   '1 for yes 0 for no
    Corrected As Integer  '1 for yes 0 for no
End Type

Type AP1099PayerRecType
    Fedid    As String * 30
    Name     As String * 30
    ADDR     As String * 30
    ADDR2    As String * 30
    CSZ      As String * 30
'new for V205
    BaseAmt  As Double
End Type

Type TranRecInfoType
    TranDate  As Integer
    TranRecNo As Long
End Type
Type MiscCodeRecType
    MiscCode As String * 7
    Description As String * 25
    GlAcctNumb As String * 14
    InActiveFlag As String * 1
    NotUsed As String * 17
End Type
'New CM type as of June 2004
Type CMTransRecType
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer           '1-Misc 24-Util 27-UtilDep 31-Tax 131-Newtax 41-License 141-NewBL 51-decal
    ''''''''''''''''''''''''''''''''''201-void Misc 224-void util 227-void dep 241-void lic 231-void tax
    ''''''''''''''''''''''''''''''''''251-void Decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Mod
    TransDetNum  As Long               'Holds Record Number of Transaction Det
    TransRevAmt(1 To 15) As Double
    TransOperNum As Long
    Trans2GL     As String * 1
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
'added charge 4 above and transvoid for new void payment procedure PS 4/14/04
    TransVoidNum As Long        'Voided trans link to record voided or void trans
    ChkByte      As String * 1
    TransPad     As String * 18
End Type
'also new as of June 2004
Type CMSetupType
    CMTOWNNAME   As String * 30
    GLInterface  As String * 1
    Pass4Voids   As String * 1  'Y -yes, N- no, F - full access only
    VoidPW       As String * 10
    Pass4Adj     As String * 1  'Y -yes, N- no, F - full access only
    AdjPW        As String * 10
    Filler       As String * 75  '128
End Type


Type SetUpAcctType
   RevName    As String * 15
   DebitAcct  As String * 14
   CreditAcct As String * 14
End Type

Type RevSetUpType
    RevName As String * 15
    UseDep   As String * 1
    USERATE  As String * 1
    TAXRATE  As Single
    UseMtr   As String * 1
    DistOr   As Integer
    Prorate  As String * 1
End Type

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    RecpPort        As String * 1
    RECPDEFT        As String * 1
    ESTREAD         As String * 1
    BANKDFT         As String * 1
    UseSeq          As String * 1
    BILLCYCL        As String * 1
    DefLook         As String * 1
    MethAcct        As String * 1      'new 02-14-97
    SkipInactive    As String * 1
    SkipSeparator   As String * 1
    Make99File      As String * 1
    LowRead         As Integer
    HighRead        As Integer
    HHDEVICE        As String * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 To 15) As RevSetUpType
    BillAcct(1 To 15) As SetUpAcctType
    PayAcct(1 To 15)  As SetUpAcctType
    DepAcct(1 To 15)  As SetUpAcctType
End Type
'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount

Type UBXferInfoType
  DAcctNo   As String * 14
  DebitAmt  As Double
  DRecNo    As Integer
  DTitle    As String * 30
  CAcctNo   As String * 14
  CreditAmt As Double
  CRecNo    As Integer
  CTitle    As String * 30
End Type

'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount
Type UBTransRecType
   TransDate              As Integer      '
   TransType              As Integer      '
   TransDesc              As String * 21  'may change
   Transamt               As Double       'total revenue amount
   RevAmt(1 To 15)        As Double       'Revenue amounts
   TaxAmt(1 To 15)        As Single       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 To 7)       As Integer
'*******************
   CurRead(1 To 7)        As Long         'Last/Current meter readings
   PrevRead(1 To 7)       As Long         'Previous readings
   ESTREAD(1 To 7)        As String * 1   'Y/N Flags for meter est's
   BillNumber             As Long         'Number on the bill that Printed
   ReadDate               As Integer
   BillDate               As Integer
   PastDueDate            As Integer
   DraftDate              As Integer      '
'111398
   ProRatePCT             As Integer
   ChkByte                As String * 1   'Added check byte
   EPPFlag                As String * 1   'Equal Payment Flag
   CustStatus             As String * 1   'Customer Status at Time of Transaction
'020199
   EPPTrans               As Long         'Pointer to Equal Pay trans
   PenAtBill              As Single       'Used to flag IRR Meter (Sunset)
'****************
   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check, 3=Cash/Check, 4=Charge
   OperatorNumber         As Integer      '
   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
   PrevTrans              As Long
   VoidFlag               As Integer       'Changed for wadesboro
   FromCMFlag             As Integer
   ActiveFlag             As Integer      'Valid transaction flag
   RunBalance             As Double
   CheckAmount            As Double
   CashAmount             As Double
   BillMsg                As String * 20
   ApplyDepFlag           As String * 1
   Posted2GL              As String * 1
   PrevDate               As Integer
   PenalFlag              As String * 1
   TaxExempt              As String * 1
   NONProfit              As String * 1
End Type
Type GLUBTempRecType
   Grabbatch              As String * 8  'this will be date and num of batch that day
   TransDate              As Integer      '
   TransType              As Integer      '
   TransDesc              As String * 21  'may change
   Transamt               As Double       'total revenue amount
   RevAmt(1 To 15)        As Double       'Revenue amounts
   TaxAmt(1 To 15)        As Single       'Tax Amounts
   CustStatus             As String * 1   'Customer Status at Time of Transaction
   CustName               As String * 35
   OperatorNumber         As Integer      '
   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
End Type

  Type GJXferRecType
    RevText    As String * 15
    BAcctInfo  As UBXferInfoType     'Billing Accounts
    PAcctInfo  As UBXferInfoType     'Payment Accounts
    DAcctInfo  As UBXferInfoType     'Deposit Accounts
  End Type
Type UBCustIndexRecType
  RecNum As Long
End Type
Type UBServiceAddressIndexType
  ServiceAddress  As String * 14
  RecNum   As Long
End Type

'Trans Types
Public Const TranUtilityBill = 1          '   1=Utility bill
Public Const TranLateCharge = 2           '   2=late charge
Public Const TranReconnectFee = 3         '   3=reconnect fee
Public Const TranBillPayment = 4          '   4=Bill Payment
Public Const TranAppliedDeposit = 5       '   5=Applied Deposit
Public Const TranPenaltyCharge = 6        '   6=Penalty Charge
Public Const TranDepositPayment = 7       '   7=Deposit Payment
Public Const TranDraftPayment = 8         '   8=Draft Payment
Public Const TranRefundDeposit = 9        '   9=Refund Deposit
Public Const TranBeginBalance = 10        '  10=Beginning Balance
Public Const TranUpwardAdjustment = 11    '  11=Bill Adjustments
Public Const TranDownwardAdjustment = 12  '  12=Bill Adjustments
'added this for new over payment adjustment on Aug 11,2003
Public Const TranOverPayAdjustment = 33   '  33=OverPayment Adjustment
Public Const TranDepCreditRemoval = 37    '  37= Deposit Credit Removal - Not Interfaced W/GL
Public Const TranMiscPayment = 99         '  99=Misc Payment
Public Const MaxRevsCnt = 15              '  Max num of Utility Revenues
Public Const TranDepPaymentVoid = 39      '  39= Deposit Payment Void  - same gl as deposit refund
Type ServicesType
    Ratecode As String * 4
    RMtrType As String * 1
End Type

Type FlatRateType
    FRDESC   As String * 18
    FRAMT    As Double
    FRFREQ   As String * 1
    REVSRC   As Integer
    NumMin   As Integer
End Type

Type RevDataType
    RevName    As String * 20
    Ratecode   As String * 4
    RevMtrType As String * 1
End Type

Type LocMeterType
    MtrNum    As String * 12
    MTRMulti  As Integer
    MTRType   As String * 1
    MtrUnit   As String * 1
    NumUser   As Integer
    InsDate   As Integer
    CurRead   As Long
    PrevRead  As Long
    CurDate   As Integer
    PastDate  As Integer       'hidden & protected
    ReadFlag  As String * 1    'hidden & protected
    AvgUse    As Long          'hidden & protected
    UseCnt    As Integer       'hidden & protected
    MtrIDNO   As String * 11
    MtrLat    As Double
    MtrLng    As Double
End Type

Type MonthlyPayType
    AMTOWED      As Double
    TotAmtPD     As Double
    PayAmt       As Double
    RevSource    As Integer
End Type

Type NewUBCustRecType
    Book          As String * 2
    SEQNUMB       As String * 6
    Status        As String * 1
    OpenDate      As Integer
    SEARCH        As String * 10
    CustName      As String * 35
    ADDR1         As String * 35
    ADDR2         As String * 35
    ServAddr      As String * 35
    City          As String * 18
    STATE         As String * 2
    ZIPCODE       As String * 10
    HPHONE        As String * 14
    WPHONE        As String * 14
    SOSEC         As String * 11
    DRVLIC        As String * 16
    CUSTTYPE      As String * 3
    Addr911       As String * 14
'051498 added bill to field. Removed 1 byte from 911 addr
    BillTo        As String * 1
'********************************************************
    BILLCOPY      As Integer
    POSTRTE       As String * 4
    BILLCYCL      As Integer
    ZONE          As String * 3
    Seq           As Long
'Page 2
    CASHONLY      As String * 1
    LATEFEE       As String * 1
    CUTOFFYN      As String * 1
    TAXEXPT       As String * 1
    SRCIT         As String * 1
    EPPFlag       As String * 1
'032299 Modified for Bank draft account type
'    EPPAMT        AS DOUBLE
'added GroupCoderec 2/1/05 for pointer to bookcode
    GroupCodeRec  As Integer
    Filler1       As String * 5
   ' Filler1       As String * 7
    USEDRAFT      As String * 1
    AcctType      As String * 1
'032299 Inserted account type
    BankName      As String * 34
    BANKLOC       As String * 30
    TRANSIT       As String * 9
    BankAcct      As String * 20
    BILLCMNT      As String * 25
    PAYCMNT       As String * 25
    PumpCode      As String * 4
    USERCODE1     As String * 4
    USERCODE2     As String * 2
    ProRatePCT    As Integer
    HHMSG1        As String * 20
    HHMSG2        As String * 20
    HHMSG3        As String * 20
'Page 3
    serv(1 To 15)      As ServicesType
    FlatRates(1 To 4)  As FlatRateType
'Page 4
    Monthly(1 To 2)    As MonthlyPayType
    MFEE1         As Double
    MFEE2         As Double
    LocMeters(1 To 7)  As LocMeterType
'END OF Quick Screen Form
    CustPin       As Long
    LastTrans     As Long
    CurrBalance   As Double
    PrevBalance   As Double
    CurrRevAmts(1 To 15) As Double   'includes the tax amount
    PrevRevAmts(1 To 15) As Double
    DepositAmt    As Double
    DelFlag       As Integer
    PreNoteFlag   As Integer
    WOLastTrans   As Long            'work order last trans pointer
    EstFlag       As String * 1
    MessageRec    As Long            ' Points to Message Record
    OldRec        As Long
    EPPLastTran   As Long
    NewNotes      As Integer
    DPCode        As String * 2
    FillPad       As String * 112
    ChkByte       As String * 1
End Type


'''New one for Bob's new BL
Type ARCustRecType
    CustNumb As String * 10
    SortName As String * 10
    BillName As String * 35
    ADDRESS1 As String * 35
    ADDRESS2 As String * 35
    City     As String * 20
    STATE    As String * 2
    ZIPCODE  As String * 10
    CustName As String * 35
    Contact  As String * 30
    BILLCAT1     As String * 5
    DESC1        As String * 35
    REV1         As Long
    Fee1         As Double 'latest charge
    FeeLicBal1   As Double
    FeeLicPay1   As Double
    BILLCAT2     As String * 5
    DESC2        As String * 35
    REV2         As Long
    Fee2         As Double
    FeeLicBal2   As Double
    FeeLicPay2   As Double
    BILLCAT3     As String * 5
    DESC3        As String * 35
    REV3         As Long
    Fee3         As Double
    FeeLicBal3   As Double
    FeeLicPay3   As Double
    BILLCAT4     As String * 5
    DESC4        As String * 35
    REV4         As Long
    Fee4         As Double
    FeeLicBal4   As Double
    FeeLicPay4   As Double
    BILLCAT5     As String * 5
    DESC5        As String * 35
    REV5         As Long
    Fee5         As Double
    FeeLicBal5   As Double
    FeeLicPay5   As Double
    IssuanceFee  As Double
    CustLocation As String * 1
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    VALID        As Integer
    Inactive     As String * 1    '"Y" if account is inactive
    Prorate      As Integer       'prorate percentage
    AcctBal      As Double
    IssueLicense As String * 1    'y/n
    Deleted      As String * 1    '(yY)=deleted, anything else isn't
    FirstTrans   As Long
    LastTrans    As Long
    LicBal       As Double
    FeeBal       As Double
    PenBal       As Double
    RoomtoGrow   As String * 136
    ChkByte      As String * 1
    IssuanceBal  As Double
    IssuancePay  As Double
    ServAdd      As String * 35
    SSNFID       As String * 15 'Paula...add this
End Type

Type ARNewCatCodeRecType
    CatCode    As String * 5    'Not Used in Version 8.5 work2 directory
    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
    CODEDESC   As String * 35
    Fee        As Single
    BaseAmt1   As Single
    Recpt1     As Double
    Percent1   As Single
    Maximum1   As Double
    BaseAmt2   As Single
    Recpt2     As Double
    Percent2   As Single
    Maximum2   As Double
    BaseAmt3   As Single
    Recpt3     As Double
    Percent3   As Single
    Maximum3   As Double
    BaseAmt4   As Single
    Recpt4     As Double
    Percent4   As Single
    Maximum4   As Double
    BaseAmt5   As Single
    Recpt5     As Double
    Percent5   As Single
    Maximum5   As Double
    REVGLNUM   As Long
    CashAcct   As Long
    ARGLACCT   As Long

    BaseAmt6   As Single
    Recpt6     As Double
    Percent6   As Single
    Maximum6   As Double
    RateStep   As Long
    Extra      As String * 36
End Type

Type ARCustIDXRecType
    IDXNAME     As String * 10
    IDXRECORD   As Integer
    ExtraRoom   As String * 52
End Type

Type ARTransRecType
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    TransType           As Integer
    TransDesc           As String * 35 '5
    CashAmount          As Double
    ChkAmount           As Double
    BalanceAfterTrans   As Double
    Posted2GL           As String * 1
    CatCodeRec1         As Long  '10         'Place to Grab G/L Acct #'s
    CatCodeRec2         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec3         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec4         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec5         As Long           'Place to Grab G/L Acct #'s
    CatLicAmt1          As Double '15
    CatLicAmt2          As Double
    CatLicAmt3          As Double
    CatLicAmt4          As Double
    CatLicAmt5          As Double
    CatLicBal1          As Double '25
    CatLicBal2          As Double
    CatLicBal3          As Double
    CatLicBal4          As Double
    CatLicBal5          As Double
    PenBal              As Double
    LicBal              As Double
    IssBal           As Double
    FeeAmt              As Double
    LicAmt              As Double
    PenAmt              As Double
    IssAmt              As Double
    ExtraRoom           As String * 8
    NextTrans           As Long
    DetailTransType     As Integer 'used for reading transaction types inside BL program for reports...not GL
    'Codes for General Ledger:
    '1 = all non-penalty charges; 2 = all payments; 6 = all penalty charges; 13 = adjust payment down
    '23 = adjust billing down; 24 = adjust billing up
    'Codes for internal Business License:
    '101 = Charge Penalty ; 110 = Charge Lic; 201 = Pay Penalty; 210 = Pay Lic; 211 = Pay Lic and Penalty; 301 = Adjust Down Pen; 310 = Adjust Down Lic
    '311 = Adjust Down Pen and Lic; '401 = Adjust Up Pen; 410 = Adjust Up Lic; 411 = Adjust Up Lic and Penalty
End Type


Type AREditPaymentRecType
    TranType        As Integer
    TranDate        As Integer
    CustNumber      As String * 10
    CustName        As String * 35
    Add1            As String * 35
    City            As String * 25
    STATE           As String * 2
    ZIPCODE         As String * 10
    Amount          As Double
    CASHCHK         As String * 9
    CASHAMT         As Double
    ChkAmt          As Double
    CREDITAM        As Double
    AMTPAID         As Double
    Change          As Double
    ISSUELIC        As String * 1
    SetFee          As String * 1
    ISSueFEE        As Double
    Desc            As String * 20
    LICDUE          As Double
    LICDUE1         As Double
    LICDUE2         As Double
    LICDUE3         As Double
    LICDUE4         As Double
    LICDUE5         As Double
    LICPAID         As Double
    LICPAID1        As Double
    LICPAID2        As Double
    LICPAID3        As Double
    LICPAID4        As Double
    LICPAID5        As Double
    TOTDUE          As Double
    TotPaid         As Double
    CatDesc1        As String * 35
    CatDesc2        As String * 35
    CatDesc3        As String * 35
    CatDesc4        As String * 35
    CatDesc5        As String * 35
    PENDUE          As Double
    PENPAID         As Double
    ISSDUE          As Double
    ISSPAID         As Double
End Type

Type CatCodeIdxType
  CatCodeRec As Integer
  CatCodeNum As String * 20
End Type


Type CustNameIdxType
   BillingName As String * 35
   CustRec As Integer
End Type
Type TownSetUpType
    TownName As String * 38 'allow for TOWN OF
    TownAdd1 As String * 30
    TownAdd2 As String * 30
    Contact As String * 30
    City As String * 30
    STATE As String * 2
    ZIPCODE As String * 10
    TownPhone As String * 14
    AppForm As Integer
    DLQNotice As Integer
    SpareSpace As String * 60
    AppAdd1 As String * 30
    AppCity As String * 30
    AppState As String * 2
    AppPhone As String * 14
    AppAdminName As String * 25
    AppAdminTitle As String * 25 '17
    AppBaseFee(1 To 10) As Double
    AppCentsPer(1 To 4) As Double
    AppGrsRcpts(1 To 4) As Double '29
    AppFirstDay As String * 7
    AppLastDay As String * 7
    AppTownOf As String * 38
    AppZip As String * 10
    AppPct As Double
    AppGrsPct As Double
    AppDenom As Integer
    AppNumer As Integer '37
    AppColFee As Double
    AppPayBy As Integer
    AppDiscPct As Double
    AppDiscMonth As String * 9
    AppDiscDay As Integer
    AppPenMonth As String * 9
    AppPenDay As Integer
    AppFiscMonth As String * 9
    AppFiscDay As Integer
    AppMayorCouncil As String * 25
    AppWholeMonth As Integer
    AppWholeDay As Integer '52
    AppRetailMonth As Integer
    AppRetailDay As Integer
    AppFinMonth As Integer
    AppFinDay As Integer
    AppContMonth As Integer
    AppContDay As Integer
    AppRepairMonth As Integer
    AppRepairDay As Integer
    AppStartMonth As String * 9
    AppStartDay As Integer
    AppLicRetMonth As String * 9
    AppLicRetDay As Integer
    AppAdoptDate As Integer
    AppCityOrd As String * 40
    AppYrUpDown(1 To 10) As String * 4
    DlqTownName As String * 38
    DlqAdd1 As String * 30 '68
    DlqCity As String * 30
    DlqState As String * 2
    DlqZip As String * 10
    DlqPhone As String * 14
    DlqPhone2 As String * 14
    DlqFax As String * 14
    DlqAdminName As String * 25
    DlqAdminTitle As String * 25
    DlqFirstDay As String * 9
    DlqLastDay As String * 9
    DlqFirstHour As String * 9
    DlqLastHour As String * 9
    DlqClerkName As String * 25
    DlqMayorCouncil As String * 25 '82
    LicNumPermYN  As String * 3
    UseAmtPctYN   As String * 3
    PENREVGLNUM   As Long
    PENRECGLNUM   As Long
    PENCASHACCT   As Long
    IssFee        As Double
    AcctMeth      As String * 1
    LaserLtr      As String * 1
    GL2Cats       As String * 1
End Type
'******************************end bl


'Type TaxMasterType      'Master Default Information in Setup
'  Name As String * 35
'  Add1 As String * 35
'  Add2 As String * 35
'  ADD3 As String * 35
'  TaxSt As String * 2
'  TaxForm As String * 20
'  CurRate As Single
'  PastRate As Single
'  PenRate As Single
'  RcptPort As Integer
'  AcctgMethod As String * 1
'  Padding As String * 253
'End Type

'Type RevSourceType
'  Principle1    As Double                 'Va Personal Prop
'  Principle2    As Double    'For Va Only     Mach/Tools
'  Principle3    As Double    'For Va Only     Merch Cap
'  Principle4    As Double    'For Va Only     Farm Equip
'  Principle5    As Double    'For Va Only     Mobile Homes
'  Interest      As Double
'  Penalty       As Double
'  Collection    As Double
'  Future1       As Double
'  Future2       As Double
'  Principle1Pd  As Double
'  Principle2Pd  As Double    'For Va Only
'  Principle3Pd  As Double    'For Va Only
'  Principle4Pd  As Double    'For Va Only
'  Principle5Pd  As Double    'For Va Only
'  InterestPd    As Double
'  PenaltyPd     As Double
'  CollectionPd  As Double
'  Future1Pd     As Double
'  Future2Pd     As Double
'End Type

'Type TaxTransactionType
'  TransDate    As Integer          'Transaction Date
'  TaxYear      As Integer          'Must Contain Full 4 digit Tax Year Here
'  TranType     As Integer          '1=Bill 2=Payment 3=Release 4=Interest
'                                   '5=Penalty 8=Collection/Ad Cost Billing
'                                   '7=Adjustment
'  BillType     As String * 1       'R=Real P=Personal Property C=Combined (NC/GA)
'  Amount       As Double           'Total Transaction Amount
'  Revenue      As RevSourceType    'See Revenue Source Type File above
'  Description  As String * 30      'Description of Transaction
'  Posted2GL    As String * 1       'I/F to G/L Yes or No
'  CustomerRec  As Long             'Pointer Back to Customer Record
'  LastTrans    As Long             'Points to Previous Trans in History actually Previous pointer
'  BelongTo     As Long             'Points to Record of Bill this Transaction belongs to:'will be 0 for Bill
'  DMVSubmitted As String * 1       'Y if Sent to DMV
'  DMVBatch     As Integer          'Records which batch contained the DMV Tranmission
'  Altered      As Integer          'Flag <> 0  If TR altered at any time
'  Padding      As String * 123     'Allow for Future Expansion
'End Type

Type DCCatCodeRecType
    CatCode    As String * 3
    CODEDESC   As String * 35
    APPNUMB    As Integer
    BILLCODE   As Integer
    REVGLNUM   As String * 14
    CashAcct   As String * 14
    Fee        As Single
    Extra      As String * 54
End Type
Type DCTransRecType
    CustomerNumber As String * 10
    TransDate As Integer
    TransAmount As Double
    TransType As Integer     '1-charge,2-pay,3-voidchrg,4-voidpay
    TRVinDesc As String * 40
    CashAmount As Double
    ChkAmount As Double
    BalanceAfterTrans As Double
    makemodel As String * 25
    StateTag As String * 35
    ExpireDate As Integer
    Sticker As String * 12
    NextTrans As Long
    OperNum   As Long
    GLInterfaced  As String * 1
    DecalCat As String * 5
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
''added tendertype and 3,4 transtypes above and  chkbyte to prevent reconverting PS 7/8/05
    VoidFlag As String * 1   'Y if voided
    ChkByte  As String * 1   'this is chr$(1)
    ExtraDesc As String * 26   'added extra
    VehRecord As Long
    ExtraRoom As String * 48
End Type
'old
''Type DCTransRecType
''    CustomerNumber As String * 10
''    TransDate As Integer
''    TransAmount As Double
''    TransType As Integer
''    TransDesc As String * 35
''    CashAmount As Double
''    ChkAmount As Double
''    BalanceAfterTrans As Double
''    makemodel As String * 25
''    StateTag As String * 25
''    ExpireDate As Integer
''    Sticker As String * 12
''    NextTrans As Long
''    OperNum   As Long
''    GLInterfaced  As String * 1
''    DecalCat As String * 5
''    ExtraRoom As String * 97
''End Type

Type EPTransRecType
    EPMonth  As String * 2
    EPDay    As String * 2
    EPYear   As String * 4
    EPAcct   As String * 14
    EPDebit  As String * 10
    EPCredit As String * 10
    EPDesc   As String * 32
    epextra  As String * 2 'this is for carriage return they had to include#*&@($&($&@(*#
End Type
'not used any longer
'Type IFRecType
'   AcctNum As String * 9      '9 AS tranacct$
'   TRDATE As String * 8       '8 AS trandate$
'   Desc As String * 20        '20 AS trandesc$
'   CrAmt As Double            '8 AS cramt$
'   DrAmt As Double            '8 AS dramt$
'   Ref As String * 8          '8 AS detail$
'   Src As String * 8          '8 AS source$
'   Filler As String * 14      '4 AS nexttr$
'   Posted As Integer
'End Type
'Type APPOType
'    DELFLAG  As Integer
'    Vendor   As String * 10
'    VENDNAME As String * 20
'    PONum    As String * 15
'    POAMT   As Double
'    PODATE   As Integer
'    Dist(1 To 36) As DistType
'End Type

'--Type for Invoice Edit
'Type APInvType
'    DELFLAG   As Integer
'    Vendor    As String * 10
'    VENDNAME  As String * 20
'    VRecNum   As String * 2
'    PONum     As String * 18
'    INVNUM    As String * 25
'    INVAMT    As Double
'    PAYCODE   As String * 1
'    INVDATE   As Integer
'    DUEDATE   As Integer
'    DISTDATE  As Integer
'    Dist(1 To 36) As DistType
'    POAPLRecNum  As String * 4        'this is a long integer pointer
'    POFLAG    As Integer           'Flag to indicate an active PO
'End Type

'Changed Vendor file structure for DBA field for Holly Springs
'for 1099's to show Doing Business As line.
'Type VendorRecTypeOld
'   VIN        As Long
'   vnum       As String * 10
'   VNAME      As String * 30
'   Addr1      As String * 30
'   Addr2      As String * 30
'   City       As String * 22
'   State      As String * 2
'   Zip        As String * 10
'   PaytoName  As String * 30
'   PaytoAddr  As String * 30
'   PaytoAddr2 As String * 30
'   PayToCity  As String * 22
'   PaytoState As String * 2
'   PaytoZip   As String * 10
'   Phone      As String * 14
'   VTerms     As Integer
'   pad2 As String * 5
'   Fedid      As String * 12
'   CoCode     As String * 3
'   StCode     As String * 2
'   YTDPay     As Double
'   Get1099  As String * 1
'   CurrBal    As Double
'   FrstTran   As Long
'   LastTran   As Long
'   FrstPO     As Long
'   LastPO     As Long
'   DefDist    As Long
'   DelFlag    As Integer
'   'Pad        AS STRING * 45
'   Contact    As String * 30
'   Fax        As String * 14
'   Pad As String * 1
'End Type
'--A/P Ledger
'Type APLedgerRecType  'version 8.0
'   VIN        As Integer
'   VendorCode As String * 10
'   VRecNum    As Integer
'   TRDATE     As Integer
'   GLDistDate As Integer
'   DUEDATE    As Integer
'   TrCode     As Integer
'   DOCNum     As String * 25
'   PONum      As String * 20
'   PAYCODE    As Integer
'   PrintCode  As Integer
'   PDCheckNum   As Long
'   PDCheckDate  As Integer
'   MiscCode     As String * 23
'   Amt          As Double
'   FrstDist     As Long
'   LastDist   As Long
'   NextTrans  As Long
'End Type

'Type APLedgerRecType0
'  VendorCode As Integer                 '2 AS vennum$
'  TRDATE As String * 8             '8 AS INVDATE$
'  DOCREF As String * 25           '25 AS invnum$
'  PONum As String * 10             '10 AS ponum$
'  Amt As Double                    '8 AS amount$
'  PAYCODE As Integer               '2 AS PAYCODE$
'  FirstDist As Single              '4 AS fdist$
'  LastDist As Single               '4 AS ldist$
'  NextTrans As Single              '4 AS ndata$
'  CoTaxCode As String * 3          '3 AS tctycode$
'  StTaxCode As String * 3          '3 AS tstcode$
'  CoTaxAmt As Single               '4 AS ctaxamt$
'  StTaxAmt As Single               '4 AS staxamt$
'  Fill As String * 15
'End Type
'Type MiscCodeRecType
'    MiscCode As String * 7
'    Description As String * 25
'    GlAcctNumb As String * 14
'    NotUsed As String * 18
'End Type
'
'Type CMTransRecType
'    TransDate    As Integer
'    TransAmount  As Double
'    TransCash    As Double
'    TransCheck   As Double
'    TransAmtOwed As Double
'    TransDesc    As String * 25
'    TransSource  As Integer            '1-Misc 2-Utility 3-Tax 4-License
'                                       '5-decal
'    TransName    As String * 25
'    TransAcctNum As Long               'Holds Master Acct Record Number in Module
'    TransDetNum  As Long               'Holds Record Number of Transaction Detail in Module
'    TransRevAmt(1 To 15) As Double
'    TransOperNum As Long
'    Trans2GL      As String * 1
'    TransPad     As String * 25
'End Type
'Business license stuff**********************
'OLD BL
''Type ARCatCodeRecType
''    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
''    CODEDESC   As String * 35
''    Fee     As Single
''    REVGLNUM   As String * 14
''    CashAcct   As String * 14
''    ARGLACCT   As String * 14
''    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
''    Percent    As Single
''    Maximum    As Double
''    Extra      As String * 157
''End Type
''
''Type ARNewCatCodeRecType
''    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
''    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
''    CODEDESC   As String * 35
''    Fee        As Single
''    BaseAmt1   As Single
''    Recpt1     As Double
''    Percent1   As Single
''    Maximum1   As Double
''    BaseAmt2   As Single
''    Recpt2     As Double
''    Percent2   As Single
''    Maximum2   As Double
''    BaseAmt3   As Single
''    Recpt3     As Double
''    Percent3   As Single
''    Maximum3   As Double
''    BaseAmt4   As Single
''    Recpt4     As Double
''    Percent4   As Single
''    Maximum4   As Double
''    BaseAmt5   As Single
''    Recpt5     As Double
''    Percent5   As Single
''    Maximum5   As Double
''    REVGLNUM   As String * 14
''    CashAcct   As String * 14
''    ARGLACCT   As String * 14
''    Extra      As String * 64
''End Type
''
''Type ARTransRecType
''    CustomerNumber      As String * 10
''    TransDate           As Integer
''    TransAmount         As Double
''    TransType           As Integer
''    TransDesc           As String * 35
''    CashAmount          As Double
''    ChkAmount           As Double
''    BalanceAfterTrans   As Double
''    Posted2GL           As String * 1
''    CatCodeRec          As Integer           'Place to Grab G/L Acct #'s
''    ExtraRoom           As String * 40
''    NextTrans           As Long
''End Type
'********************************************************
'Type TAXGLAcctRecType
'  TaxYear       As Integer        'protected
'  TaxDBAcct     As String * 14
'  TaxCRAcct     As String * 14
'  IntDBAcct     As String * 14
'  IntCRAcct     As String * 14
'  AdvDBAcct     As String * 14
'  AdvCRAcct     As String * 14
'  Fill1         As String * 1     'protected
'End Type

'Type TaxAcctsType
'  TaxAcct(1 To 31)   As TAXGLAcctRecType
'  '1980 thru 2010  Inclusive
'End Type

'New Tax Stuff for v2.05
Type WinTAXGLAcctRecType
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
Type TaxAcctsType
  TaxAcct(1 To 51) As WinTAXGLAcctRecType
End Type

Type TaxMasterType      'Master Default Information in Setup
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
  Padding As String * 189
End Type
Type WinRevSourceType
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
  Pad           As String * 80
End Type

Type TaxTransactionType
  TransDate    As Integer          'Transaction Date
  TaxYear      As Integer          'Must Contain Full 4 digit Tax Year Here
  TranType     As Integer          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
                                   '10=DwnAdjPay 11=UpAdjPay
                                   '22=PrePayment 23=Refund Prepayment added 3-25-03
  BillType     As String * 1       'R=Real P=Personal Property C=Combined (NC/
  Amount       As Double           'Total Transaction Amount
  Revenue      As WinRevSourceType    'See Revenue Source Type File above
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
')(*)(*)*)*)(*)(*)(*)(*)(*)(*)()(*)
'New VATAX V2.05
Type WinRVATAXGLAcctRecType
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
Type WinPVATAXGLAcctRecType
  TaxYear       As Integer        'protected
  PersDBAcct    As String * 14
  PersCRAcct    As String * 14
  MTDBAcct      As String * 14
  MTCRAcct      As String * 14
  MCDBAcct      As String * 14
  MCCRAcct      As String * 14
  Fill1         As String * 1     'protected
  FEDBAcct      As String * 14
  FECRAcct      As String * 14
  MHDBAcct      As String * 14
  MHCRAcct      As String * 14
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

Type TaxRVAAcctsType
  TaxAcct(1 To 51) As WinRVATAXGLAcctRecType
End Type

Type TaxPVAAcctsType
  TaxAcct(1 To 51) As WinPVATAXGLAcctRecType
End Type
Type PINRecType
  PIN As Long
End Type
Type TaxVAMasterType      'Master Default Information in Setup
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

Type TaxVATransactionType
  TransDate    As Integer          'Transaction Date
  TaxYear      As Integer          'Must Contain Full 4 digit Tax Year Here
  TranType     As Integer          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
                                   '10=DwnAdjPay 11=UpAdjPay
                                   '22=PrePayment 23=Refund Prepayment added 3-25-03
  BillType     As String * 1       'R=Real P=Personal Property C=Combined (NC/
  Amount       As Double           'Total Transaction Amount
  Revenue      As WinRevSourceType    'See Revenue Source Type File above
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

