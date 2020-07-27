'G/L File Type Structures

TYPE GLSetupRecType                 'still under const.
   UserName    AS STRING * 30
   TotAcctLen  AS INTEGER
   FundLen     AS INTEGER
   AcctLen     AS INTEGER
   DetLen      AS INTEGER
   CashAcct    AS STRING * 14
   APAcct      AS STRING * 14
   EncAcct     AS STRING * 14
   FBAcct      AS STRING * 14
   FYBeg       AS INTEGER
   FYEnd       AS INTEGER
   NYBeg       AS INTEGER
   NYEnd       AS INTEGER
   CDCash      AS STRING * 14  'new
   CDDue       AS STRING * 14
   CDActive    AS STRING * 1
   CRCashAcct  AS STRING * 14
   DeptCode    AS STRING * 1
   LPDate      AS INTEGER
   HPDate      AS INTEGER
   CDCashAcct  AS STRING * 14
   CDsbCash    AS STRING * 14
END TYPE

TYPE GLFundIndexType                'Fund Index : 16 bytes
   FundNum     AS STRING * 4        'Fund Number
   RecNum      AS INTEGER           'Pointer to record
END TYPE

TYPE GLFundRecType                  'Fund Record Type: 64 bytes
   Deleted     AS INTEGER           'Deleted Flag
   FundNum     AS STRING * 4        'Fund Code
   Title       AS STRING * 30       'Fund Title
   Res         AS STRING * 28       'Reserve for future needs
END TYPE

TYPE GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     AS STRING * 14       'Formatted account Number string
   RecNum      AS INTEGER           'Pointer to record
END TYPE

TYPE GLAcctRecType                  'Account Record Type: ? bytes
   Deleted     AS INTEGER           'Active Account Flag
   Num         AS STRING * 14       'Formatted Account Number
   Title       AS STRING * 30       'Account Description
   Typ         AS STRING * 1        'Account Type
   FrstTran    AS LONG              'Pointer to First Trans
   LastTran    AS LONG              'Pointer to Last Trans
   PYAct       AS DOUBLE            'Prior Year Actual
   BegBal      AS DOUBLE            'Beginning Balance
   Bgt         AS DOUBLE            'Budget Amount
   Bal         AS DOUBLE            'Running Balance
   Encumb      AS DOUBLE            'Encumbered Amount
   MTD         AS DOUBLE            'Month to Date Bal (calc as needed)
   YTD         AS DOUBLE            'Year to Date Bal (calc as needed)
   NYEst       AS DOUBLE            'Bgt Estimate
   NYReq       AS DOUBLE            'New Year Bgt Requested
   NYRec       AS DOUBLE            'New Year Bgt Recommended
   NYApp       AS DOUBLE            'New Year Bgt Approved
   FrstBTran   AS INTEGER           'Pointer to First Budget Trans
   LastBTran   AS INTEGER           'Pointer to Last Budget Trans
   FrstPTran   AS INTEGER           'Pointer to First Budget Trans
   LastPTran   AS INTEGER           'Pointer to Last Budget Trans
   'Res         AS STRING * 25       'Reserved for future needs
   Work        AS DOUBLE            'Temp added 08/17/96 for closeout
   Res         AS STRING * 17
   Marked      AS INTEGER           '
END TYPE

TYPE GLDeptIndexType                'Dept Index
   DeptNum     AS STRING * 8        'Dept Number
   RecNum      AS INTEGER           'Pointer to record
END TYPE

TYPE GLDeptRecType                  'Dept Record Type
   Deleted     AS INTEGER           'Deleted Flag
   DeptNum     AS STRING * 8        'Fund Code
   Title       AS STRING * 30       'Fund Title
   Res         AS STRING * 20       'Reserve for future needs
END TYPE

TYPE GLTransRecType                 'Transaction Record: 96 bytes
   AcctRec     AS INTEGER           'Pointer to Acct Record
   AcctNum     AS STRING * 14       'Formatted Acct Number string
   TrDate      AS INTEGER           'Date2Num function
   Desc        AS STRING * 20       'Transaction Description
   Ref         AS STRING * 8        'Document Reference
   DrAmt       AS DOUBLE            'Debit Amount
   CrAmt       AS DOUBLE            'Credit Amount
   Src         AS STRING * 8        'Module Source Code
   NextTran    AS LONG              'Pointer to Next Trans
   Res         AS STRING * 20       'Reserved for future needs
   Marked      AS INTEGER
END TYPE

TYPE TrEditRecType                  'Experimental GJ edit record:
   Deleted     AS INTEGER           'Deleted transaction flag
   Posted      AS INTEGER           'Posted flag
   AcctRec     AS INTEGER           'Pointer to acct rec
   AcctNum     AS STRING * 14       'Formatted Acct number string
   AcctName    AS STRING * 30       'Account Title
   TrDate      AS INTEGER           'Date2Num
   DrAmt       AS DOUBLE            'Transaction Debit Amount
   CrAmt       AS DOUBLE            'Transaction Credit Amount
   EType       AS STRING * 1        'Entry Type (Debit/Credit)
   Desc        AS STRING * 20       'Transaction Description
   Ref         AS STRING * 8        'Document Reference #
   Src         AS STRING * 8        'Module Source Code
   'Res         AS STRING *          'Reserve for future needs
END TYPE

TYPE TrSortType                     'Used for sorting trans in history rpt
   TrDate     AS INTEGER            'Transaction Date
   Record     AS LONG               'Pointer to transaction record
END TYPE

TYPE TrSortType1                    'Used for sorting trans in history rpt
   TrDate     AS STRING * 12             'Transaction Date
   Record     AS LONG               'Pointer to transaction record
END TYPE


TYPE IFRecType
   AcctNum AS STRING * 9      '9 AS tranacct$
   TrDate AS STRING * 8       '8 AS trandate$
   Desc AS STRING * 20        '20 AS trandesc$
   CrAmt AS DOUBLE            '8 AS cramt$
   DrAmt AS DOUBLE            '8 AS dramt$
   Ref AS STRING * 8          '8 AS detail$
   Src AS STRING * 8          '8 AS source$
   filler AS STRING * 14      '4 AS nexttr$
   Posted AS INTEGER
END TYPE

TYPE GLFBAdjRecType
   AcctNum  AS STRING * 16
   AdjAmt AS DOUBLE
END TYPE

TYPE GLFundCloseRecType
   FundNum AS STRING * 3
END TYPE

TYPE BankRecType   '128 bytes
   Deleted AS INTEGER
   BankNum AS INTEGER
   BankName AS STRING * 25
   BankAcct AS STRING * 25
   GLAcct AS STRING * 25
   Pad AS STRING * 49
END TYPE

TYPE GLSumSetupRecType                 'still under const.
   Beg1  AS STRING * 6
   End1  AS STRING * 6
   Desc1 AS STRING * 30
   Beg2  AS STRING * 6
   End2  AS STRING * 6
   Desc2 AS STRING * 30
   Beg3  AS STRING * 6
   End3  AS STRING * 6
   Desc3 AS STRING * 30
   Beg4  AS STRING * 6
   End4  AS STRING * 6
   Desc4 AS STRING * 30
   Beg5  AS STRING * 6
   End5  AS STRING * 6
   Desc5 AS STRING * 30
   Beg6  AS STRING * 6
   End6  AS STRING * 6
   Desc6 AS STRING * 30
   Beg7  AS STRING * 6
   End7  AS STRING * 6
   Desc7 AS STRING * 30
   Beg8  AS STRING * 6
   End8  AS STRING * 6
   Desc8 AS STRING * 30
   Beg9  AS STRING * 6
   End9  AS STRING * 6
   Desc9 AS STRING * 30
   Pad   AS STRING * 75
END TYPE

TYPE InvTaxAcct
    AcctNo  AS STRING * 16
    TaxAmt  AS DOUBLE
END TYPE

'TYPE InvoiceTaxInfoType
'    InvTax(1 TO 2)  AS InvTaxAcct
'    AutoDist        AS STRING * 1
'END TYPE

