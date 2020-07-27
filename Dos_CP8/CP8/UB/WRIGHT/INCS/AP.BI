'AP.BI
'--Vendor Index
TYPE VendorIdxRecType
   VendorCode AS STRING * 10
   RecNum AS INTEGER
END TYPE

'--Vendor Record
TYPE oVendorRecType
   VIN        AS LONG
   VNum       AS STRING * 10
   VNAME      AS STRING * 30
   Addr1      AS STRING * 30
   Addr2      AS STRING * 30
   City       AS STRING * 22
   State      AS STRING * 2
   Zip        AS STRING * 10
   PayToName  AS STRING * 30
   PayToAddr  AS STRING * 30
   PaytoAddr2 AS STRING * 30
   PayToCity  AS STRING * 22
   PaytoState AS STRING * 2
   PaytoZip   AS STRING * 10
   phone      AS STRING * 14
   FedID      AS STRING * 12
   CoCode     AS STRING * 3
   StCode     AS STRING * 2
   YTDPurch   AS DOUBLE
   YTDPay     AS DOUBLE
   CurrBal    AS DOUBLE
   FrstTran   AS LONG
   LastTran   AS LONG
   FrstPO     AS LONG
   LastPO     AS LONG
   DefDist    AS LONG
   DELFLAG    AS INTEGER
   Pad        AS STRING * 45
END TYPE

TYPE VendorRecType
   VIN        AS LONG
   VNum       AS STRING * 10
   VNAME      AS STRING * 30
   Addr1      AS STRING * 30
   Addr2      AS STRING * 30
   City       AS STRING * 22
   State      AS STRING * 2
   Zip        AS STRING * 10
   PayToName  AS STRING * 30
   PayToAddr  AS STRING * 30
   PaytoAddr2 AS STRING * 30
   PayToCity  AS STRING * 22
   PaytoState AS STRING * 2
   PaytoZip   AS STRING * 10
   phone      AS STRING * 14
   VTERMS     AS INTEGER
   pad2 AS STRING * 5
   FedID      AS STRING * 12
   CoCode     AS STRING * 3
   StCode     AS STRING * 2
   YTDPay     AS DOUBLE
   Get1099  AS STRING * 1
   CurrBal    AS DOUBLE
   FrstTran   AS LONG
   LastTran   AS LONG
   FrstPO     AS LONG
   LastPO     AS LONG
   DefDist    AS LONG
   DELFLAG    AS INTEGER
   'Pad        AS STRING * 45
   Contact    AS STRING * 30
   Fax        AS STRING * 14
   Pad AS STRING * 1
END TYPE

'--Distribution type for invoice edit
TYPE DistType
   DACN       AS STRING * 16
   DACNM      AS STRING * 20
   DACREC     AS STRING * 2      'this is an integer rec number
   'DACREC     AS INTEGER
   DAMT       AS DOUBLE
END TYPE

TYPE APPOType
    DELFLAG  AS INTEGER
    VENDOR   AS STRING * 10
    VENDNAME AS STRING * 20
    PONUM    AS STRING * 15
    POAMT   AS DOUBLE
    PODATE   AS INTEGER
    Dist(1 TO 36) AS DistType
END TYPE

'--Type for Invoice Edit
TYPE APInvType
    DELFLAG   AS INTEGER
    VENDOR    AS STRING * 10
    VENDNAME  AS STRING * 20
    VRecNum   AS STRING * 2
    PONUM     AS STRING * 18
    INVNUM    AS STRING * 25
    INVAMT    AS DOUBLE
    PAYCODE   AS STRING * 1
    INVDATE   AS INTEGER
    DUEDATE   AS INTEGER
    DISTDATE  AS INTEGER
    Dist(1 TO 36) AS DistType
    POAPLRecNum  AS STRING * 4        'this is a long integer pointer
    POFLAG    AS INTEGER           'Flag to indicate an active PO
END TYPE

'--Type for Invoice Edit
TYPE APInv81Type
    DELFLAG  AS INTEGER           '1  protected
    VENDOR   AS STRING * 10       '2
    VENDNAME AS STRING * 20       '3
    VRecNum  AS STRING * 2        '4  protected
    INVNUM   AS STRING * 25       '5
    PONUM    AS STRING * 18       '6
    INVAMT   AS DOUBLE            '7
    PAYCODE  AS STRING * 1        '8  protected
    INVDESC  AS STRING * 33       '9
    INVDATE  AS INTEGER           '10
    DUEDATE  AS INTEGER           '11
    DISTDATE AS INTEGER           '12
    PSLFlag  AS STRING * 1        '13
    Get1099  AS STRING * 1        '14
    Dist(1 TO 24) AS DistType     '
    POAPLRecNum  AS STRING * 4     'this is a long integer pointer
    POFLAG    AS INTEGER           'Flag to indicate an active PO
'Added this for multi PO's
'    PORecs(1 TO 6)  AS STRING * 4
END TYPE


'--Type for Invoice Edit
TYPE APInv85Type   'this is actually inv85
    DELFLAG  AS INTEGER           '1  protected
    VENDOR   AS STRING * 10       '2
    VENDNAME AS STRING * 20       '3
    VRecNum  AS STRING * 2        '4  protected
    INVNUM   AS STRING * 25       '5
    PONUM    AS STRING * 18       '6
    INVAMT   AS DOUBLE            '7
    INVDESC  AS STRING * 33       '9
    TaxYN    AS STRING * 1
    PAYCODE  AS STRING * 1        '8  protected
    INVDATE  AS INTEGER           '10
    DUEDATE  AS INTEGER           '11
    DISTDATE AS INTEGER           '12
    PSLFlag  AS STRING * 1        '13
    Get1099  AS STRING * 1        '14
    STaxAmt  AS DOUBLE
    CTaxAmt  AS DOUBLE
    GrandTot AS DOUBLE
    Dist(1 TO 24) AS DistType     '
    POAPLRecNum  AS STRING * 4     'this is a long integer pointer
    POFLAG    AS INTEGER           'Flag to indicate an active PO
'Added this for multi PO's
    PORecs(1 TO 6)  AS STRING * 4
END TYPE

TYPE APLedger81RecType             'version for Troy's paid supply list
   VIN          AS INTEGER
   VendorCode   AS STRING * 10
   VRecNum      AS INTEGER
   TrDate       AS INTEGER
   GLDistDate   AS INTEGER
   DUEDATE      AS INTEGER
   TrCode       AS INTEGER      '1=Invoice, 4=PO, 3=Check, -3=Void Check, -4=Cleared PO
   DOCNum       AS STRING * 25
   PONUM        AS STRING * 20
   PAYCODE      AS INTEGER
   PrintCode    AS INTEGER
   PDCheckNum   AS LONG
   PDCheckDate  AS INTEGER
   Comment      AS STRING * 31          'Reduced to 31 bytes to allow for dept to be stored (wrightsville beach)
   DeptNumb     AS LONG
   PSLFlag      AS STRING * 1
   Get1099      AS STRING * 1
   Amt          AS DOUBLE
   FrstDist     AS LONG
   LastDist     AS LONG
   NextTrans    AS LONG
   TaxAmt       AS DOUBLE
   Pad           AS STRING * 2
END TYPE

'--A/P Ledger
TYPE APLedgerRecType  'version 8.0
   VIN        AS INTEGER
   VendorCode AS STRING * 10
   VRecNum    AS INTEGER
   TrDate     AS INTEGER
   GLDistDate AS INTEGER
   DUEDATE    AS INTEGER
   TrCode     AS INTEGER
   DOCNum     AS STRING * 25
   PONUM      AS STRING * 20
   PAYCODE    AS INTEGER
   PrintCode  AS INTEGER
   PDCheckNum   AS LONG
   PDCheckDate  AS INTEGER
   MiscCode     AS STRING * 23
   Amt          AS DOUBLE
   FrstDist     AS LONG
   LastDist   AS LONG
   NextTrans  AS LONG
END TYPE

TYPE APLedger7RecType
  VendorCode AS INTEGER                 '2 AS vennum$
  TrDate AS STRING * 8             '8 AS INVDATE$
  DOCRef AS STRING * 25           '25 AS invnum$
  PONUM AS STRING * 10             '10 AS ponum$
  Amt AS DOUBLE                    '8 AS amount$
  PAYCODE AS INTEGER               '2 AS PAYCODE$
  FirstDist AS SINGLE              '4 AS fdist$
  LastDist AS SINGLE               '4 AS ldist$
  NextTrans AS SINGLE              '4 AS ndata$
  CoTaxCode AS STRING * 3          '3 AS tctycode$
  StTaxCode AS STRING * 3          '3 AS tstcode$
  CoTaxAmt AS SINGLE               '4 AS ctaxamt$
  StTaxAmt AS SINGLE               '4 AS staxamt$
  Fill AS STRING * 15
END TYPE

TYPE oAPLedgerRecType
   VIN        AS INTEGER
   VendorCode AS STRING * 10
   VRecNum    AS INTEGER
   INVDATE    AS INTEGER
   GLDistDate AS INTEGER
   DUEDATE    AS INTEGER
   TrCode     AS INTEGER
   INVNUM     AS STRING * 25
   PONUM      AS STRING * 20
   PAYCODE    AS INTEGER
   PrintCode  AS INTEGER
   MiscCode   AS STRING * 29
   Amt        AS DOUBLE
   FrstDist   AS LONG
   LastDist   AS LONG
   NextTrans  AS LONG
END TYPE

'--A/P Ledger Accounting Distributions
TYPE APDistRecType
   APLedgerRec AS LONG
   DistAcctRec AS INTEGER
   'DistAcctRec AS STRING * 2
   DistAcctNum AS STRING * 16
   DistAmt     AS DOUBLE
   NextDist    AS LONG
END TYPE

'--Distributions Summary work array type
TYPE DistSumType
   DistAcctNum AS STRING * 16
   AcctTitle   AS STRING * 20
   DistAmt     AS DOUBLE
END TYPE

'--Vendor Default Distribution
'need to revise

TYPE DefDistAcctsRecType
   DefAcct      AS STRING * 16
   DefAcctName  AS STRING * 20
   DefPct       AS SINGLE

END TYPE

TYPE VendorDefDistRecType
   VRecNum    AS INTEGER
   DefDist(1 TO 8)  AS DefDistAcctsRecType
END TYPE


