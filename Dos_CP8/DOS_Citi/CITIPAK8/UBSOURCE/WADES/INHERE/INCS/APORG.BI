'AP.BI

'--Vendor Index
TYPE VendorIdxRecType
   VendorCode AS STRING * 10
   RecNum AS INTEGER
END TYPE

'--Vendor Record
TYPE VendorRecType
   VIN        AS LONG
   VNum       AS STRING * 10
   VName      AS STRING * 30
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
   Phone      AS STRING * 14
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
    InvNum    AS STRING * 25
    INVAMT    AS DOUBLE
    PayCode   AS STRING * 1
    InvDate   AS INTEGER
    DueDate   AS INTEGER
    DISTDATE  AS INTEGER
    Dist(1 TO 36) AS DistType
    POAPLRecNum  AS STRING * 4        'this is a long integer pointer
    POFLAG    AS INTEGER           'Flag to indicate an active PO
END TYPE

'--A/P Ledger
TYPE APLedgerRecType
   VIN        AS INTEGER
   VendorCode AS STRING * 10
   VRecNum    AS INTEGER
   TRDate     AS INTEGER
   GLDistDate AS INTEGER
   DueDate    AS INTEGER
   TrCode     AS INTEGER
   DOCNum     AS STRING * 25
   PONUM      AS STRING * 20
   PayCode    AS INTEGER
   PrintCode  AS INTEGER
   PDCheckNum   AS LONG
   PDCheckDate  AS INTEGER
   MiscCode     AS STRING * 23
   Amt          AS DOUBLE
   FrstDist     AS LONG
   lastDist   AS LONG
   NextTrans  AS LONG
END TYPE

TYPE oAPLedgerRecType
   VIN        AS INTEGER
   VendorCode AS STRING * 10
   VRecNum    AS INTEGER
   InvDate    AS INTEGER
   GLDistDate AS INTEGER
   DueDate    AS INTEGER
   TrCode     AS INTEGER
   InvNum     AS STRING * 25
   PONUM      AS STRING * 20
   PayCode    AS INTEGER
   PrintCode  AS INTEGER
   MiscCode   AS STRING * 29
   Amt        AS DOUBLE
   FrstDist   AS LONG
   lastDist   AS LONG
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

'--Vendor Default Distribution
'need to revise
TYPE VendorDefDistRecType
   DefAcct    AS STRING * 16
   DefPct     AS SINGLE
END TYPE

