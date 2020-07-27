TYPE PINSearchType
  PIN   AS STRING * 20
  Cust  AS LONG
END TYPE

TYPE PropertyRecType
    REALPIN  AS STRING * 20
    PROPDATE AS INTEGER
    GISPOS   AS STRING * 20
    MAP      AS STRING * 6
    BLOCK    AS STRING * 6
    LOTNUMB  AS STRING * 6
    LOTACRE  AS STRING * 1
    PROPSIZE AS DOUBLE
    PROPDISC AS STRING * 1
    LATELIST AS STRING * 1
    MORTCODE AS STRING * 8
    PROPVALU AS DOUBLE
    EXMPSENI AS DOUBLE 'bldgval
    EXMPOTHR AS DOUBLE
    PROPNOT1 AS STRING * 31
    PROPNOT2 AS STRING * 31
    PROPNOT3 AS STRING * 31
    StormFee AS SINGLE 'not being used in VA
'End of form
    CustPin  AS LONG
    NextRec  AS LONG
    LastYrPrinted AS INTEGER
    Deleted  AS INTEGER
    Blank    AS STRING * 124
END TYPE

TYPE PersonalRecType
   PROPPIN  AS STRING * 20
   PROPDATE AS INTEGER
   PersVal  AS DOUBLE
   MHVALUE  AS DOUBLE
   MCVALUE  AS DOUBLE
   CVALUE   AS DOUBLE
   MTVALUE  AS DOUBLE
   EXMPSENI AS DOUBLE
   EXMPOTHR AS DOUBLE
   DISCOV   AS STRING * 1
   LATELIST AS STRING * 1
   DESC1    AS STRING * 30
   DESC2    AS STRING * 30
   DESC3    AS STRING * 30
   Desc4    AS STRING * 30
   Desc5    AS STRING * 30
'end of form
   CustPin        AS LONG
   NextRec        AS LONG
   LastYrPrinted  AS INTEGER
   Deleted        AS INTEGER
   VehTaxYear     AS INTEGER
   DMVSubmitted   AS STRING * 1
   Blank          AS STRING * 121
END TYPE

TYPE ChilRecType
 blk AS STRING * 6
 Acct AS STRING * 7
 Blk1 AS STRING * 4
 bus AS STRING * 1
 nam AS STRING * 35
 a1 AS STRING * 30
 a2 AS STRING * 30
 a3 AS STRING * 31
 z1 AS STRING * 5
 Blk2 AS STRING * 1
 z2 AS STRING * 4
 Blk3 AS STRING * 19
 ss1 AS STRING * 9
 Blk4 AS STRING * 1
 ss2 AS STRING * 9
END TYPE

TYPE ChilPRecType
   blk AS STRING * 6
   Acct AS STRING * 7
   Blk1 AS STRING * 5
   PType AS STRING * 2
   Vin AS STRING * 23
   Blk2 AS STRING * 5
   Value AS STRING * 10
   Blk3 AS STRING * 11
   Make AS STRING * 10
   Model AS STRING * 15
   Body AS STRING * 16
   Year AS STRING * 4
   Blk4 AS STRING * 4
   Lic AS STRING * 12
   Blk5 AS STRING * 30
   Qual AS STRING * 1
END TYPE

TYPE ElkppType
 AcctNumber     AS STRING * 8
 Name           AS STRING * 35
 Add1           AS STRING * 35
 Add2           AS STRING * 33
 City           AS STRING * 12
 Zip            AS STRING * 5
 Blank          AS STRING * 28
 TaxDue         AS STRING * 9
 BillNo         AS STRING * 7
 ss1            AS STRING * 12
 ss2            AS STRING * 12
END TYPE

