TYPE REALTYPE                              'ST PAUL VA
 TAXID AS STRING * 10
 Map AS STRING * 50
 LASTNAME AS STRING * 14
 FirstName AS STRING * 14
 INITIAL AS STRING * 2
 LASTNAME2 AS STRING * 14
 FIRSTNAME2 AS STRING * 14
 INITIAL2 AS STRING * 2
 ADDRESS1 AS STRING * 30
 ADDRESS2 AS STRING * 30
 City AS STRING * 18
 State AS STRING * 2
 ZipCode AS STRING * 10
 DESCRIPTION AS STRING * 36
 ENTRYTYPE AS STRING * 2
 DISTRICT AS STRING * 2
 CLASS AS STRING * 2
 ACRE AS STRING * 8
 BLDGVALUE AS STRING * 8
 LANDVALUE AS STRING * 8
 RESVAL AS STRING * 8
 DEVVAL AS STRING * 8
 TOTVAL AS STRING * 9
 EXEMPT AS STRING * 6
 DIABILITY AS STRING * 2
 REMARKS AS STRING * 60
 BLK AS STRING * 3
END TYPE

TYPE PERSTYPE                   'ST PAUL VA
  LASTNAME AS STRING * 32
  FirstName AS STRING * 20
  INITIAL AS STRING * 2
  BLK AS STRING * 54
  ADDRESS1 AS STRING * 32
  ADDRESS2 AS STRING * 32
  City AS STRING * 20
  State AS STRING * 2
  Zip AS STRING * 10
  SSN AS STRING * 12
  CUSTCODE AS STRING * 2
  SSN2 AS STRING * 12
  PROPTYPE AS STRING * 2
  PRICE AS STRING * 9
  AVALUE AS STRING * 9
  LEVY AS STRING * 8
  EXEMPT AS STRING * 7
  NETTAX AS STRING * 8
  PAGENO AS STRING * 5
  DISTRICT AS STRING * 2
  TaxType AS STRING * 2
  BLK13 AS STRING * 19
  MVLIC AS STRING * 10
  MVMAKE AS STRING * 16
  MVVIN AS STRING * 22
  MVYR AS STRING * 4
  MVPPTRA AS STRING * 2
  BLK5 AS STRING * 4

END TYPE

TYPE ZipType
    ZipCode AS STRING * 5
    City   AS STRING * 40
END TYPE

TYPE TBCustType              'len=512
     Acct       AS STRING * 8
     pin        AS STRING * 4
     FirstName  AS STRING * 25
     LName      AS STRING * 35
     Addr1       AS STRING * 35
     Addr2       AS STRING * 35
     City       AS STRING * 20
     State      AS STRING * 2
     Zip        AS STRING * 10
     Phone  AS STRING * 12
     pdesc      AS STRING * 39
     pmap       AS STRING * 22
     Pdesc2     AS STRING * 35
     plot       AS STRING * 4
     PDesc3     AS STRING * 35
     psize      AS SINGLE
     pland      AS DOUBLE
     PBldg         AS DOUBLE
     MH         AS DOUBLE
     FM         AS DOUBLE
     LS         AS DOUBLE
     OP         AS SINGLE
     ME         AS SINGLE
     homeexp    AS SINGLE
     UseExp    AS SINGLE
     CLate AS SINGLE
     CTax  AS SINGLE
     CInt  AS SINGLE
     CAdv  AS SINGLE
     CCol  AS SINGLE
     CNotice AS SINGLE
     ftr        AS SINGLE
     ltr        AS SINGLE
     SSN        AS STRING * 11
     late       AS STRING * 1
     END TYPE
TYPE TBCust2RecType     'len=512
      D1        AS STRING * 30
      D2        AS STRING * 30
      D3        AS STRING * 30
      D4        AS STRING * 30
      D5        AS STRING * 30
      V1        AS STRING * 18
      V2        AS STRING * 18
      V3        AS STRING * 18
      V4        AS STRING * 18
      V5        AS STRING * 18
      VAL1      AS DOUBLE
      VAL2      AS DOUBLE
      VAL3      AS DOUBLE
      VAL4      AS DOUBLE
      VAL5      AS DOUBLE
      D6        AS STRING * 30
      V6        AS STRING * 18
      Val6      AS DOUBLE
      D7        AS STRING * 30
      V7        AS STRING * 18
      Val7      AS DOUBLE
      D8        AS STRING * 30
      V8        AS STRING * 18
      Val8      AS DOUBLE
      DYr1      AS INTEGER
      DYr2      AS INTEGER
      DYr3      AS INTEGER
      DYr4      AS INTEGER
      DYr5      AS INTEGER
      DYr6      AS INTEGER
      DYr7      AS INTEGER
      DYr8      AS INTEGER
      DMV1      AS STRING * 1
      DMV2      AS STRING * 1
      DMV3      AS STRING * 1
      DMV4      AS STRING * 1
      DMV5      AS STRING * 1
      DMV6      AS STRING * 1
      DMV7      AS STRING * 1
      DMV8      AS STRING * 1
END TYPE
TYPE TBCust1Type     'len=256
  PTaxAmt0      AS SINGLE
  PIntAmt0      AS SINGLE
  PAdvAmt0      AS SINGLE
  PColAmt0      AS SINGLE
  PLateAmt0     AS SINGLE
  PTaxNotice0   AS SINGLE
  PTaxAmt1      AS SINGLE
  PIntAmt1      AS SINGLE
  PAdvAmt1      AS SINGLE
  PColAmt1      AS SINGLE
  PLateAmt1     AS SINGLE
  PTaxNotice1   AS SINGLE
  PTaxAmt2      AS SINGLE
  PIntAmt2      AS SINGLE
  PAdvAmt2      AS SINGLE
  PColAmt2      AS SINGLE
  PLateAmt2     AS SINGLE
  PTaxNotice2   AS SINGLE
  PTaxAmt3      AS SINGLE
  PIntAmt3      AS SINGLE
  PAdvAmt3      AS SINGLE
  PColAmt3      AS SINGLE
  PLateAmt3     AS SINGLE
  PTaxNotice3   AS SINGLE
  PTaxAmt4      AS SINGLE
  PIntAmt4      AS SINGLE
  PAdvAmt4      AS SINGLE
  PColAmt4      AS SINGLE
  PLateAmt4     AS SINGLE
  PTaxNotice4   AS SINGLE
  PTaxAmt5      AS SINGLE
  PIntAmt5      AS SINGLE
  PAdvAmt5      AS SINGLE
  PColAmt5      AS SINGLE
  PLateAmt5     AS SINGLE
  PTaxNotice5   AS SINGLE
  PTaxAmt6      AS SINGLE
  PIntAmt6      AS SINGLE
  PAdvAmt6      AS SINGLE
  PColAmt6      AS SINGLE
  PLateAmt6     AS SINGLE
  PTaxNotice6   AS SINGLE
  PTaxAmt7      AS SINGLE
  PIntAmt7      AS SINGLE
  PAdvAmt7      AS SINGLE
  PColAmt7      AS SINGLE
  PLateAmt7     AS SINGLE
  PTaxNotice7   AS SINGLE
  PTaxAmt8      AS SINGLE
  PIntAmt8      AS SINGLE
  PAdvAmt8      AS SINGLE
  PColAmt8      AS SINGLE
  PLateAmt8     AS SINGLE
  PTaxNotice8   AS SINGLE
  PTaxAmt9      AS SINGLE
  PIntAmt9      AS SINGLE
  PAdvAmt9      AS SINGLE
  PColAmt9      AS SINGLE
  PLateAmt9     AS SINGLE
  PTaxNotice9   AS SINGLE
END TYPE


TYPE IOwner
 Owner AS SINGLE
 NameID AS SINGLE
END TYPE







TYPE Oldtbcustdat
  LName     AS STRING * 30
  FirstName AS STRING * 25
  Addr1     AS STRING * 30
  Addr2     AS STRING * 30
  City      AS STRING * 20
  State     AS STRING * 2
  Zip       AS STRING * 5
  Phone     AS STRING * 12
  PropDesc  AS STRING * 20
  Map       AS STRING * 15
  RealValue AS DOUBLE
  PersValue AS DOUBLE
  SrCitValue AS DOUBLE
  CarTag    AS INTEGER
  Acct      AS STRING * 8
  CurBal    AS DOUBLE
  PastDue   AS DOUBLE
END TYPE

TYPE OldTBCUST1DAT
 CurYrTax AS DOUBLE
 PastYrTax AS DOUBLE
 LastList AS SINGLE
END TYPE

