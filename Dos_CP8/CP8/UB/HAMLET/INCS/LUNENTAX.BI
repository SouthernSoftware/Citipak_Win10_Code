
TYPE TBRealCustomer
  Account AS STRING * 7
  MapNumber AS STRING * 14
  District AS STRING * 2
  CustName AS STRING * 35
  Address1 AS STRING * 35
  Address2 AS STRING * 35
  address3 AS STRING * 35
  Zipcode AS STRING * 5
  Desc1 AS STRING * 25
  Desc2 AS STRING * 25
  Desc3 AS STRING * 25
  Desc4 AS STRING * 25
  OldLandValue AS STRING * 8
  OldBldgValue AS STRING * 8
  OldLandUseValue AS STRING * 8
  Notused AS STRING * 4
  Zone AS STRING * 2
  MGCode AS STRING * 2
  ClassCode AS STRING * 2
  Tract AS STRING * 18
  SqFtSize AS STRING * 8
  Notused1 AS STRING * 4
  Acres AS STRING * 12
  Notused2 AS STRING * 8
  TransferYr AS STRING * 2
  Comment1 AS STRING * 20
  Comment2 AS STRING * 20
  Comment3 AS STRING * 20
  LandValue AS STRING * 9
  BldgValue AS STRING * 9
  LandUseValue AS STRING * 9
  RV AS STRING * 1
  NewConstruction AS STRING * 1
  NewValue AS STRING * 1
  NewAddress AS STRING * 1
  LastDate AS STRING * 8
  BillNumber AS STRING * 5
  TaxAmount AS STRING * 11
END TYPE

TYPE TBPPCustomer
  Account   AS STRING * 6
  CustName  AS STRING * 25
  Address1  AS STRING * 25
  Address2  AS STRING * 25
  City      AS STRING * 20
  State     AS STRING * 2
  Zip       AS STRING * 10
  SocSec    AS STRING * 12
  SocSec2   AS STRING * 12
  TType      AS STRING * 1
  Short     AS STRING * 10
  PersProp  AS STRING * 11
  MachTools AS STRING * 11
  MobileHome AS STRING * 11
  MerchCap  AS STRING * 11
  FarmEquip AS STRING * 11
  Business  AS STRING * 1
  LastChangeDate AS STRING * 8
  District  AS STRING * 2
  LateFile  AS STRING * 1        'Y/N
  BillNumb  AS STRING * 4
  PersPropTax   AS STRING * 11
  FarmEquipTax  AS STRING * 11
  MachToolsTax  AS STRING * 11
  MerchCapTax   AS STRING * 11
  MobileHomeTax AS STRING * 11
  LateFeeTax    AS STRING * 11
  SocSec3       AS STRING * 12
END TYPE


TYPE OldTBCust1Data
        acct AS STRING * 12
        lnam AS STRING * 60
        add1 AS STRING * 35
        add2 AS STRING * 35
        City AS STRING * 20
        State AS STRING * 2
        Zip AS STRING * 10
        Tel AS STRING * 12
        PropDesc AS STRING * 39
        Map AS STRING * 22
        SubDiv AS STRING * 35
        plot AS STRING * 4
        County AS STRING * 35
        Size AS SINGLE
        LandValue AS DOUBLE
        BldgValue AS DOUBLE
        MobileHome AS DOUBLE
        FarmMach AS DOUBLE
        PersProp AS DOUBLE
        MachTools AS SINGLE
        MerchCap AS SINGLE
        Elderly AS SINGLE
        Notused AS STRING * 2
        taxtype AS STRING * 1           'R=Real   P=Pers
        TypeTax AS STRING * 3           'Business Y/N
        CURLATEAMT AS SINGLE
        CURTAXAMT AS SINGLE
        CURintAMT AS SINGLE
        CURADVAMT AS SINGLE
        CURCOLAMT AS SINGLE
        CurTaxNotice AS SINGLE
        ftr AS SINGLE
        ltr AS SINGLE
        SSN AS STRING * 11
        LateFee AS STRING * 1
        pin AS STRING * 15
        Taxable AS STRING * 1
        blank AS STRING * 69
END TYPE

TYPE OldTBCust3Data
 TaxAmt(1 TO 10) AS SINGLE
 IntAmt(1 TO 10) AS SINGLE
 AdvAmt(1 TO 10) AS SINGLE
 ColAmt(1 TO 10) AS SINGLE
 LateAmt(1 TO 10) AS SINGLE
 TaxNotice(1 TO 10) AS SINGLE
END TYPE

TYPE OldARTrans
 TransDate AS STRING * 6
 TransAmt AS SINGLE
 TaxYear AS SINGLE
 Desc AS STRING * 20
 TrType AS INTEGER
 TrPrinciple AS SINGLE
 TrInterest AS SINGLE
 TrAdv AS SINGLE
 TrCol AS SINGLE
 TrLate AS SINGLE
 MasterCustomerRecord AS SINGLE
 NextTrans AS SINGLE
 RunningBal AS SINGLE
 TrDisc AS SINGLE
 blank AS STRING * 26
END TYPE


