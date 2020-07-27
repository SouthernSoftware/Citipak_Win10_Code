
TYPE TBRealCustomer
  Account AS STRING * 7
  MapNumber AS STRING * 14
  District AS STRING * 2
  CustName AS STRING * 35
  Address1 AS STRING * 35
  Address2 AS STRING * 35
  address3 AS STRING * 35
  zipcode AS STRING * 9
  PropDesc1 AS STRING * 25
  PropDesc2 AS STRING * 25
  PropDesc3 AS STRING * 25
  PropDesc4 AS STRING * 25
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
  CurYrTaxAmount AS STRING * 11
 'FirstTrans AS SINGLE
 'LastTrans AS SINGLE
 'Future AS STRING * 128
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


