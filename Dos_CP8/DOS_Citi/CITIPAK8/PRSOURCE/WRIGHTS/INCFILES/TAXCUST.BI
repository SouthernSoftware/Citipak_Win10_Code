CONST CustRecVerNO = -100

TYPE PINRecType
   PIN AS LONG
END TYPE

TYPE TaxMasterType      'Master Default Information in Setup
   Name AS STRING * 35
   Add1 AS STRING * 35
   Add2 AS STRING * 35
   Add3 AS STRING * 35
   TaxSt AS STRING * 2
   TaxForm AS STRING * 20
   CurRate AS SINGLE
   PastRate AS SINGLE
   PenRate AS SINGLE
   Padding AS STRING * 256
END TYPE

TYPE TaxValuesType
    Value    AS DOUBLE
    OthVal   AS DOUBLE
    ExmVal   AS DOUBLE
END TYPE

TYPE TaxCustType
    ACCT       AS LONG
    OPENDATE   AS INTEGER
    FNAME      AS STRING * 15
    LName      AS STRING * 25
    SNAME      AS STRING * 10
    HPHONE     AS STRING * 14
    WPHONE     AS STRING * 14
    CSSN       AS STRING * 11
    SSSN       AS STRING * 11
    Addr1      AS STRING * 35
    Addr2      AS STRING * 35
    City       AS STRING * 20
    State      AS STRING * 2
    Zip        AS STRING * 10
    Active     AS STRING * 1    'Y if Active N if Inactive
    Interest   AS STRING * 1    'Y/N to Charge Interest
    TaxExempt  AS STRING * 1    'Y/N to Charge Taxes Period
    Penalty    AS STRING * 1    'Y/N to Charge Penalty
'end of form

    TotalReal(1 TO 1)  AS TaxValuesType
    TotalPers(1 TO 1)  AS TaxValuesType

    PAD1         AS STRING * 236
    
    FirstPropRec  AS LONG      'Pointer to last property rec
    FirstPersRec  AS LONG      'Pointer to last personal rec
    PIN        AS LONG        'Cust program id number.
    Deleted    AS INTEGER     'deleted flag
    FileVer    AS INTEGER     'this is the file struct version number

END TYPE

TYPE RevSourceType
     Principle1    AS DOUBLE
     Principle2    AS DOUBLE    'For Va Only
     Principle3    AS DOUBLE    'For Va Only
     Principle4    AS DOUBLE    'For Va Only
     Principle5    AS DOUBLE    'For Va Only
     Interest      AS DOUBLE
     Penalty       AS DOUBLE
     Collection    AS DOUBLE
     AdCost        AS DOUBLE
     LateFee       AS DOUBLE
     Principle1Pd  AS DOUBLE
     Principle2Pd  AS DOUBLE    'For Va Only
     Principle3Pd  AS DOUBLE    'For Va Only
     Principle4Pd  AS DOUBLE    'For Va Only
     Principle5Pd  AS DOUBLE    'For Va Only
     InterestPd    AS DOUBLE
     PenaltyPd     AS DOUBLE
     CollectionPd  AS DOUBLE
     AdCostPd      AS DOUBLE
     LateFeePd     AS DOUBLE
END TYPE


TYPE TaxTransactionType
     TransDate    AS INTEGER          'Transaction Date
     TaxYear      AS INTEGER          'Must Contain Full 4 digit Tax Year Here
     TranType     AS INTEGER          '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty
     BillType     AS STRING * 1       'R=Real P=Personal Property
     Amount       AS DOUBLE           'Total Transaction Amount
     Revenue      AS RevSourceType    'See Revenue Source Type File Below
     Description  AS STRING * 30      'Description of Transaction
     Posted2GL    AS STRING * 1       'I/F to G/L Yes or No
     CustomerRec  AS LONG             'Pointer Back to Customer Record
     LastTrans    AS LONG             'Points to Last Transaction in Customer History
     Padding      AS STRING * 128     'Allow for Future Expansion
END TYPE

TYPE MortCodeRecType
     MortCode AS STRING * 5
     Name AS STRING * 32
     Add1 AS STRING * 32
     Add2 AS STRING * 32
     Add3 AS STRING * 32
     Contact AS STRING * 32
     Phone AS STRING * 14
     Pad AS STRING * 254
END TYPE





'Old Customer Layouts

TYPE TBCustType
   ACCT      AS STRING * 12
   FirstName AS STRING * 25
   LName     AS STRING * 35
   Addr1     AS STRING * 35
   Addr2     AS STRING * 35
   City      AS STRING * 20
   State     AS STRING * 2
   Zip       AS STRING * 10
   Phone     AS STRING * 12
   Fill      AS STRING * 326
END TYPE



