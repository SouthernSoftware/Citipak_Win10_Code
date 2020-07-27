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
    RcptPort AS INTEGER
    Padding AS STRING * 254
END TYPE

TYPE TaxValuesType
    Value    AS DOUBLE
    OthVal   AS DOUBLE
    ExmVal   AS DOUBLE
END TYPE

TYPE TaxCustType
    Acct       AS LONG
    OPENDATE   AS INTEGER
    FNAME      AS STRING * 15
    LName      AS STRING * 25
    SName      AS STRING * 10
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

    PAD1         AS STRING * 232
    LastTrans     AS LONG        'Pointer to last transaction
    FirstPropRec  AS LONG        'Pointer to first property rec
    FirstPersRec  AS LONG        'Pointer to first personal rec
    PIN           AS LONG        'Cust internal id number.
    Deleted       AS INTEGER     'deleted flag
    FileVer       AS INTEGER     'this is the file struct version number
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
    Future1       AS DOUBLE
    Future2       AS DOUBLE
    Principle1Pd  AS DOUBLE
    Principle2Pd  AS DOUBLE    'For Va Only
    Principle3Pd  AS DOUBLE    'For Va Only
    Principle4Pd  AS DOUBLE    'For Va Only
    Principle5Pd  AS DOUBLE    'For Va Only
    InterestPd    AS DOUBLE
    PenaltyPd     AS DOUBLE
    CollectionPd  AS DOUBLE
    Future1Pd     AS DOUBLE
    Future2Pd     AS DOUBLE
END TYPE

TYPE TaxTransactionType
    TransDate    AS INTEGER          'Transaction Date
    TaxYear      AS INTEGER          'Must Contain Full 4 digit Tax Year Here
    TranType     AS INTEGER          '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
    BillType     AS STRING * 1       'R=Real P=Personal Property C=Combined (NC/GA)
    Amount       AS DOUBLE           'Total Transaction Amount
    Revenue      AS RevSourceType    'See Revenue Source Type File above
    Description  AS STRING * 30      'Description of Transaction
    Posted2GL    AS STRING * 1       'I/F to G/L Yes or No
    CustomerRec  AS LONG             'Pointer Back to Customer Record
    LastTrans    AS LONG             'Points to Previous Trans in History
    'actually Previous pointer
    BelongTo     AS LONG             'Points to Record of Bill this Transaction Belongs to : 'Will be 0 for Bill
    Padding      AS STRING * 128     'Allow for Future Expansion
END TYPE

TYPE InterestRecType
     CustRec            AS LONG                 'Acct #
     CustName           AS STRING * 40
     TaxYear            AS INTEGER
     Amount             AS DOUBLE
     BillNumber         AS STRING * 10
     CurYear            AS INTEGER
'end of form
     BillRec            AS LONG
     DelFlag            AS INTEGER
     Padding            AS STRING * 159
END TYPE

TYPE TaxMTransactionType
    Account      AS LONG
    TransDate    AS INTEGER
    TaxYear      AS INTEGER
    Desc         AS STRING * 30
    TaxAmount    AS DOUBLE
    IntAmount    AS DOUBLE
    AdColAmount  AS DOUBLE
    BillType     AS STRING * 1   'R=REAL P=PERS C=COMB
    SName        AS STRING * 30
    TName        AS STRING * 30
    Padding      AS STRING * 128
END TYPE

TYPE MortCodeRecType
    MortCode AS STRING * 8
    BName    AS STRING * 32
    Add1     AS STRING * 32
    Add2     AS STRING * 32
    Add3     AS STRING * 32
    Contact  AS STRING * 32
    Phone    AS STRING * 14
    Pad      AS STRING * 254
END TYPE


'Old Customer Layouts

TYPE TBCustType
   Acct      AS STRING * 11
   FirstName AS STRING * 26
   LName     AS STRING * 35
   Addr1     AS STRING * 35
   Addr2     AS STRING * 35
   City      AS STRING * 20
   State     AS STRING * 2
   Zip       AS STRING * 10
   Phone     AS STRING * 12
   PDesc     AS STRING * 39
   PMap      AS STRING * 22
   SubDiv    AS STRING * 35
   Lot       AS SINGLE
   County    AS STRING * 35
   PSize     AS SINGLE
   PLand     AS DOUBLE
   PBldg     AS DOUBLE
   PBus      AS DOUBLE
   PRes      AS DOUBLE
   PPub      AS DOUBLE
   PersExp   AS SINGLE
   SCExp     AS SINGLE
   HomeExp   AS SINGLE
   UseExp    AS SINGLE
   CLate     AS SINGLE
   CTax      AS SINGLE
   CInt      AS SINGLE
   CAdv      AS SINGLE
   CCol      AS SINGLE
   CNotice   AS SINGLE
   FTr       AS SINGLE
   Ltr       AS SINGLE
   SSN       AS STRING * 11
   Late      AS STRING * 1
   PIN       AS STRING * 16
END TYPE






   

