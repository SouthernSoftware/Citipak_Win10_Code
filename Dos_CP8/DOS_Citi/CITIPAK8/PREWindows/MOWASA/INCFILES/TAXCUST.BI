
CONST CustRecVerNO = -100

TYPE PINRecType
   pin AS LONG
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

TYPE TaxInterestDateType
    InterestDate AS INTEGER
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
    CITY       AS STRING * 20
    STATE      AS STRING * 2
    Zip        AS STRING * 10
    Active     AS STRING * 1    'Y if Active N if Inactive
    Interest   AS STRING * 1    'Y/N to Charge Interest
    TaxExempt  AS STRING * 1    'Y/N to Charge Taxes Period
    Penalty    AS STRING * 1    'Y/N to Charge Penalty
'end of form
    TotalReal(1 TO 1)  AS TaxValuesType
    TotalPers(1 TO 1)  AS TaxValuesType

    PAD1         AS STRING * 228
    CountyAcct    AS LONG        'County Account Number to Link to County Records if applic
    LastTrans     AS LONG        'Pointer to last transaction
    FirstPropRec  AS LONG        'Pointer to first property rec
    FirstPersRec  AS LONG        'Pointer to first personal rec
    pin           AS LONG        'Cust internal id number.
    Deleted       AS INTEGER     'deleted flag
    FileVer       AS INTEGER     'this is the file struct version number

END TYPE

TYPE RevSourceType
    Principle1    AS DOUBLE                 'Va Personal Prop
    Principle2    AS DOUBLE    'For Va Only     Mach/Tools
    Principle3    AS DOUBLE    'For Va Only     Merch Cap
    Principle4    AS DOUBLE    'For Va Only     Farm Equip
    Principle5    AS DOUBLE    'For Va Only     Mobile Homes
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
    TranType     AS INTEGER          '1=Bill 2=Payment 3=Release 4=Interest
                                     '5=Penalty 8=Collection/Ad Cost Billing
                                     '7=Adjustment
    BillType     AS STRING * 1       'R=Real P=Personal Property C=Combined (NC/GA)
    Amount       AS DOUBLE           'Total Transaction Amount
    Revenue      AS RevSourceType    'See Revenue Source Type File above
    Description  AS STRING * 30      'Description of Transaction
    Posted2GL    AS STRING * 1       'I/F to G/L Yes or No
    CustomerRec  AS LONG             'Pointer Back to Customer Record
    LastTrans    AS LONG             'Points to Previous Trans in History
    'actually Previous pointer
    BelongTo     AS LONG             'Points to Record of Bill this Transaction Belongs to : 'Will be 0 for Bill
    DMVSubmitted AS STRING * 1       'Y if Sent to DMV
    DMVBatch     AS INTEGER          'Records which batch contained the DMV Transmission
    Padding      AS STRING * 125     'Allow for Future Expansion
END TYPE

TYPE InterestRecType
     CustRec            AS LONG                 'Acct #
     CustName           AS STRING * 40
     TaxYear            AS INTEGER
     Amount             AS DOUBLE
     BillNumber         AS STRING * 10
     CurYear            AS INTEGER
'end of form
     billrec            AS LONG
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
    MHTax        AS DOUBLE
    MCTax        AS DOUBLE
    FETax        AS DOUBLE
    MTTax        AS DOUBLE
    Padding      AS STRING * 96
END TYPE

TYPE MortCodeRecType
    MortCode AS STRING * 8
    BName    AS STRING * 32
    Add1     AS STRING * 32
    Add2     AS STRING * 32
    Add3     AS STRING * 32
    Contact  AS STRING * 32
    PHONE    AS STRING * 14
    Pad      AS STRING * 254
END TYPE



TYPE TBCust1Type
    ptaxamt0 AS SINGLE
    pintamt0 AS SINGLE
    padvamt0 AS SINGLE
    pcolamt0 AS SINGLE
    plateamt0 AS SINGLE
    ptaxnotice0 AS SINGLE
    ptaxamt1 AS SINGLE
    pintamt1 AS SINGLE
    padvamt1 AS SINGLE
    pcolamt1 AS SINGLE
    plateamt1 AS SINGLE
    ptaxnotice1 AS SINGLE
    ptaxamt2 AS SINGLE
    pintamt2 AS SINGLE
    padvamt2 AS SINGLE
    pcolamt2 AS SINGLE
    plateamt2 AS SINGLE
    ptaxnotice2 AS SINGLE
    ptaxamt3 AS SINGLE
    pintamt3 AS SINGLE
    padvamt3 AS SINGLE
    pcolamt3 AS SINGLE
    plateamt3 AS SINGLE
    ptaxnotice3 AS SINGLE
    ptaxamt4 AS SINGLE
    pintamt4 AS SINGLE
    padvamt4 AS SINGLE
    pcolamt4 AS SINGLE
    plateamt4 AS SINGLE
    ptaxnotice4 AS SINGLE
    ptaxamt5 AS SINGLE
    pintamt5 AS SINGLE
    padvamt5 AS SINGLE
    pcolamt5 AS SINGLE
    plateamt5 AS SINGLE
    ptaxnotice5 AS SINGLE
    ptaxamt6 AS SINGLE
    pintamt6 AS SINGLE
    padvamt6 AS SINGLE
    pcolamt6 AS SINGLE
    plateamt6 AS SINGLE
    ptaxnotice6 AS SINGLE
    ptaxamt7 AS SINGLE
    pintamt7 AS SINGLE
    padvamt7 AS SINGLE
    pcolamt7 AS SINGLE
    plateamt7 AS SINGLE
    ptaxnotice7 AS SINGLE
    ptaxamt8 AS SINGLE
    pintamt8 AS SINGLE
    padvamt8 AS SINGLE
    pcolamt8 AS SINGLE
    plateamt8 AS SINGLE
    ptaxnotice8 AS SINGLE
    ptaxamt9 AS SINGLE
    pintamt9 AS SINGLE
    padvamt9 AS SINGLE
    pcolamt9 AS SINGLE
    plateamt9 AS SINGLE
    ptaxnotice9 AS SINGLE
END TYPE

      

TYPE TBCustType
   Acct      AS STRING * 12
   firstnamE AS STRING * 25
   LName     AS STRING * 35
   Addr1     AS STRING * 35
   Addr2     AS STRING * 35
   CITY      AS STRING * 20
   STATE     AS STRING * 2
   Zip       AS STRING * 10
   PHONE     AS STRING * 12
   pdesc     AS STRING * 39
   MAP      AS STRING * 22
   SubDiv    AS STRING * 35
   Lot       AS SINGLE
   County    AS STRING * 35
   psize     AS SINGLE
   pland     AS DOUBLE
   mv        AS DOUBLE
   mh        AS DOUBLE
   fm        AS DOUBLE
   ls        AS DOUBLE
   op        AS SINGLE
   me        AS SINGLE
   Homeexp   AS STRING * 6
   TaxTYpe   AS STRING * 1
   NU        AS STRING * 3
   CLate     AS SINGLE
   CTax      AS SINGLE
   CInt      AS SINGLE
   CAdv      AS SINGLE
   CCol      AS SINGLE
   CNotice   AS SINGLE
   FTR       AS SINGLE
   LTR       AS SINGLE
   SSN       AS STRING * 11
   late      AS STRING * 1
   pin       AS STRING * 15
END TYPE

TYPE RAMAST
   SSN AS STRING * 9
   BILLNUM AS STRING * 10
   MAP AS STRING * 18
   taxyr AS STRING * 2
   SUP AS STRING * 2
   Name AS STRING * 38
   Add1 AS STRING * 32
   Add2 AS STRING * 32
   CITY AS STRING * 21
   STATE AS STRING * 2
   Zip AS STRING * 9
   ACRES AS STRING * 9
   DESC1 AS STRING * 32
   DESC2 AS STRING * 32
   LANDVAL AS STRING * 9
   BLDGVAL AS STRING * 9
   filler AS STRING * 16
   CLASS AS STRING * 2
   DISTRICT AS STRING * 3
   filler2 AS STRING * 104
END TYPE

TYPE RABILL
 FYear AS STRING * 2
 TYear AS STRING * 2
 CustRec AS STRING * 5
 filler AS STRING * 1
 billrec AS STRING * 3
 filler2 AS STRING * 12
 DueDate AS STRING * 6
 TaxDue AS STRING * 11
 late AS STRING * 1
 filler3 AS STRING * 22
 AmtDue AS STRING * 11
 filler4 AS STRING * 89
END TYPE
   
TYPE Oldtbcustdat  '256 byte file
   LName     AS STRING * 30
   firstnamE AS STRING * 25
   Addr1     AS STRING * 30
   Addr2     AS STRING * 30
   CITY      AS STRING * 20
   STATE     AS STRING * 2
   Zip       AS STRING * 5
   PHONE     AS STRING * 12
   pdesc     AS STRING * 20
   MAP       AS STRING * 15
   pland     AS DOUBLE
   mv        AS DOUBLE
   SrCit     AS DOUBLE
   TaxTYpe   AS STRING * 2
   Acct      AS STRING * 12
   CurBal    AS DOUBLE
   PastDue   AS DOUBLE
   FTR       AS SINGLE
   LTR       AS SINGLE
END TYPE
TYPE OldTBCUST1DAT              '64 byte record tbcust1.dat
        CurYrTax    AS DOUBLE
        PastYrTax   AS DOUBLE
        LateList    AS SINGLE
        TaxTYpe     AS STRING * 1
        SSN         AS STRING * 11
END TYPE

TYPE HamletTaxRec
 Acct AS STRING * 9
 Name AS STRING * 30
 Add1 AS STRING * 30
 Add2 AS STRING * 30
 Add3 AS STRING * 30
 Add4 AS STRING * 30
 Zip AS STRING * 9
 PropDesc AS STRING * 30
 Lots AS STRING * 5
 Real AS STRING * 8
 Pers AS STRING * 7
 Discnt AS STRING * 5
 NU AS STRING * 43
 SSN AS STRING * 9
 NU1 AS STRING * 282
 SSN1 AS STRING * 9
END TYPE

