TYPE DayReceiptsType
   Date         AS INTEGER
   AmtIn        AS DOUBLE
   AmtOwed      AS DOUBLE
   CashAmt      AS DOUBLE
   ChkAmt       AS DOUBLE
   TranType     AS INTEGER  '1=Tax Pay Real 2=Tax Pay Pers 3=Sup Tax Real 4=Sup Tax Pers 5=Decal 6=Misc
   TranCode     AS INTEGER  '1=Misc Code Rec #
   TaxYear      AS INTEGER
   Desc         AS STRING * 30
   RealTxPd     AS DOUBLE
   PersTxPd     AS DOUBLE
   MTTxPd       AS DOUBLE
   MCTxPd       AS DOUBLE
   MHTxPd       AS DOUBLE
   FETxPd       AS DOUBLE
   IntTxPd      AS DOUBLE
   PenTxPd      AS DOUBLE
   DecalNumb    AS STRING * 10
   Drawer       AS INTEGER
   Fudge        AS STRING * 110
END TYPE

TYPE NADA
    Blank AS STRING * 1
    NCIC AS STRING * 5
    Year AS STRING * 2
    VIN AS STRING * 17
    State AS STRING * 2
    VehRecord AS STRING * 5
    CusRecord AS STRING * 5
    NotUsed AS STRING * 59
    NADAFILLER AS STRING * 222
    ValPeriod AS STRING * 4
    NADAFILLER1 AS STRING * 378
END TYPE

TYPE TBRECust
  MapNumber         AS STRING * 14
  District          AS STRING * 2
  CustName          AS STRING * 35
  Address1          AS STRING * 35
  Address2          AS STRING * 35
  Address3          AS STRING * 35
  Zipcode           AS STRING * 5
  PropDesc1         AS STRING * 25
  PropDesc2         AS STRING * 25
  PropDesc3         AS STRING * 25
  PropDesc4         AS STRING * 25
  Zone              AS STRING * 2
  MGCode            AS STRING * 2
  ClassCode         AS STRING * 2
  Tract             AS STRING * 18
  SqFtSize          AS STRING * 8
  Acres             AS STRING * 12
  TransferYr        AS STRING * 2
  Comment1          AS STRING * 20
  Comment2          AS STRING * 20
  Comment3          AS STRING * 20
  LandValue         AS DOUBLE
  BldgValue         AS DOUBLE
  LandUseValue      AS DOUBLE
  NewConst          AS STRING * 1                  'Y/N
  NewValue          AS STRING * 1                  'Y/N
  NewAddress        AS STRING * 1                  'Y/N
  LastDateChange    AS INTEGER
  CurYrBillNumber   AS STRING * 7
  CurYrTaxAmount    AS DOUBLE
  CurYrIntAmount    AS DOUBLE
  CurYrPenAmount    AS DOUBLE
  CurYrOthAmount    AS DOUBLE
  FirstTrans        AS SINGLE
  LastTrans         AS SINGLE
  Deleted           AS STRING * 1
  LBCurYrBillNumber AS STRING * 7
  LBCurYrTaxAmount  AS DOUBLE
  RELandValue       AS DOUBLE
  REBldgValue       AS DOUBLE
  NewBldgValue      AS DOUBLE
  NewCom1           AS STRING * 25
  NewCom2           AS STRING * 25
  BBCurYrBillNumber AS STRING * 7
  BBCurYrTaxAmount  AS DOUBLE
  MPCurYrBillNumber AS STRING * 7       'Municipality Bills (Victoria/Kenbridge)
  MPCurYrTaxAmount  AS DOUBLE
  LASTYEARBILLED    AS INTEGER
  MortgageYN        AS STRING * 1
  Pad               AS STRING * 5
END TYPE

TYPE TBREBAL
 TaxAmt(1 TO 10) AS SINGLE
 IntAmt(1 TO 10) AS SINGLE
 PenAmt(1 TO 10) AS SINGLE
 OthAmt(1 TO 10) AS SINGLE
 LateAmt(1 TO 10) AS SINGLE
 TaxNotice(1 TO 10) AS STRING * 7
END TYPE

TYPE TBPPCust
        CustName         AS STRING * 25
        Address1         AS STRING * 25
        Address2         AS STRING * 25
        City             AS STRING * 20
        State            AS STRING * 2
        Zip              AS STRING * 10
        SocSec           AS STRING * 12
        SocSec2          AS STRING * 12
        Short            AS STRING * 10
        PersProp         AS DOUBLE
        MachTools        AS DOUBLE
        MobileHome       AS DOUBLE
        MerchCap         AS DOUBLE
        FarmEquip        AS DOUBLE
        Business         AS STRING * 1  'y/n
        LastChangeDate   AS INTEGER
        District         AS STRING * 2
        LateFile         AS STRING * 1  'y/n
        CurYrBillNumber  AS STRING * 7
        CurYrPersPropTax        AS DOUBLE
        CurYrFarmEquipTax       AS DOUBLE
        CurYrMachToolsTax       AS DOUBLE
        CurYrMerchCapTax        AS DOUBLE
        CurYrMobileHomeTax      AS DOUBLE
        CurYrLateFeeTax         AS DOUBLE
        SocSec3          AS STRING * 12
        FirstTrans       AS SINGLE
        LastTrans        AS SINGLE
        Deleted          AS STRING * 1
        CurYrIntAmount   AS DOUBLE
        CurYrPenAmount   AS DOUBLE
        LBCurYrBillNumber AS STRING * 7
        LBCurYrPersPropTaxAmount AS DOUBLE
        LBCurYrFarmEquipTaxAmount AS DOUBLE
        LBCurYrMachToolsTaxAmount AS DOUBLE
        LBCurYrMerchCapTaxAmount AS DOUBLE
        LBCurYrMobileHomeTaxAmount AS DOUBLE
        LBLateFeeAmount AS DOUBLE
        LName2 AS STRING * 35
        FName2 AS STRING * 35
        FirstVeh AS LONG
        LastVeh AS LONG
        HeavyEquip   AS DOUBLE
        Boats AS DOUBLE
        BusFurn AS DOUBLE
        TBCurYrBillNumber AS STRING * 7
        TBCurYrPersPropTaxAmount AS DOUBLE
        TBCurYrFarmEquipTaxAmount AS DOUBLE
        TBCurYrMachToolsTaxAmount AS DOUBLE
        TBCurYrMerchCapTaxAmount AS DOUBLE
        TBCurYrMobileHomeTaxAmount AS DOUBLE
        TBLateFeeAmount  AS DOUBLE
        LASTYEARBILLED   AS INTEGER
        DMVSubmitted     AS STRING * 1          '1998 vehicles
        DMVBatch         AS STRING * 4          '1998 vehicles
        DecalBought      AS STRING * 1
        DecalNumber      AS STRING * 10
        LBPPTRAValue     AS DOUBLE
        LBPPTRADiscnt    AS DOUBLE
        TBPPTRAValue     AS DOUBLE
        TBPPTRADiscnt    AS DOUBLE
        DMVSubmitted1999 AS STRING * 1          '1999 vehicles
        DMVBatch1999     AS STRING * 4          '1999 vehicles
        DMVSubmitted2000 AS STRING * 1          '2000 vehicles
        DMVBatch2000     AS STRING * 4          '2000 vehicles
        DMVSubmitted2001 AS STRING * 1          '2001 vehicles
        DMVBatch2001     AS STRING * 4          '2001 vehicles
        DMVSubmitted2002 AS STRING * 1          '2002 vehicles
        DMVBatch2002     AS STRING * 4          '2002 vehicles
        BoatFirstTrans   AS SINGLE
        BoatLastTrans    AS SINGLE
        MHFirstTrans     AS SINGLE
        MHLastTrans      AS SINGLE
        DMVSubmitted2003 AS STRING * 1          '2003 vehicles
        DMVBatch2003     AS STRING * 4          '2003 vehicles
        DMVSubmitted2004 AS STRING * 1          '2004 vehicles
        DMVBatch2004     AS STRING * 4          '2004 vehicles
        DMVSubmitted2005 AS STRING * 1          '2005 vehicles
        DMVBatch2005     AS STRING * 4          '2005 vehicles
        DMVSubmitted2006 AS STRING * 1          '2006 vehicles
        DMVBatch2006     AS STRING * 4          '2006 vehicles
        DMVSubmitted2007 AS STRING * 1          '2007 vehicles
        DMVBatch2007     AS STRING * 4          '2007 vehicles
        DMVSubmitted2008 AS STRING * 1          '2008 vehicles
        DMVBatch2008     AS STRING * 4          '2008 vehicles
        DMVSubmitted2009 AS STRING * 1          '2009 vehicles
        DMVBatch2009     AS STRING * 4          '2009 vehicles
        Sup2003          AS STRING * 1          '2003 Sup Bill
        Sup2004          AS STRING * 1          '2003 Sup Bill

        Sup2005          AS STRING * 1
        Pad              AS STRING * 159
        END TYPE

TYPE TownType
       TBCurYrBillNumber AS STRING * 7
       TBCurYrPersPropTaxAmount AS DOUBLE
       TBCurYrFarmEquipTaxAmount AS DOUBLE
       TBCurYrMachToolsTaxAmount AS DOUBLE
       TBCurYrMerchCapTaxAmount AS DOUBLE
       TBCurYrMobileHomeTaxAmount AS DOUBLE
       TBLateFeeAmount  AS DOUBLE
       TBPPTRAValue     AS DOUBLE
       TBPPTRADiscnt    AS DOUBLE
END TYPE
TYPE TBExceptionType
   SSN1 AS STRING * 12
   SSN2 AS STRING * 12
   LName AS STRING * 45
   FName AS STRING * 33
   MName AS STRING * 31
   VType AS STRING * 1
   Add1  AS STRING * 35
   Add2  AS STRING * 35
   City  AS STRING * 22
   State AS STRING * 2
   Zip AS STRING * 9
   VIN AS STRING * 22
   VehMake AS STRING * 20
   VehModl AS STRING * 15
   VehYr AS STRING * 4
   VehWt AS STRING * 5
   TradeVal AS STRING * 6
END TYPE

TYPE PPVehType
  VehTyp   AS STRING * 1                'D=Deleted Do Not Use
  VehVin   AS STRING * 22
  VehMake  AS STRING * 20
  VehYear  AS STRING * 4
  VehModel AS STRING * 15
  VehWght  AS LONG
  VehValue AS LONG
  VehQ     AS STRING * 1
  VehLic   AS STRING * 10
  VehNext  AS LONG
  VehAcct  AS LONG
  VehUpd   AS STRING * 1                'used to mark record each year for dmv process
  VehTaxYr AS INTEGER                   'used to keep vehicles actual billing year seperated
  VehTrail AS STRING * 1                'Trailor Code
  VehMiles AS STRING * 1                'Flag to indicate high mileage
  VehSupFlag AS STRING * 1              'Indicates Vehicle is a supplemental veh for that tax year only
  VehPad   AS STRING * 33
END TYPE

TYPE PPBoat
 BoatYear AS STRING * 4
 BoatMake AS STRING * 15
 BoatHull AS STRING * 11
 BoatLgth AS STRING * 6
 BoatDate AS STRING * 8
 BoatCost AS DOUBLE
 BoatMotorMake AS STRING * 15
 BoatMotorHP AS STRING * 6
 BoatMotorDate AS STRING * 8
 BoatMotorCost AS DOUBLE
 BoatActive AS STRING * 1               'Y if Active
 BoatRecord AS SINGLE
 BoatNext AS SINGLE
 BoatPad AS STRING * 30
END TYPE
TYPE PPMHome
 MHLandOwner AS STRING * 24
 MHYear AS STRING * 4
 MHMake AS STRING * 15
 MHLgth AS STRING * 6
 MHWdth AS STRING * 6
 MHDate AS STRING * 8
 MHCost AS DOUBLE
 MHActive AS STRING * 1         'Y=Yes
 MHRecord AS SINGLE
 MHNext AS SINGLE
 MHPad AS STRING * 48
END TYPE



TYPE TBPPBal
 PerPropTaxAmt(1 TO 10) AS SINGLE
 FarmEquipAmt(1 TO 10) AS SINGLE
 MachToolsAmt(1 TO 10) AS SINGLE
 MerchCapAmt(1 TO 10) AS SINGLE
 MobileHomeAmt(1 TO 10) AS SINGLE
 LateFeeAmt(1 TO 10) AS SINGLE
 IntAmt(1 TO 10) AS SINGLE
 PenAmt(1 TO 10) AS SINGLE
 TaxNotice(1 TO 10) AS STRING * 7
END TYPE



TYPE PaymentEdit
 PaymentAcct AS SINGLE           'Hold Customer Acct #
 PaymentType AS INTEGER         '1=Payment 2=Release
 PaymentDate AS STRING * 6
 PaymentAmt AS DOUBLE
 PaymentDesc AS STRING * 15
 PaymentSource AS STRING * 1        'R=Real  P=Personal
 PaymentDist AS STRING * 1          'A=Auto  M=Manual
 PaymentManYear AS STRING * 4
 PaymentManTaxAmt AS DOUBLE
 PaymentManIntAmt AS SINGLE
 PaymentManPenAmt AS SINGLE
 Pad AS STRING * 7
END TYPE


TYPE TBTrans
 TransDate AS INTEGER
 TransAmt AS DOUBLE
 TrTaxYear AS STRING * 4
 TrDesc AS STRING * 15
 TrType AS INTEGER      '1=bill 2=payment 3=interest 4=penalty 5=release 6=bal fwd
 TrTax AS SINGLE
 TrInterest AS SINGLE
 TrPenalty AS SINGLE
 TrLate AS SINGLE
 TrPersProp AS SINGLE
 TrMachTools AS SINGLE
 TrMobileHome AS SINGLE
 TrMerchCap AS SINGLE
 TrFarmEquip AS SINGLE
 TrCustomerRecord AS SINGLE
 TrNextTrans AS SINGLE
 TrRunningBal AS SINGLE
 TrSource AS STRING * 1         'R=Real P=Personal
 TrPPTRA AS SINGLE
 TrPad AS STRING * 12
END TYPE
 
TYPE TBLog
 LogAcct AS SINGLE
 LogOrigAmt AS DOUBLE
 LogNewAmt AS DOUBLE
 LogName AS STRING * 20
 LogCode AS INTEGER
END TYPE

TYPE TBPage
 Page AS INTEGER
 Class1 AS DOUBLE
 Class2 AS DOUBLE
 Class3 AS DOUBLE
 Class4 AS DOUBLE
 Class5 AS DOUBLE
 Class6 AS DOUBLE
 Class7 AS DOUBLE
 Class8 AS DOUBLE
 Class9 AS DOUBLE
 Class0 AS DOUBLE
 Customers AS SINGLE
 LandValue AS DOUBLE
 BldgValue AS DOUBLE
 TaxValue AS DOUBLE
 ExLandValue AS DOUBLE
 ExBldgValue AS DOUBLE
 ExTaxValue AS DOUBLE
END TYPE

TYPE TBPage1
 Page AS INTEGER
 AlphaWho AS STRING * 1
 PPValue AS DOUBLE
 PPAmount AS DOUBLE
 FEValue AS DOUBLE
 FEAmount AS DOUBLE
 MTValue AS DOUBLE
 MTAmount AS DOUBLE
 MCValue AS DOUBLE
 MCAmount AS DOUBLE
 MHValue AS DOUBLE
 MHAmount AS DOUBLE
 LateAmount AS DOUBLE
 TotalTaxAmount AS DOUBLE
 NotUsed AS STRING * 64
END TYPE

TYPE TBRealExport
  MapNumber      AS STRING * 14
  District       AS STRING * 2
  CustName       AS STRING * 35
  Address1       AS STRING * 35
  Address2       AS STRING * 35
  Address3       AS STRING * 35
  Zipcode        AS STRING * 5
  PropDesc1      AS STRING * 25
  PropDesc2      AS STRING * 25
  PropDesc3      AS STRING * 25
  PropDesc4      AS STRING * 25
  Zone           AS STRING * 2
  MGCode         AS STRING * 2
  ClassCode      AS STRING * 2
  Tract          AS STRING * 18
  SqFtSize       AS STRING * 8
  Acres          AS STRING * 12
  TransferYr     AS STRING * 2
  Comment1       AS STRING * 20
  Comment2       AS STRING * 20
  Comment3       AS STRING * 20
  LandValue      AS STRING * 10
  BldgValue      AS STRING * 10
  LandUseValue   AS STRING * 10
  Account        AS STRING * 10
END TYPE
TYPE TBPersonalExport
        CustName         AS STRING * 25
        Address1         AS STRING * 25
        Address2         AS STRING * 25
        City             AS STRING * 20
        State            AS STRING * 2
        Zip              AS STRING * 10
        SocSec           AS STRING * 12
        SocSec2          AS STRING * 12
        Short            AS STRING * 10
        PersProp         AS STRING * 10
        MachTools        AS STRING * 10
        MobileHome       AS STRING * 10
        MerchCap         AS STRING * 10
        FarmEquip        AS STRING * 10
        Business         AS STRING * 1  'y/n
        District         AS STRING * 2
        LateFile         AS STRING * 1  'y/n
        Account          AS STRING * 10
END TYPE




REM These Are Old Programmers File Layouts
TYPE TBRealCustomer
  Account AS STRING * 7
  MapNumber AS STRING * 14
  District AS STRING * 2
  CustName AS STRING * 35
  Address1 AS STRING * 35
  Address2 AS STRING * 35
  Address3 AS STRING * 35
  Zipcode AS STRING * 5
  PropDesc1 AS STRING * 25
  PropDesc2 AS STRING * 25
  PropDesc3 AS STRING * 25
  PropDesc4 AS STRING * 25
  OldLandValue AS STRING * 8
  OldBldgValue AS STRING * 8
  OldLandUseValue AS STRING * 8
  NotUsed AS STRING * 4
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
END TYPE

TYPE TBPPCustomer
        Account AS STRING * 6
        CustName AS STRING * 25
        Address1 AS STRING * 25
        Address2 AS STRING * 25
        City AS STRING * 20
        State AS STRING * 2
        Zip AS STRING * 10
        SocSec AS STRING * 12
        SocSec2 AS STRING * 12
        TType AS STRING * 1
        Short AS STRING * 10
        PersProp AS STRING * 11
        MachTools AS STRING * 11
        MobileHome AS STRING * 11
        MerchCap AS STRING * 11
        FarmEquip AS STRING * 11
        Business AS STRING * 1
        LastChangeDate AS STRING * 8
        District AS STRING * 2
        LateFile AS STRING * 1
        BillNumb AS STRING * 4
        PersPropTax AS STRING * 11
        FarmEquipTax AS STRING * 11
        MachToolsTax AS STRING * 11
        MerchCapTax AS STRING * 11
        MobileHomeTax AS STRING * 11
        LateFeeTax AS STRING * 11
        SocSec3 AS STRING * 12
END TYPE





REM End of Old Programmers Layouts



TYPE OldTBCust1Data
        acct AS STRING * 12
        lnam AS STRING * 60
        Add1 AS STRING * 35
        Add2 AS STRING * 35
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
        NotUsed AS STRING * 2
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
        Blank AS STRING * 69
END TYPE

TYPE OldTBCust3Data1
 TaxAmt  AS SINGLE
 IntAmt AS SINGLE
 AdvAmt AS SINGLE
 ColAmt AS SINGLE
 LateAmt AS SINGLE
 TaxNotice AS SINGLE
END TYPE

        
'TYPE OldTBCust3Data
' sAA(1 TO 10) AS OldTBCust3Data1
' sFill AS STRING * 16
'END TYPE

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
 Blank AS STRING * 26
END TYPE


TYPE Interest
LastTime AS INTEGER
END TYPE

TYPE DMVInformationType
 PerRate AS SINGLE
 Batch   AS LONG
 JCode   AS STRING * 4
END TYPE

TYPE TBREKENCust
  AccountNumber     AS STRING * 5
  MapNumber         AS STRING * 14
  District          AS STRING * 2
  CustName          AS STRING * 35
  Address1          AS STRING * 35
  Address2          AS STRING * 35
  Address3          AS STRING * 35
  Zipcode           AS STRING * 5
  PropDesc1         AS STRING * 25
  PropDesc2         AS STRING * 25
  PropDesc3         AS STRING * 25
  PropDesc4         AS STRING * 25
  Zone              AS STRING * 2
  ClassCode         AS STRING * 2
  Tract             AS STRING * 18
  SqFtSize          AS STRING * 8
  Acres             AS STRING * 12
  Comment1          AS STRING * 20
  Comment2          AS STRING * 20
  Comment3          AS STRING * 20
  LandValue         AS STRING * 10
  BldgValue         AS STRING * 10
  LandUseValue      AS STRING * 10
  MPCurYrBillNumber AS STRING * 7       'Municipality Bills (Victoria/Kenbridge)
  MPCurYrTaxAmount  AS STRING * 10

END TYPE

TYPE TBKENPPCust
        AcctNumber       AS STRING * 10
        CustName         AS STRING * 25
        Address1         AS STRING * 25
        Address2         AS STRING * 25
        City             AS STRING * 20
        State            AS STRING * 2
        Zip              AS STRING * 10
        SocSec           AS STRING * 12
        SocSec2          AS STRING * 12
        PersProp         AS STRING * 10
        MachTools        AS STRING * 10
        MobileHome       AS STRING * 10
        MerchCap         AS STRING * 10
        FarmEquip        AS STRING * 10
        Business         AS STRING * 1  'y/n
        District         AS STRING * 2
        LateFile         AS STRING * 1  'y/n
        Deleted          AS STRING * 1
        LName2           AS STRING * 35
        FName2           AS STRING * 35
        HeavyEquip       AS STRING * 10
        Boats            AS STRING * 10
        BusFurn          AS STRING * 10
        TBCurYrBillNumber AS STRING * 7
        TBCurYrPersPropTaxAmount AS STRING * 10
        TBCurYrFarmEquipTaxAmount AS STRING * 10
        TBCurYrMachToolsTaxAmount AS STRING * 10
        TBCurYrMerchCapTaxAmount AS STRING * 10
        TBCurYrMobileHomeTaxAmount AS STRING * 10
        TBLateFeeAmount  AS STRING * 10
        TBPPTRAValue     AS STRING * 10
        TBPPTRADiscnt    AS STRING * 10
END TYPE

TYPE PPKenVehType
  VehTyp   AS STRING * 1                'D=Deleted Do Not Use
  VehVin   AS STRING * 22
  VehMake  AS STRING * 20
  VehYear  AS STRING * 4
  VehModel AS STRING * 15
  VehWght  AS STRING * 10
  VehValue AS STRING * 10
  VehQ     AS STRING * 1
  VehLic   AS STRING * 10
  VehAcct  AS STRING * 6
END TYPE


