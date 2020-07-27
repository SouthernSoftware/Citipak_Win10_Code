TYPE CACUSTType
    ACCT       AS LONG
    CustName   AS STRING * 40
    ADD1       AS STRING * 40
    ADD2       AS STRING * 40
    LEGALDES   AS STRING * 40
    PROPDESC   AS STRING * 40
    PARCEL     AS STRING * 40

    STRAMT     AS SINGLE
    SIDEAMT    AS SINGLE
    DRIVEAMT   AS SINGLE
    DEMOAMT    AS SINGLE
    WATERAMT   AS SINGLE
    SEWERAMT   AS SINGLE

    STRDATE    AS INTEGER
    SIDEDATE   AS INTEGER
    DRIVEDATE  AS INTEGER
    DEMODATE   AS INTEGER
    WATERDATE  AS INTEGER
    SEWERDATE  AS INTEGER

    STRRATE    AS SINGLE
    SIDERATE   AS SINGLE
    DRIVERATE  AS SINGLE
    DEMORATE   AS SINGLE
    WATERRATE  AS SINGLE
    SEWERRATE  AS SINGLE

    STRTERM    AS INTEGER
    SIDETERM   AS INTEGER
    DRIVETERM  AS INTEGER
    DEMOTERM   AS INTEGER
    WATERTERM  AS INTEGER
    SEWERTERM  AS INTEGER
    ADDNOTES AS STRING * 70
   'End Screen Here
    BegTrans AS LONG
    LastTrans AS LONG
    Deleted     AS INTEGER
    IStreet     AS SINGLE
    ISide       AS SINGLE
    IDrive      AS SINGLE
    IDemo       AS SINGLE
    IWater      AS SINGLE
    ISewer      AS SINGLE
    RPBStreet   AS SINGLE
    RPBSide     AS SINGLE
    RPBDrive    AS SINGLE
    RPBDemo     AS SINGLE
    RPBWater    AS SINGLE
    RPBSewer    AS SINGLE
    RIBStreet   AS SINGLE
    RIBSide     AS SINGLE
    RIBDrive    AS SINGLE
    RIBDemo     AS SINGLE
    RIBWater    AS SINGLE
    RIBSewer    AS SINGLE
    LastPayDate AS INTEGER
    LastIntDate AS INTEGER
    Pad         AS STRING * 52
END TYPE

TYPE CABegTransType
    ACCT     AS LONG
    CustName AS STRING * 40
    STRBegBal   AS SINGLE
    STRIntAcc   AS SINGLE
    STRPrinPd  AS SINGLE
    STRIntPd   AS SINGLE
    SIDEBegBal  AS SINGLE
    SIDEIntAcc  AS SINGLE
    SIDEPrinPd  AS SINGLE
    SIDEIntPd   AS SINGLE
    DRIVEBegBal  AS SINGLE
    DRIVEIntAcc  AS SINGLE
    DRIVEPrinPd  AS SINGLE
    DRIVEIntPd   AS SINGLE
    DEMOBegBal   AS SINGLE
    DEMOIntAcc   AS SINGLE
    DEMOPrinPd   AS SINGLE
    DEMOIntPd    AS SINGLE
    WATERBegBal  AS SINGLE
    WATERIntAcc  AS SINGLE
    WATERPrinPd  AS SINGLE
    WATERIntPd   AS SINGLE
    SEWERBegBal  AS SINGLE
    SEWERIntAcc  AS SINGLE
    SEWERPrinPd  AS SINGLE
    SEWERIntPd   AS SINGLE
    'End Screen Form
    Pad AS STRING * 64
END TYPE

TYPE CaTransType
     TranType     AS INTEGER               '1=Interest 2=Payment 3=Adjustment
     Date         AS INTEGER
     IStreetAmt   AS SINGLE
     ISideAmt     AS SINGLE
     IDriveAmt    AS SINGLE
     IDemoAmt     AS SINGLE
     IWaterAmt    AS SINGLE
     ISewerAmt    AS SINGLE
     PStreetAmt   AS SINGLE
     PSideAmt     AS SINGLE
     PDriveAmt    AS SINGLE
     PDemoAmt     AS SINGLE
     PWaterAmt    AS SINGLE
     PSewerAmt    AS SINGLE
     NextTrans    AS LONG
     CustRecord   AS LONG
     Pad          AS STRING * 68
END TYPE

TYPE CaIntEditType
     IStreetAmt   AS SINGLE
     ISideAmt     AS SINGLE
     IDriveAmt    AS SINGLE
     IDemoAmt     AS SINGLE
     IWaterAmt    AS SINGLE
     ISewerAmt    AS SINGLE
     CustRecord   AS LONG
     CustName     AS STRING * 40
     DelFlag      AS INTEGER
     Pad          AS STRING * 22
END TYPE

TYPE CaPayEditType
     CustAcct     AS LONG
     PStreetAmt   AS SINGLE
     PSideAmt     AS SINGLE
     PDriveAmt    AS SINGLE
     PDemoAmt     AS SINGLE
     PWaterAmt    AS SINGLE
     PSewerAmt    AS SINGLE
     IStreetAmt   AS SINGLE
     ISideAmt     AS SINGLE
     IDriveAmt    AS SINGLE
     IDemoAmt     AS SINGLE
     IWaterAmt    AS SINGLE
     ISewerAmt    AS SINGLE
     TotalAmtPd   AS SINGLE
     DelFlag      AS INTEGER
     RecptNumber  AS STRING * 15
     PayDesc      AS STRING * 15
     PayDate      AS INTEGER
     Pad          AS STRING * 38
END TYPE

