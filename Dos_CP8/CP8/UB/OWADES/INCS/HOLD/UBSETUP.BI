CONST MaxRevsCnt = 15

TYPE NRevSetUpType
    RevName AS STRING * 15
    UseDep   AS STRING * 1
    USERATE  AS STRING * 1
    TAXRATE  AS SINGLE
    UseMtr   AS STRING * 1
    DistOr   AS INTEGER
END TYPE

TYPE RevSetUpType
    RevName AS STRING * 15
    UseDep   AS STRING * 1
    USERATE  AS STRING * 1
    TAXRATE  AS SINGLE
    UseMtr   AS STRING * 1
    DistOr   AS INTEGER
END TYPE

TYPE SetUpAcctType
   RevName    AS STRING * 15
   DebitAcct  AS STRING * 14
   CreditAcct AS STRING * 14
END TYPE

TYPE nUBSetupRecType
    UTILNAME        AS STRING * 35
    DEFCITY         AS STRING * 18
    DEFSTATE        AS STRING * 2
    ZIPCODE         AS STRING * 10
    PreByBook       AS STRING * 1
    RecpPort        AS STRING * 1
    RECPDEFT        AS STRING * 1
    ESTREAD         AS STRING * 1
    BANKDFT         AS STRING * 1
    UseSeq          AS STRING * 1
    BILLCYCL        AS STRING * 1
    DefLook         AS STRING * 1
    MethAcct        AS STRING * 1      'new 02-14-97
    SkipInactive    AS STRING * 1
    SkipSeparator   AS STRING * 1
    Fill1           AS STRING * 1
    LowRead         AS INTEGER
    HighRead        AS INTEGER
    HHDEVICE        AS STRING * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
'new 051098
    FileVer         AS DOUBLE
'****
    Revenues(1 TO 15) AS RevSetUpType
    BillAcct(1 TO 15) AS SetUpAcctType
     PayAcct(1 TO 15) AS SetUpAcctType
     DepAcct(1 TO 15) AS SetUpAcctType
END TYPE

TYPE UBSetupRecType
    UTILNAME        AS STRING * 35
    DEFCITY         AS STRING * 18
    DEFSTATE        AS STRING * 2
    ZIPCODE         AS STRING * 10
    PreByBook       AS STRING * 1
    RecpPort        AS STRING * 1
    RECPDEFT        AS STRING * 1
    ESTREAD         AS STRING * 1
    BANKDFT         AS STRING * 1
    UseSeq          AS STRING * 1
    BILLCYCL        AS STRING * 1
    DefLook         AS STRING * 1
    MethAcct        AS STRING * 1      'new 02-14-97
    SkipInactive    AS STRING * 1
    SkipSeparator   AS STRING * 1
    Fill1           AS STRING * 1
    LowRead         AS INTEGER
    HighRead        AS INTEGER
    HHDEVICE        AS STRING * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 TO 15) AS RevSetUpType
    BillAcct(1 TO 15) AS SetUpAcctType
    PayAcct(1 TO 15)  AS SetUpAcctType
    DepAcct(1 TO 15)  AS SetUpAcctType
END TYPE

TYPE UBSensusRecType
    PathWay AS STRING * 20
END TYPE

TYPE UBPC3000RecType
    PathWay AS STRING * 20
END TYPE

TYPE DistArrayType
   DistOrder AS INTEGER
   DistCnt   AS INTEGER
END TYPE

' These Sensus Layout Files are Spec'd Out Exactly to Long View NC

TYPE UBSensusReadRecType        ' File Layout for Sending Out Records
    ServAddress  AS STRING * 20
    MeterID      AS STRING * 8
    LowRead      AS STRING * 8
    HighRead     AS STRING * 8
    Account      AS STRING * 8
    SensusType   AS STRING * 1        ' B=Touch Read : M=Manual
    CustName     AS STRING * 25
    SerialNumb   AS STRING * 8        'Added Per Mickey on 6-23-97
END TYPE

TYPE UBSensusGetReadRecType     ' File Layout For Retreiving Records
    Account AS STRING * 8
    NotUsed AS STRING * 5
    MeterID AS STRING * 8
    Filler AS STRING * 1
    Reading AS STRING * 8
    NotUse1 AS STRING * 1
    DateRead AS STRING * 4
    NotUse2 AS STRING * 6
END TYPE

TYPE UBPC3000ReadRecType        ' File Layout for Sending Out Records
    CustName AS STRING * 20
    ServAddress AS STRING * 20
    MeterID AS STRING * 8
    LowRead AS DOUBLE
    HighRead AS DOUBLE
    Account AS STRING * 8
    ReadFlag AS STRING * 1        'Y/N
    MeterType AS STRING * 1
    Book AS INTEGER
    CurRead AS DOUBLE
    PastRead AS DOUBLE
    ReadDate AS INTEGER
    ReadTime AS STRING * 5
    Note1 AS STRING * 20
    Note2 AS STRING * 20
    Note3 AS STRING * 20
    NoteStatus AS STRING * 1    'T=Temp Note    P=Perm Note
END TYPE

TYPE UBDGRecType
    PathWay AS STRING * 20
END TYPE

TYPE UBDGProcRecType        ' File Layout for Sending Out Records
    RouteID AS STRING * 20
    SvcTyp AS STRING * 1
    CustName AS STRING * 25
    SvcLoc AS STRING * 21
    MeterSN AS STRING * 20
    MeterType AS STRING * 1       ' C for reg mtr   D for demand elec
    High AS STRING * 10
    Low AS STRING * 10
    Msg AS STRING * 110
    Account AS STRING * 10
    NewRdng AS STRING * 10
    NewDmnd AS STRING * 10
    Date AS STRING * 6
    Time AS STRING * 6
    NewAcctRte AS STRING * 20
END TYPE

