CONST MaxRevsCnt = 15

TYPE NRevSetUpType
    RevName AS STRING * 15
    UseDep   AS STRING * 1
    USERATE  AS STRING * 1
    TAXRATE  AS SINGLE
    UseMtr   AS STRING * 1
    DistOr   AS INTEGER
END TYPE

TYPE NUBSetupRecType
    UTILNAME    AS STRING * 35
    DEFCITY     AS STRING * 18
    DEFSTATE    AS STRING * 2
    ZIPCODE     AS STRING * 10
    FLYDeposit  AS STRING * 1
    EQPAY       AS STRING * 1
    WORKORD     AS STRING * 1
    ESTREAD     AS STRING * 1
    BANKDFT     AS STRING * 1
    UseSeq      AS STRING * 1
    BILLCYCL    AS STRING * 1
    BARCODE     AS STRING * 1
    RemLookFlag AS STRING * 1
    LowRead     AS SINGLE
    HighRead    AS SINGLE
    HHDEVICE    AS STRING * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 TO 15) AS NRevSetUpType
    Fill AS STRING * 65
END TYPE

TYPE RevSetUpType
    RevName AS STRING * 15
    UseDep   AS STRING * 1
    USERATE  AS STRING * 1
    TAXRATE  AS SINGLE
    UseMtr   AS STRING * 1
    DistOr   AS INTEGER
END TYPE

TYPE UBSetupRecType
    UTILNAME    AS STRING * 35
    DEFCITY     AS STRING * 18
    DEFSTATE    AS STRING * 2
    ZIPCODE     AS STRING * 10
    FLYDeposit  AS STRING * 1
    RLookFlag   AS STRING * 1
    WORKORD     AS STRING * 1
    ESTREAD     AS STRING * 1
    BANKDFT     AS STRING * 1
    UseSeq      AS STRING * 1
    BILLCYCL    AS STRING * 1
    BARCODE     AS STRING * 1
    LowRead     AS SINGLE
    HighRead    AS SINGLE
    HHDEVICE    AS STRING * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 TO 15) AS RevSetUpType
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
    ServAddress AS STRING * 20
    MeterID AS STRING * 8
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    Account AS STRING * 8
    SensusType AS STRING * 1        ' B=Touch Read : M=Manual
END TYPE

TYPE UBSensusGetReadRecType     ' File Layout For Retreiving Records
    Account AS STRING * 8
    NotUsed AS STRING * 4
    MeterID AS STRING * 8
    Reading AS STRING * 8
    'NotUse1 AS STRING * 6
    NotUse1 AS STRING * 14
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


