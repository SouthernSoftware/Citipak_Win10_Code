CONST MaxRevsCnt = 15

TYPE oRevSetUpType
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

TYPE RevSetUpType
    RevName AS STRING * 15
    UseDep   AS STRING * 1
    USERATE  AS STRING * 1
    TAXRATE  AS SINGLE
    UseMtr   AS STRING * 1
    DistOr   AS INTEGER
    ProRate  AS STRING * 1
END TYPE

TYPE oUBSetupRecType
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
    Make99File      AS STRING * 1
    LowRead         AS INTEGER
    HighRead        AS INTEGER
    HHDEVICE        AS STRING * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 TO 15) AS oRevSetUpType
    BillAcct(1 TO 15) AS SetUpAcctType
    PayAcct(1 TO 15)  AS SetUpAcctType
    DepAcct(1 TO 15)  AS SetUpAcctType
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
    Make99File      AS STRING * 1
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

TYPE UBLogiconRecType
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

TYPE UBLogiconReadRecType
 RecType   AS STRING * 1
 RouteNo   AS STRING * 2
 AcctNo    AS STRING * 6
 RecName   AS STRING * 30
 ServAddress AS STRING * 25
 ReadDate    AS STRING * 6
 ReadTime    AS STRING * 6
 Consumption AS STRING * 8
 PrevRead    AS STRING * 8
 CurRead     AS STRING * 8
 LowRead     AS STRING * 8
 HighRead   AS STRING * 8
 MtrNumb    AS STRING * 12
 CountChg   AS STRING * 1
 ForceFlag  AS STRING * 1
 ReportCode AS STRING * 2
 Remark     AS STRING * 40
 Label      AS STRING * 19
 PrintFlag  AS STRING * 1
 MessageOut AS STRING * 30
 book       AS STRING * 2
 Future     AS STRING * 29
 Recend     AS STRING * 1               'Must be 'X'
 CRLF       AS STRING * 2
END TYPE


TYPE UBLogiconGetReadRecType
 RecType   AS STRING * 1
 RouteNo   AS STRING * 2
 AcctNo    AS STRING * 6
 RecName   AS STRING * 30
 ServAddress AS STRING * 25
 ReadDate    AS STRING * 6
 ReadTime    AS STRING * 6
 Consumption AS STRING * 8
 PrevRead    AS STRING * 8
 CurRead     AS STRING * 8
 LowRead     AS STRING * 8
 HighRead   AS STRING * 8
 MtrNumb    AS STRING * 12
 CountChg   AS STRING * 1
 ForceFlag  AS STRING * 1
 ReportCode AS STRING * 2
 Remark     AS STRING * 40
 Label      AS STRING * 19
 PrintFlag  AS STRING * 1
 MessageOut AS STRING * 30
 book       AS STRING * 2
 Future     AS STRING * 29
 Recend     AS STRING * 1               'Must be 'X'
 CRLF       AS STRING * 2
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
    book AS INTEGER
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


'Itron Layouts

TYPE UBItronRecType
    PathWay AS STRING * 20
END TYPE


TYPE ItronFType                         'Header Record Type
    RecordCode AS STRING * 1
    Route AS STRING * 8
    Message AS STRING * 64
    Filler AS STRING * 5
    CRLF AS STRING * 2
END TYPE

TYPE ItronAType                         'Customer Record One
    RecordCode AS STRING * 1
    Route AS STRING * 8
    AcctNumb AS STRING * 10
    Geo AS STRING * 12
    SeqNumb AS STRING * 5
    Message AS STRING * 1
    AType AS STRING * 1
    Filler AS STRING * 40
    CRLF AS STRING * 2
END TYPE

TYPE ItronBType                         'Customer Record Two
    RecordCode AS STRING * 1
    CustName AS STRING * 32
    CustAddr AS STRING * 32
    Filler AS STRING * 13
    CRLF AS STRING * 2
END TYPE

TYPE ItronHType
    RecordCode AS STRING * 1
    MeterNumb AS STRING * 12
    Dials AS STRING * 1
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    LastRead AS STRING * 8
    MeterType AS STRING * 1
    Multiplier AS STRING * 6
    NoMths AS STRING * 1
    MtrMake AS STRING * 2
    DispCode AS STRING * 2
    NumbDec AS STRING * 1
    MustRead AS STRING * 1
    Status AS STRING * 1
    Filler AS STRING * 25
    CRLF AS STRING * 2
END TYPE

TYPE ItronZType
    RecordCode AS STRING * 1
    Route AS STRING * 8
    NumberAccts AS STRING * 4
    NumberMeters AS STRING * 4
    Filler AS STRING * 61
    CRLF AS STRING * 2
END TYPE

'Itron Read Layouts
TYPE ItronCType
     RecordCode AS STRING * 1           'Must be a C
     Route      AS STRING * 8
     Acct       AS STRING * 10
     SpecInst   AS STRING * 2
     Survey     AS STRING * 1
     ReadStatus AS STRING * 1
     ReSeqFlag  AS STRING * 1
     ReadDate   AS STRING * 4           'mmdd
     AcctChg    AS STRING * 1
     Filler     AS STRING * 51
END TYPE
TYPE ItronDType
     RecordCode AS STRING * 1           'Must be a D
     MeterNum   AS STRING * 12
     LocCode    AS STRING * 2
     MtrCon1    AS STRING * 2
     MtrCon2    AS STRING * 2
     Mult       AS STRING * 6
     ChgeFlag   AS STRING * 1
     Filler     AS STRING * 54
END TYPE
TYPE ItronEType
     RecordCode AS STRING * 1           'Must be a E
     NbrDials   AS STRING * 1
     CurReading AS STRING * 8           'Zero's if No Read
     DecPlaces  AS STRING * 1
     ReadTime   AS STRING * 6
     ReadChg    AS STRING * 1
     DispCode   AS STRING * 2
     ReadCount  AS STRING * 1
     ReadVerify AS STRING * 1
     NoReadCode AS STRING * 2
     Filler     AS STRING * 56
END TYPE

