TYPE BookSeqRecType
  BookSeq  AS LONG
END TYPE

TYPE ServicesType
    RateCode AS STRING * 4
    RMTRTYPE AS STRING * 1
END TYPE

TYPE LocMeterType
    MTRNum    AS STRING * 12
    MTRMulti  AS INTEGER
    MTRType   AS STRING * 1
    MTRUnit   AS STRING * 1
    NumUser   AS INTEGER
    InsDate   AS INTEGER
    CurRead   AS LONG
    PrevRead  AS LONG
    CurDate   AS INTEGER
    PastDate  AS INTEGER       'hidden & protected
    ReadFlag  AS STRING * 1    'hidden & protected
    AvgUse    AS LONG          'hidden & protected
    UseCnt    AS INTEGER       'hidden & protected
    '  MtrPad    AS STRING * 14   'hidden & protected
END TYPE

TYPE FlatRateType
    FRDesc   AS STRING * 18
    FRAmt    AS DOUBLE
    FRFreq   AS STRING * 1
    RevSrc   AS INTEGER
    NumMin   AS INTEGER
END TYPE

TYPE UBLocationRecType
    BOOK           AS STRING * 2
    SEQNUMB        AS STRING * 6
    LocatStatus    AS STRING * 1    ' 1
    SERVADDR       AS STRING * 35
    BILLCYCL       AS INTEGER
    ZONE           AS STRING * 3
    SEQ            AS LONG
    BILLWHO        AS STRING * 1
    PUMPCODE       AS STRING * 4
    USERCODE1      AS STRING * 4
    USERCODE2      AS STRING * 4
    MFEE1          AS DOUBLE
    MFEE2          AS DOUBLE

    Serv(1 TO 15)     AS ServicesType
    LocMeters(1 TO 7) AS LocMeterType
    FlatRates(1 TO 4) AS FlatRateType

    OwnerRec          AS LONG          ' 5
    PrevLocRec        AS LONG          ' 9
    NextLocRec        AS LONG          ' 13
    CustRec           AS LONG          ' 17
    PendCustRec       AS LONG          ' 21
    OwnerPIN          AS LONG          ' 25
    LocaPIN           AS LONG          ' 29
    NotesRec          AS LONG          ' 33
    DelFlag           AS INTEGER
    EstFlag           AS STRING * 1
'add message
    MessageRec        AS INTEGER
    Pad               AS STRING * 1   ' 48

END TYPE

