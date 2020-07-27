TYPE FinalCurrPrevType
    MeterNum   AS STRING * 12
    MeterType  AS STRING * 7
    MeterUnit  AS STRING * 7
    CurrRead   AS LONG
    PrevRead   AS LONG
END TYPE

TYPE UBFinalReadType
    CUSTREC  AS LONG
    CUSTNAME AS STRING * 35
    READDATE AS INTEGER
    BOOK     AS STRING * 2
    SEQNUM   AS STRING * 6
    CurrPrev(1 TO 7) AS FinalCurrPrevType
END TYPE

