TYPE TblBreakRecType
    UNITS      AS LONG
    UNITAMT    AS DOUBLE
END TYPE

TYPE UBRateTblRecType
    RATECODE AS STRING * 4
    RATEDESC AS STRING * 29
    ChkByte  AS STRING * 1
    MINAMT   AS DOUBLE
    MINUNITS AS LONG
    MaxAmt   AS DOUBLE
    TblBreaks(1 TO 10) AS TblBreakRecType
END TYPE

TYPE oUBRateTblRecType
    RATECODE AS STRING * 4
    RATEDESC AS STRING * 30
    MINAMT   AS DOUBLE
    MINUNITS AS LONG
    DiscPct  AS INTEGER
    TblBreaks(1 TO 10) AS TblBreakRecType
END TYPE

TYPE BookConsumpType
    Book             AS INTEGER
    CustCnt          AS LONG
    Consump(1 TO 15) AS DOUBLE
    'Consump(1 TO 15) AS LONG
    RevAmt(1 TO 15)  AS DOUBLE
    TaxAmt(1 TO 15)  AS DOUBLE
END TYPE

TYPE PumpConsumpType
    PumpCode         AS STRING * 4
    CustCnt          AS LONG
    Consump         AS DOUBLE
    'Consump(1 TO 15) AS DOUBLE
END TYPE

