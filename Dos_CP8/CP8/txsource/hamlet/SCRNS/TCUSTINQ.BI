TYPE TCUSTINQ
    ACCT     AS LONG
    OPENDATE AS INTEGER
    CUSTNAME AS STRING * 35
    HPHONE   AS STRING * 14
    CSSN     AS STRING * 11
    WPHONE   AS STRING * 14
    ADDR1    AS STRING * 35
    ADDR2    AS STRING * 35
    CITY     AS STRING * 20
    STATE    AS STRING * 2
    ZIP      AS STRING * 10
    ACTIVE   AS STRING * 1
    INTEREST AS STRING * 1
    EXEMPT   AS STRING * 1
    PENALTY  AS STRING * 1
    SRCITDIS AS DOUBLE
    ODISCOUN AS DOUBLE
    PROPVALU AS DOUBLE
    PERSVAL  AS DOUBLE
    MHVALUE  AS DOUBLE
    MCVALUE  AS DOUBLE
    CVALUE   AS DOUBLE
    MTVALUE  AS DOUBLE
    BALANCE  AS DOUBLE
END TYPE
