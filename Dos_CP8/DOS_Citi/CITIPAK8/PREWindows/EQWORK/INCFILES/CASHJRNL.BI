TYPE DistType
    DACN      AS STRING * 16
    DACNM     AS STRING * 20
    DACREC    AS INTEGER
    DAMT      AS DOUBLE
END TYPE

TYPE CJEditRecType
    DELFLAG   AS INTEGER
    'VENDOR CODE as string * 10
    'VDRRC     as integer
    TRDATE    AS INTEGER
    'INVDATE   as integer
    'DUEDATE   as integer
    DESC      AS STRING * 20
    DOCREF    AS STRING * 8
    AMT       AS DOUBLE
    BATCHNUM  AS STRING * 8
    RECCODE   AS STRING * 2 'AS INTEGER
    DIST(1 TO 36)  AS DistType
END TYPE

