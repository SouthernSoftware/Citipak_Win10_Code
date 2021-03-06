TYPE DistType
    DACN      AS STRING * 16
    DACNM     AS STRING * 20
    DACREC    AS INTEGER
    DAMT      AS DOUBLE
END TYPE

TYPE CJEditRecType
    DELFLAG   AS INTEGER
    TRDATE    AS INTEGER
    DESC      AS STRING * 20
    DOCREF    AS STRING * 8
    AMT       AS DOUBLE
    BATCHNUM  AS STRING * 8
    RECCODE   AS INTEGER
    DIST(1 TO 36)  AS DistType
END TYPE

