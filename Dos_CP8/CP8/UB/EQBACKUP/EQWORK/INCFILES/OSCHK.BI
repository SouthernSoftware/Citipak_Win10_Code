TYPE OSChkRecType
   ChkNum   AS SINGLE        '4 AS chknum$
   ChkDate  AS STRING * 8    '8 AS chkdate$
   Desc     AS STRING * 30   '30 AS chkdesc$
   Amt      AS SINGLE        '4 AS chkamt$
   Src      AS INTEGER       '2 AS CHKSOURCE$
   Cleared  AS INTEGER       'added by JB
   BankCode AS INTEGER
   Filler   AS STRING * 12   '16 AS nul$
END TYPE

TYPE OSChkRecType1
   ChkNum   AS LONG        '4 AS chknum$
   ChkDate  AS STRING * 8    '8 AS chkdate$
   Desc     AS STRING * 30   '30 AS chkdesc$
   Amt      AS DOUBLE        '4 AS chkamt$
   Src      AS INTEGER       '2 AS CHKSOURCE$
   Cleared  AS INTEGER       'added by JB
   BankCode AS INTEGER
   Filler   AS STRING * 8   '16 AS nul$
END TYPE


