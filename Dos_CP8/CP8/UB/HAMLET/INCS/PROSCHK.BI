TYPE OSChkRecType
   ChkNum   AS SINGLE        '4 AS chknum$
   ChkDate  AS STRING * 8    '8 AS chkdate$
   Desc     AS STRING * 30   '30 AS chkdesc$
   Amt      AS SINGLE        '4 AS chkamt$
   Src      AS INTEGER       '2 AS CHKSOURCE$
   Cleared  AS INTEGER       'added by JB
   BankCode AS INTEGER       '09/07/96 Gate City's
   Filler   AS STRING * 12 '14   '16 AS nul$
END TYPE

