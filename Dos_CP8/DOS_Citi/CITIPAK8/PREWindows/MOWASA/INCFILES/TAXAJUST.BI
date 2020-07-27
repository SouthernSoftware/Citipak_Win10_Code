TYPE TaxAdjustRecType
    AdjDate        AS INTEGER
    CustAcct       AS LONG
    CustName       AS STRING * 24
    CustAddr       AS STRING * 24
    AdjNote        AS STRING * 20
    BalAmt(1 TO 3) AS DOUBLE
    AdjAmt(1 TO 3) AS DOUBLE
    TAMTADJ        AS DOUBLE
END TYPE

