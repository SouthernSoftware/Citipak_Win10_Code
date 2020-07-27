TYPE oItemRecType
    ITEMNUMBER AS LONG
    DESC1      AS STRING * 45
    DESC2      AS STRING * 45
    Vendor     AS STRING * 30
    VItem      AS STRING * 20
    DeptNumb   AS STRING * 8
    COST       AS DOUBLE
    Unit       AS STRING * 10
    QtyOnHand  AS DOUBLE
    QtyOnOrd   AS DOUBLE
    UsedYTD    AS DOUBLE
    ReOrder    AS DOUBLE
    QTD        AS DOUBLE
    Note1      AS STRING * 45
    Note2      AS STRING * 45
    Deleted    AS STRING * 1              'Y if Deleted
    LastTran   AS LONG                    'pointer to last transaction
    Future     AS STRING * 367
END TYPE

TYPE ItemRecType
    ITEMNUMBER AS LONG
    DESC1      AS STRING * 45
    DESC2      AS STRING * 45
    Vendor     AS STRING * 30
    VItem      AS STRING * 20
    DeptNumb   AS STRING * 8
    COST       AS DOUBLE
    YTDCost    AS DOUBLE
    Unit       AS STRING * 10
    QtyOnHand  AS DOUBLE
    QtyOnOrd   AS DOUBLE
    UsedYTD    AS DOUBLE
    ReOrder    AS DOUBLE
    QTD        AS DOUBLE
    Note1      AS STRING * 45
    Note2      AS STRING * 45
    Deleted    AS STRING * 1              'Y if Deleted
    LastTran   AS LONG                    'pointer to last transaction
    Future     AS STRING * 359
END TYPE

TYPE ICTransRecType       '1=Up Adj 2=Dwn Adj 3=Redorder 4=Used 5=Purchased
   ItemRec          AS LONG
   TransDate        AS INTEGER       'date user entered as trans date
   TransType        AS INTEGER
   TransQuantity    AS DOUBLE
   TransCost        AS DOUBLE
   TransNote        AS STRING * 40
'internal fields
   EntryDate        AS INTEGER       'actual trans entry creation date
   PostedDate       AS INTEGER       'date trans was really posted
   PrevTran         AS LONG          'pointer to previous transaction
   DelFlag          AS INTEGER
   TransPad         AS STRING * 22
END TYPE

'TYPE TRANS
'    ITEMNUMB AS LONG
'    DATEIN   AS INTEGER
'    TRANTYPE AS STRING * 20
'    Quantity AS DOUBLE
'    COST     AS DOUBLE
'    TRANSDES AS STRING * 40
'    DESC1    AS STRING * 45
'    DESC2    AS STRING * 45
'END TYPE

