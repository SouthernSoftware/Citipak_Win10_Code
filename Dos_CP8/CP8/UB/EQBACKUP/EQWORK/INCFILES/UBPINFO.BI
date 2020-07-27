TYPE oPrintBillInfoType
    FrstBill AS LONG
    LastBill AS LONG
    BillDate AS INTEGER
    PastDate AS INTEGER
    PRDate   AS INTEGER
    CRDate   AS INTEGER
    DrftDate AS INTEGER
    PrnOrder AS STRING * 25
    MsgLine  AS STRING * 25
    Desc1    AS STRING * 15
    Desc2    AS STRING * 15
    Desc3    AS STRING * 15
END TYPE

TYPE PrintBillInfoType
    EPPYearEnd  AS STRING * 1
    FrstBill    AS LONG
    LastBill    AS LONG
    BillDate    AS INTEGER
    PastDate    AS INTEGER
    PRDate      AS INTEGER
    CRDate      AS INTEGER
    DrftDate    AS INTEGER
    PrnOrder    AS STRING * 25
    MsgLine1    AS STRING * 25
    MsgLine2    AS STRING * 25
    MsgLine3    AS STRING * 25
    MsgLine4    AS STRING * 25
END TYPE
 

