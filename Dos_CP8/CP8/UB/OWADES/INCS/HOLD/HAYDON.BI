TYPE CompanyName
        CompanyName     AS STRING * 30
        Address1        AS STRING * 30
        Address2        AS STRING * 30
        Pading          AS STRING * 38
END TYPE

TYPE Item
       Class            AS INTEGER
       Invoice          AS STRING * 12
       Style            AS SINGLE
       Desc             AS STRING * 40
       UnitCost         AS DOUBLE
       UnitRetail       AS DOUBLE
       QuanRecd         AS DOUBLE
       QuanSoldYTD      AS DOUBLE
       QuanSoldTP       AS DOUBLE
       RetailSalesYTD   AS DOUBLE
       RetailSalesTP    AS DOUBLE
       VendorName       AS STRING * 30
       VendorNumber     AS INTEGER
       DateRecd         AS STRING * 6
       CostofSales      AS DOUBLE
       TotalCost        AS DOUBLE
       FirstTr          AS LONG
       LastTr           AS LONG
       RoomToGrow       AS STRING * 336
END TYPE

TYPE ItemEdit
       Class            AS INTEGER
       Invoice          AS STRING * 12
       Style            AS SINGLE
       Desc             AS STRING * 40
       UnitCost         AS DOUBLE
       UnitRetail       AS DOUBLE
       QuanRecd         AS DOUBLE
       QuanSoldYTD      AS DOUBLE
       QuanSoldTP       AS DOUBLE
       RetailSalesYTD   AS DOUBLE
       RetailSalesTP    AS DOUBLE
       VendorName       AS STRING * 30
       VendorNumber     AS INTEGER
       DateRecd         AS STRING * 6
       CostofSales      AS DOUBLE
       TotalCost        AS DOUBLE
       FirstTr          AS LONG
       LastTr           AS LONG
       RoomToGrow       AS STRING * 336
END TYPE

TYPE AdjEdit
       Invoice          AS STRING * 12
       Style            AS SINGLE
       Desc             AS STRING * 25
       ItemRecord       AS SINGLE
       AdjType          AS STRING * 1   '(U)p  or (D)own
       UnitCost         AS DOUBLE
       UnitRetail       AS DOUBLE
       QuanRecd         AS DOUBLE
       DateRecd         AS STRING * 6
       Pad              AS STRING * 50
END TYPE

TYPE EditTrans
       TranType         AS STRING * 1   'S=Sell  R=Receive  A=Adjustment
       Date             AS STRING * 6
       AmountSale       AS DOUBLE
       AmountCost       AS DOUBLE
       Quan             AS DOUBLE
       Ticket           AS STRING * 12
       Salesman         AS SINGLE
       PayCode          AS INTEGER
       Taxable          AS STRING * 1  'Y=Yes  N=No  D=Transaction Deleted
       ItemRecord       AS SINGLE
       TypeSale         AS INTEGER
       CustName         AS STRING * 30
       Pad              AS STRING * 20
END TYPE


TYPE ItemTrans
       TranType         AS STRING * 1   'S=Sell  R=Receive  A=Adjustment
       Date             AS STRING * 6
       AmountSale       AS DOUBLE
       AmountCost       AS DOUBLE
       Quan             AS DOUBLE
       Ticket           AS STRING * 12
       Salesman         AS SINGLE
       PayCode          AS INTEGER
       Taxable          AS STRING * 1
       ItemRecord       AS SINGLE
       TypeSale         AS INTEGER
       NextTr           AS LONG
       CustName         AS STRING * 30
       Padding          AS STRING * 38
END TYPE

TYPE Class
       ClassName        AS STRING * 40
       Pad              AS STRING * 24
END TYPE

TYPE Salesman
       SalesName        AS STRING * 30
       DolSalesYtd      AS DOUBLE
       NumSalesYtd      AS SINGLE
       DolSalesTW       AS DOUBLE
       NumSalesTW       AS SINGLE
       DolSalesToday    AS DOUBLE
       NumSalesToday    AS SINGLE
       Pad              AS STRING * 24
END TYPE
TYPE ApVendor
   venname AS STRING * 30
   address AS STRING * 30
   City  AS STRING * 22
   St AS STRING * 2
   Zip AS STRING * 10
   Payto AS STRING * 30
   PayAdd AS STRING * 30
   PayCity AS STRING * 22
   PaySt AS STRING * 2
   PayZip AS STRING * 10
   Curbal AS DOUBLE
   FirstTr AS SINGLE
   LastTr AS SINGLE
   HPhone AS STRING * 10
   WPhone AS STRING * 10
   Contact AS STRING * 30
   APPad AS STRING * 258
END TYPE

TYPE APTrans
 Vennum AS INTEGER
 InvDate AS STRING * 8
 InvNum AS STRING * 10
 PONum AS STRING * 10
 Amount AS DOUBLE
 PayCode AS INTEGER
 Fdist AS SINGLE
 Ldist AS SINGLE
 NData AS SINGLE
END TYPE

TYPE APTrans1
 ApDistDate AS STRING * 8
 ApDueDate AS STRING * 8
END TYPE





TYPE DaySales
   TaxSales AS SINGLE
   CreditSales AS SINGLE
   TaxSalesTax AS SINGLE
   CreditSalesTax AS SINGLE
   NonTaxMerch AS SINGLE
   NonTaxLabor AS SINGLE
   Postage AS SINGLE
   ChargeSale AS SINGLE
   RecdonAcct AS SINGLE
   CreditAcct AS SINGLE
   LessDep AS SINGLE
   LayPaid AS SINGLE
   LayPayment AS SINGLE
   Refund AS SINGLE
   PaidOut AS SINGLE
   Amex AS DOUBLE
   Cash AS DOUBLE
   Charge AS DOUBLE
   Posted AS STRING * 1         'Y/N  Allow clearing a 'Y' only
   NotUsed AS STRING * 19
END TYPE

TYPE GLDefAcct
   AccountNumber AS STRING * 9
   AccountType AS STRING * 1            'D or C for debit/credit
   Pad AS STRING * 30
END TYPE

TYPE BAAcct
        mastacct AS STRING * 9
        ACCTNAME AS STRING * 25
        ACCTTYPE AS STRING * 1
        acctbal  AS DOUBLE
        mtdact AS DOUBLE
        ytdact AS DOUBLE
        budamt AS DOUBLE
        FIRST AS SINGLE
        LAST AS SINGLE
        begbal AS DOUBLE
        RecalcBal AS DOUBLE
        Padding AS STRING * 9
END TYPE

TYPE BATrans
        TRANACCT AS STRING * 9
        TRANDATE AS STRING * 8
        TRANDESC AS STRING * 20
        CRAMT AS DOUBLE
        DRAMT AS DOUBLE
        DETAIL AS STRING * 8
        SOURCE AS STRING * 8
        NextTr AS SINGLE
        Padding AS STRING * 12
END TYPE

TYPE Customer
        LastName AS STRING * 25
        FirstName AS STRING * 25
        Address1 AS STRING * 30
        Address2 AS STRING * 30
        Address3 AS STRING * 30
        Phone AS STRING * 12
        Purchases AS DOUBLE
        LastPurchaseDate AS STRING * 8
        CustType AS STRING * 2
        Padding AS STRING * 75
END TYPE

