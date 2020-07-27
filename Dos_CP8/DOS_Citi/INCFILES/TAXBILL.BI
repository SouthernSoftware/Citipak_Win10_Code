TYPE TaxBillType
     CustRec            AS LONG                 'Acct #
     CustName           AS STRING * 40
     CustAdd1           AS STRING * 35
     CustAdd2           AS STRING * 35
     CustAdd3           AS STRING * 35
     CustZip            AS STRING * 10
     RDesc1             AS STRING * 30
     RDesc2             AS STRING * 30
     RealPIN            AS STRING * 16
     RealValue          AS DOUBLE
     PersValue          AS DOUBLE
     ExptValue          AS DOUBLE
     RealTaxDue         AS DOUBLE
     PersTaxDue         AS DOUBLE
     LateTaxDue         AS DOUBLE
     TotalBillDue       AS DOUBLE
     BillNumber         AS LONG                          'Recpt #
     TaxYear            AS INTEGER
     BillPrinted        AS INTEGER            '-1 = printed
     RealPropRecord     AS LONG
     PersPropRecord     AS LONG
     PriorYrBalance     AS DOUBLE
     RealTaxRate        AS DOUBLE
     PersTaxRate        AS DOUBLE
     CustPin            AS LONG         'additional Protection for relinking
     TownShip           AS STRING * 20
     MortCode           AS STRING * 2
     LotOrAcre          AS STRING * 1
     LASize             AS STRING * 9
     MortRec            AS INTEGER
     CarShore           AS DOUBLE
     RDesc3             AS STRING * 30
     Padding            AS STRING * 105
END TYPE


TYPE VAPPTaxBillType
     CustRec            AS LONG                 'Acct #
     CustName           AS STRING * 40
     CustAdd1           AS STRING * 35
     CustAdd2           AS STRING * 35
     CustAdd3           AS STRING * 35
     CustZip            AS STRING * 10
     RDesc1             AS STRING * 30
     RDesc2             AS STRING * 30
     RealPIN            AS STRING * 16
     PersValue          AS DOUBLE
     MHValue            AS DOUBLE
     MCValue            AS DOUBLE
     FEValue            AS DOUBLE
     MTValue            AS DOUBLE
     ExptValue          AS DOUBLE
     PersTaxDue         AS DOUBLE
     MHTaxDue           AS DOUBLE
     MCTaxDue           AS DOUBLE
     FETaxDue           AS DOUBLE
     MTTaxDue           AS DOUBLE
     LateTaxDue         AS DOUBLE
     TotalBillDue       AS DOUBLE
     BillNumber         AS LONG         'Recpt #
     TaxYear            AS INTEGER
     BillPrinted        AS INTEGER      '-1 = printed
     PersPropRecord     AS LONG
     PriorYrBalance     AS DOUBLE
     PersTaxRate        AS DOUBLE
     CustPin            AS LONG         'Same as Record #
     ChillHowieFudge    AS SINGLE
     PPTRAValue         AS DOUBLE
     PPTRADiscnt        AS DOUBLE
     Padding            AS STRING * 157
END TYPE

