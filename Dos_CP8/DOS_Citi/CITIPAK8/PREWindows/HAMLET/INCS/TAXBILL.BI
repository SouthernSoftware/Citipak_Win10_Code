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

     Padding            AS STRING * 177
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
     Padding            AS STRING * 173
END TYPE

