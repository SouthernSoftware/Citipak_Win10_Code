TYPE TaxCustomerRecType
     Account    AS STRING * 12
     NameLine1  AS STRING * 35
     NameLine2  AS STRING * 35
     Address1   AS STRING * 35
     Address2   AS STRING * 35
     State      AS STRING * 2
     ZipCode    AS STRING * 10
     Telephone  AS STRING * 12
     PropDesc1  AS STRING * 35
     PropDesc2  AS STRING * 35
     PropDesc3  AS STRING * 35
     PropDesc4  AS STRING * 35
     MapNumber  AS STRING * 35
     NumberLots AS SINGLE
     Acres      AS SINGLE
     LandValue  AS DOUBLE
     BldgValue  AS DOUBLE
     MobileHome AS DOUBLE
     FarmMach   AS DOUBLE
     LandUse    AS DOUBLE
     Personal   AS DOUBLE
     MachTools  AS DOUBLE
     Exemption  AS DOUBLE
     TaxType    AS STRING * 1
     UseType    AS STRING * 3
     CurLateAmt AS SINGLE
     CurTaxAmt  AS SINGLE
     CurIntAmt  AS SINGLE
     CurAdvAmt  AS SINGLE
     CurColAmt  AS SINGLE
     CurTaxNotice AS SINGLE
     FirstTrans AS SINGLE
     LastTrans  AS SINGLE
     SocSec     AS STRING * 11
     LateFlag   AS STRING * 1     'Y/N
     PIN        AS STRING * 15
     StormFee   AS DOUBLE
     pad        AS STRING * 18
END TYPE

TYPE CtyTaxRecType
     Book       AS STRING * 2
     Acct       AS STRING * 5
     Name       AS STRING * 34
     Name2      AS STRING * 16
     Add1       AS STRING * 34
     Add2       AS STRING * 28
     State      AS STRING * 2
     Zip        AS STRING * 5
     Deed       AS STRING * 8
     PropDesc1  AS STRING * 34
     PropDesc2  AS STRING * 34
     Size       AS STRING * 9
     PropDesc3  AS STRING * 46
     LandVal    AS STRING * 8
     BldgVal    AS STRING * 8
     UseCode    AS STRING * 3
     Exp        AS STRING * 2
     UseType    AS STRING * 1
     SD         AS STRING * 2
     MktVal     AS STRING * 8
     pad        AS STRING * 2
END TYPE



