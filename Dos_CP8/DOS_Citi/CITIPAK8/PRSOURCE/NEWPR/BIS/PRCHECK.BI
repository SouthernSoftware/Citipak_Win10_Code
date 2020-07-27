TYPE PRDEDType
   DCode       AS STRING * 10
   DAmt        AS DOUBLE
'added
   YTDDAmt     AS DOUBLE
'--------
END TYPE

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

TYPE PRCheckRecType
   CActive       AS INTEGER
   CheckNum      AS LONG
   CheckDate     AS INTEGER

   EmpName       AS STRING * 33
   EmpNo         AS STRING * 10
   EmpSSN        AS STRING * 11

'=-=-=-=-=-
   EmpAddr1 AS STRING * 36
   EmpCity  AS STRING * 24
   EmpState AS STRING * 2
   EmpZip   AS STRING * 10
'-=-=-=-=-=-=

   PayEndDate    AS INTEGER
   BaseRate      AS DOUBLE
   GrossPay      AS DOUBLE
   FedTaxAmt     AS DOUBLE
   StaTaxAmt     AS DOUBLE
   MedTaxAmt     AS DOUBLE
   SocTaxAmt     AS DOUBLE
   TotDedAmt     AS DOUBLE

'added
   RetireAmt     AS DOUBLE

   NetPay        AS DOUBLE
   YTDGrossPay   AS DOUBLE
   YTDFederal    AS DOUBLE
   YTDState      AS DOUBLE
   YTDSocial     AS DOUBLE
   YTDMedicare   AS DOUBLE
   YTDTotDed     AS DOUBLE
   YTDNetPay     AS DOUBLE

'added
   YTDRetire     AS DOUBLE

   VactBal       AS DOUBLE   '
   SickBal       AS DOUBLE   '
   CompBal       AS DOUBLE

'added-----
   CompEarn      AS DOUBLE
   RegHrsWork    AS DOUBLE
   OTHrsPaid     AS DOUBLE
   TotRegWage    AS DOUBLE
   VacUsed       AS DOUBLE
   SickUsed      AS DOUBLE
   CompUsed      AS DOUBLE
   HolUsed       AS DOUBLE
   PerUsed       AS DOUBLE

   RegHrsPaid    AS DOUBLE
   TotOTWage     AS DOUBLE

   AEarn(1 TO 3) AS PRDEDType

   TotAdditEarn  AS DOUBLE

   EICAmt        AS DOUBLE
   TaxFring      AS DOUBLE

'----------
   CDED(1 TO 12) AS PRDEDType
   DDFlag        AS INTEGER
END TYPE

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

