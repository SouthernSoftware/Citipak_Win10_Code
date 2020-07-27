DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  WhatDate = Date2Num("06-30-2002")

  CLS
  DIM SHARED ARCustRec(1) AS ARCustRecType
  PRINT "Processing. . ."
  CustLen = LEN(ARCustRec(1))
  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = CustLen
  NumCust = LOF(1) / CustLen
  FOR zz = 1 TO NumCust

    LOCATE 5, 1
    PRINT "Customer:"; zz; "of"; NumCust;
    GET #1, zz, ARCustRec(1)
    ARCustRec(1).LicBal = 0
    ARCustRec(1).PenBal = 0
    ARCustRec(1).FeeBal = 0
    ARCustRec(1).ProRate = 100
    ARCustRec(1).AcctBal = 0
    ARCustRec(1).IssueLicense = ""
    ARCustRec(1).Inactive = "N"
    ARCustRec(1).FirstTrans = 0
    ARCustRec(1).LastTrans = 0
    ARCustRec(1).IssuanceFee = 0
    ARCustRec(1).Valid = WhatDate
    PUT #1, zz, ARCustRec(1)

  NEXT

  CLOSE
  OPEN "artrans.dat" FOR OUTPUT AS #1
  CLOSE
  LOCATE 9, 1
  PRINT "Processing Complete."
  CALL KillFile("FixBal.exe")

