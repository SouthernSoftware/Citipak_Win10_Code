DEFINT A-Z
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)

'$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS

  DIM ARCust AS ARCustRecType
  CustLen = LEN(ARCust)
  CLS
  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = CustLen
  NumOfCust = LOF(1) \ CustLen
  FOR Cnt = 1 TO NumOfCust
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; " of"; NumOfCust;
    GET #1, Cnt, ARCust
    ExpDate$ = Num2Date(ARCust.Valid)
    IF RIGHT$(ExpDate$, 4) = "2000" THEN
      NewDate$ = LEFT$(ExpDate$, 6) + "2001"
      ARCust.Valid = Date2Num(NewDate$)
      Fixed = Fixed + 1
      PUT #1, Cnt, ARCust
    END IF
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Corrected:"; Fixed

