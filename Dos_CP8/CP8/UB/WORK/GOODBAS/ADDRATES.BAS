DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  ''$INCLUDE: 'DefCnf.BI'
  ''$INCLUDE: 'formedit.BI'
  ''$INCLUDE: 'fieldinf.BI'
  ''$INCLUDE: 'PageInfo.BI'                  'Form Page info
  ''$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  ''$INCLUDE: 'ubsetup.bi'
  ''$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'Newcust.bi'
  ''$INCLUDE: 'ubrate.bi'
  ''$INCLUDE: 'ubTrans.bi'

  CLS

  PRINT "Updating Utility Customer Information. . ."

  CONST False = 0, True = NOT False
  
  TCode$ = "WRI1"
  Code7$ = "LIMB"
  Code8$ = "ADMI"

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen

  FOR Cnt = 1 TO NumOfRecs
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; "of"; NumOfRecs;
    GET UBCust, Cnt, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      '*************************************
      '   Main body of Printing goes here
      FOR RCnt = 1 TO 15
        ThisCode$ = QPTrim$(UBCustRec(1).Serv(RCnt).RATECODE)
        IF LEN(ThisCode$) > 0 THEN
          IF TCode$ = ThisCode$ THEN
            CustomerCnt = CustomerCnt + 1
            UBCustRec(1).Serv(7).RATECODE = Code7$
            UBCustRec(1).Serv(8).RATECODE = Code8$
            PUT UBCust, Cnt, UBCustRec(1)
            EXIT FOR
          END IF
        END IF
SkipThisRev:
      NEXT
    END IF
  NEXT

  CLOSE
  LOCATE 7, 1
  PRINT "Changed:"; CustomerCnt
  LOCATE 9, 1
  PRINT "Done."
  CALL KillFile("addrates.exe")

END

