DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
'DECLARE FUNCTION Date2Num (TheDate$)
'DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  CLS

  PRINT
  PRINT
  PRINT "Reading meter number file. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  UBMtrF = FREEFILE
  OPEN "newMNumb.txt" FOR INPUT AS UBMtrF
  ReadCnt& = 0
  ON ERROR GOTO BadFormat
  DO UNTIL EOF(UBMtrF)
    ReadCnt& = ReadCnt& + 1
    LOCATE , 1
    PRINT "Reading:"; ReadCnt&;
    INPUT #UBMtrF, Acct$, NewMtrNum$, OldMtrNum$
    CustAcct& = VAL(Acct$)
    IF CustAcct& < 1 OR CustAcct& > NumOfRecs& THEN
      BadCust = -1
    END IF
  LOOP
  CLOSE UBMtrF

  IF BadCust = -1 THEN
    GOTO CustBad
  END IF

  'END
  PRINT
  PRINT
  PRINT

  ReadCnt& = 0
  UBMtrF = FREEFILE
  OPEN "newMNumb.txt" FOR INPUT AS UBMtrF
  DO UNTIL EOF(UBMtrF)
    ReadCnt& = ReadCnt& + 1
    LOCATE , 1
    PRINT "Processing:"; ReadCnt&;
    INPUT #UBMtrF, Acct$, NewMtrNum$, OldMtrNum$
    CustAcct& = VAL(Acct$)
    GET UBFile, CustAcct&, UBCustRec(1)
    UBCustRec(1).LocMeters(6).MtrNum = UBCustRec(1).LocMeters(5).MtrNum
    UBCustRec(1).LocMeters(5).MtrNum = UBCustRec(1).LocMeters(4).MtrNum
    UBCustRec(1).LocMeters(4).MtrNum = UBCustRec(1).LocMeters(3).MtrNum
    UBCustRec(1).LocMeters(3).MtrNum = UBCustRec(1).LocMeters(2).MtrNum
    UBCustRec(1).LocMeters(2).MtrNum = UBCustRec(1).LocMeters(1).MtrNum
    ThisMtr$ = QPTrim$(NewMtrNum$) + "Z"
    UBCustRec(1).LocMeters(1).MtrNum = ThisMtr$
    PUT UBFile, CustAcct&, UBCustRec(1)
  LOOP
  CLOSE
  PRINT
  PRINT
  PRINT

  PRINT "Done."
  KILL "newMNumb.txt"
  END

'  KillFile "FIXCALD.EXE"

CustBad:

  CLOSE
  PRINT "Bad Customer Account"
  END


BadFormat:
  CLOSE
  PRINT "Bad File Format!!!!!"
  END

