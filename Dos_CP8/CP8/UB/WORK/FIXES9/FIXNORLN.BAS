DEFINT A-Z
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'

  CLS
  PRINT "WARNING ALL OPERATORS MUST EXIT UNTIL THIS PROGRAM IS FINISHED!!!"
  PRINT
  LOCATE 4, 1
  PRINT "Press any key when ready.";

  DO
  LOOP UNTIL LEN(INKEY$) > 0

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  IdxRecLen = 4 'we are using a long integer
  OPEN "UBCUSTBK.IDX" FOR RANDOM AS #1 LEN = IdxRecLen
  IdxFileSize& = LOF(1)
  CLOSE #1
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  'NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO IdxNumOfRecs
    LOCATE 4, 1
    PRINT "Processing:"; Cnt&; "of"; IdxNumOfRecs;
    GET UBFile, Cnt&, UBCustRec(1)
    FOR ZZ = 1 TO 7
      PrevRead& = UBCustRec(1).LocMeters(ZZ).PrevRead
      IF PrevRead& > 0 THEN
        PrevRead& = (PrevRead& / 100&)
        UBCustRec(1).LocMeters(ZZ).PrevRead = PrevRead&
        UBCustRec(1).LocMeters(ZZ).MTRMulti = 1000
      END IF
      CurrRead& = UBCustRec(1).LocMeters(ZZ).CurRead
      IF CurrRead& > 0 THEN
        CurrRead& = (CurrRead& / 100&)
        UBCustRec(1).LocMeters(ZZ).CurRead = CurRead&
      END IF

    NEXT

    PUT UBFile, IdxBuff(Cnt&).RecNum, UBCustRec(1)
  NEXT
  CLOSE
  ERASE IdxBuff

  LOCATE 6, 1
  PRINT "Done."
  CALL KillFile("FIXNORLN.EXE")

