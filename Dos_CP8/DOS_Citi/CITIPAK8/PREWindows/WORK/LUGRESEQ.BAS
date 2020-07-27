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
    GET UBFile, IdxBuff(Cnt&).RecNum, UBCustRec(1)
    Book$ = QPTrim$(UBCustRec(1).Book)
    Loca$ = QPTrim$(UBCustRec(1).SeqNumb)
    UpDateEM = 0
    IF (LEN(Book$) = 2) AND (LEN(Loca$) = 6) THEN
      UpDateEM = -1
      SeqS$ = Book$ + Loca$
      Seq& = VAL(SeqS$)
      SELECT CASE Seq&
      CASE 10000100 TO 10022750
        UBCustRec(1).Book = "01"
      CASE 10022800 TO 10048950
        UBCustRec(1).Book = "02"
      CASE 10049000 TO 10066000
        UBCustRec(1).Book = "03"
      CASE 10066050 TO 10089220
        UBCustRec(1).Book = "04"
      CASE 10089250 TO 10100187
        'no action for this group
      CASE 10100200 TO 10160030
        UBCustRec(1).Book = "05"
      CASE 20600010 TO 20636800
        UBCustRec(1).Book = "06"
      CASE 20636900 TO 20672820
        UBCustRec(1).Book = "07"
      CASE 20672900 TO 20705700
        UBCustRec(1).Book = "08"
      CASE 20705800 TO 20766700
        UBCustRec(1).Book = "09"
      END SELECT
    END IF
    FOR ZZ = 1 TO 7
      CurrRead& = UBCustRec(1).LocMeters(ZZ).CurRead
      IF PrevRead& > 0 THEN
        CurrRead& = (CurrRead& / 100&)
        UBCustRec(1).LocMeters(ZZ).CurRead = CurrRead&
        UBCustRec(1).LocMeters(ZZ).MTRMulti = 100
      END IF
    NEXT

    PUT UBFile, IdxBuff(Cnt&).RecNum, UBCustRec(1)
  NEXT
  CLOSE
  ERASE IdxBuff

  LOCATE 6, 1
  PRINT "Done."
  LOCATE 10, 1
  PRINT "Press any ket when ready.";
  DO
  LOOP UNTIL LEN(INKEY$) > 0
  PRINT
  PRINT "Reindexing data. . . "
  'END

  SHELL "UBMISC REINDEX"
  CALL KillFile("LUGRESEQ.EXE")

