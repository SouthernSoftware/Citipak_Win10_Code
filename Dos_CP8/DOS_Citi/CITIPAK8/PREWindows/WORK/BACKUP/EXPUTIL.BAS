DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  
  CONST False = 0, True = NOT False

  CLS
  PRINT "Exporting Utility Customer Information"

  q$ = CHR$(34)
  c$ = CHR$(44)

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))          'Length of Cust Record Structure

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  EXFile = FREEFILE
  OPEN "FROMUTIL.ASC" FOR OUTPUT AS EXFile

  FOR Cnt& = 1 TO NumOfRecs&
    GET UBFile, Cnt&, UBCustRec(1)
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    IF UBCustRec(1).DelFlag = False THEN
      IF INSTR(UBCustRec(1).CustName, "VACANT") = 0 THEN
        PRINT #EXFile, q$; QPTrim$(UBCustRec(1).CustName); q$; c$; q$; QPTrim$(UBCustRec(1).Addr1); q$; c$; q$; QPTrim$(UBCustRec(1).HPhone); q$
      END IF
    END IF
  NEXT
  CLOSE
LOCATE 10, 5
PRINT "Export Completed"

