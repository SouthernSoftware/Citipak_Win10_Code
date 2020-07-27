DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  CLS
  LOCATE 2, 1
  PRINT "Correcting Transaction Dates. . ."

  BadDate = Date2Num("11-14-2025")
  FixDate = Date2Num("11-14-2005")

  '$INCLUDE: 'TaxCust.BI'

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumTr& = LOF(TaxTran) / TaxTranRecLen

  FOR Cnt& = 1 TO NumTr&
    GET TaxTran, Cnt&, TaxTran(1)
    
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; NumTr&;
    IF TaxTran(1).TransDate = BadDate THEN
      TaxTran(1).TransDate = FixDate
      PUT TaxTran, Cnt&, TaxTran(1)
      BADCNT = BADCNT + 1
    END IF
SkipFix:
  NEXT
  CLOSE
  CALL KILLFile("FIXTDATE.EXE")
  LOCATE 9, 1
  PRINT "Corrected:"; BADCNT
  PRINT
  PRINT
  PRINT "Done."
  'PRINT "Press any key to exit. . .";
  'DO
  '  BEEP
  '  A$ = INKEY$
  'LOOP UNTIL LEN(A$) > 0





