DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'

  CLS

  REDIM TaxCust(1) AS TaxCustType
  TaxCustRecLen = LEN(TaxCust(1))

  OPEN "TaxCust.DAT" FOR RANDOM SHARED AS #1 LEN = TaxCustRecLen

  NumOfRec& = LOF(1) / TaxCustRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxCust(1)
    Addr1$ = QPTrim$(TaxCust(1).Addr1)
    IF LEN(Addr1$) = 0 THEN
      SWAP TaxCust(1).Addr1, TaxCust(1).Addr2
      Fixed = Fixed + 1
    END IF
    PUT #1, Cnt&, TaxCust(1)
  NEXT
  CLOSE
  PRINT
  PRINT

  PRINT " Corrected:"; Fixed
  CALL KILLFile("FIXADDR.EXE")

