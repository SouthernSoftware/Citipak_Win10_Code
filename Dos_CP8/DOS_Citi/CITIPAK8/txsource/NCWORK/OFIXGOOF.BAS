DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'

  CLS

  REDIM TaxCust(1 TO 2) AS TaxCustType
  TaxCustRecLen = LEN(TaxCust(1))

  OPEN "TaxCust.DAT" FOR RANDOM SHARED AS #1 LEN = TaxCustRecLen

  NumOfRec& = LOF(1) / TaxCustRecLen

  FOR Cnt& = 1 TO NumOfRec&
    'LOCATE 5, 1: PRINT "Processing:"; Cnt&;

    GET #1, Cnt&, TaxCust(1)
    PRINT TaxCust(1).CSSN

'    PRINT TaxCust(1).TownShip
'SLEEP 1

  '  GET #2, Cnt&, TaxCust(2)
  '  TaxCust(2).sSSN = TaxCust(1).sSSN
  '  TaxCust(2).CSSN = TaxCust(1).CSSN

    Fixed = Fixed + 1
   ' PUT #2, Cnt&, TaxCust(2)
  NEXT
  CLOSE
  PRINT
  PRINT

 ' PRINT " Corrected:"; Fixed
  CALL KILLFile("FIXGOOF.EXE")

