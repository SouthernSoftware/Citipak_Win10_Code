DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  CLS

  REDIM TaxCust(1) AS TaxCustType
  TaxCustRecLen = LEN(TaxCust(1))
'  KillFile "TaxCust.BAK"
'  NAME "TaxCust.DAT" AS "TaxCust.BAK"

'  OPEN "TaxCust.BAK" FOR RANDOM SHARED AS #1 LEN = TaxCustRecLen
  OPEN "TaxCust.DAT" FOR RANDOM SHARED AS #2 LEN = TaxCustRecLen

  NumOfRec& = LOF(2) / TaxCustRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #2, Cnt&, TaxCust(1)
    TaxCust(1).Interest = "Y"
    PUT #2, , TaxCust(1)
  NEXT
  CLOSE
  PRINT
  PRINT

  PRINT "Done."
  'KILL "FIXCUST.EXE"

