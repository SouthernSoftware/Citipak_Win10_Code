DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'

  CLS

  REDIM TaxCust(1) AS TaxCustType
  TaxCustRecLen = LEN(TaxCust(1))

  OPEN "TaxCust.DAT" FOR RANDOM SHARED AS #1 LEN = TaxCustRecLen

  NumOfRec& = LOF(1) / TaxCustRecLen

  PRINT "Processing: 1205"
  GET #1, 1205, TaxCust(1)
  TaxCust(1).Deleted = 0
  PUT #1, 1205, TaxCust(1)
  CLOSE
  PRINT
  PRINT

  PRINT "Done."
  CALL KILLFile("FIXEMUP.EXE")

