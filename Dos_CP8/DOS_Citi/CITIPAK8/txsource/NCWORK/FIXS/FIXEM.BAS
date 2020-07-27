DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  PRINT "Correcting Invalid Transactions Date. . ."
  BadDate = Date2Num("07-01-2024")
  GoodDate = Date2Num("07-01-2004")

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen

  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = BadDate THEN
      TaxTran(1).TransDate = GoodDate
      Found = Found + 1
      PUT #1, Cnt&, TaxTran(1)
    END IF
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "FIXED:"; Found

  CALL KILLFile("fixem.exe")

