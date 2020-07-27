DEFINT A-Z
  
  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT
  PRINT
  PRINT "Processing Late Notice Flags. . ."
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS #1 LEN = TaxRecLen
  NumRec& = LOF(1) / TaxRecLen
  FOR Cnt& = 1 TO NumRec&
    LOCATE 6, 1: PRINT "Account:"; Cnt&; "of"; NumRec&;
    GET #1, Cnt&, TaxCust(1)
    TaxCust(1).LateNotice = "Y"
    PUT #1, Cnt&, TaxCust(1)
  NEXT
  CLOSE
  LOCATE 9, 1
  PRINT "Processing Complete."
  KILL "setlflag.exe"

