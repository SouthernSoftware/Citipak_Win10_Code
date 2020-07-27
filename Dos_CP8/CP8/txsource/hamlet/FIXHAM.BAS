DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxBill.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXBINFO.BI'

  REDIM TaxCust(1)  AS TaxCustType
  REDIM RealRec(1)  AS PropertyRecType
  REDIM PersRec(1)  AS PersonalRecType

  RealRecLen = LEN(RealRec(1))
  PersRecLen = LEN(PersRec(1))
  TaxCustLen = LEN(TaxCust(1))

  'Open All Needed Files Here
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen

  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  
  CustFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM AS CustFile LEN = TaxCustLen

  GET CustFile, 3854, TaxCust(1)
  IF TaxCust(1).FirstPropRec > 0 THEN
    GET RealFile, TaxCust(1).FirstPropRec, RealRec(1)
    RealRec(1).LastYrPrinted = 1999
    PUT RealFile, TaxCust(1).FirstPropRec, RealRec(1)
  END IF
  IF TaxCust(1).FirstPersRec > 0 THEN
    GET PersFile, TaxCust(1).FirstPersRec, PersRec(1)
    PersRec(1).LastYrPrinted = 1999
    PUT PersFile, TaxCust(1).FirstPersRec, PersRec(1)
  END IF

  GET CustFile, 3765, TaxCust(1)
  IF TaxCust(1).FirstPropRec > 0 THEN
    GET RealFile, TaxCust(1).FirstPropRec, RealRec(1)
    RealRec(1).LastYrPrinted = 1999
    PUT RealFile, TaxCust(1).FirstPropRec, RealRec(1)
  END IF
  IF TaxCust(1).FirstPersRec > 0 THEN
    GET PersFile, TaxCust(1).FirstPersRec, PersRec(1)
    PersRec(1).LastYrPrinted = 1999
    PUT PersFile, TaxCust(1).FirstPersRec, PersRec(1)
  END IF

  CLOSE

  kill "fixham.exe"
  END






