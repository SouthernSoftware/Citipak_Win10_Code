DEFINT A-Z

  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'PROPAbst.BI'

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))
  

  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  NumOfRec& = LOF(PropFile) / PropRecLen

  FOR CNT& = 1 TO (NumOfRec& - 1)
    GET #PropFile, CNT&, PropRec(1)
     PropRec(1).LastYrPrinted = 1999
    PUT #PropFile, CNT&, PropRec(1)

  NEXT

   
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  NumOfRec& = LOF(PersFile) / PersRecLen
  FOR CNT& = 1 TO (NumOfRec& - 4)
    GET #PersFile, CNT&, PersRec(1)
    PersRec(1).LastYrPrinted = 1999
    PUT #PersFile, CNT&, PersRec(1)
   
  NEXT CNT&
  CLOSE

  END



