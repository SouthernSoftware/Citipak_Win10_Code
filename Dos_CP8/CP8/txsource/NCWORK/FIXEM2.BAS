
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  FOR Cnt& = 1 TO LOF(PropFile) / PropRecLen
    GET PropFile, Cnt&, PropRec(1)
    PropRec(1).Deleted = 0
    PUT PropFile, Cnt&, PropRec(1)
  NEXT
  CLOSE

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  FOR Cnt& = 1 TO LOF(PersFile) / PersRecLen
    GET PersFile, Cnt&, PersRec(1)
    PersRec(1).Deleted = 0
    PUT PersFile, Cnt&, PersRec(1)
  NEXT
  CLOSE

