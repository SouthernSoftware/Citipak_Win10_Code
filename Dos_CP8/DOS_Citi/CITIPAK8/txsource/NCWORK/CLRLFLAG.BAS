DEFINT A-Z

  '$INCLUDE: 'bobTaxes.BI'
  
  DIM RealRec AS PropertyRecType
  DIM PersRec AS PersonalRecType

  RealRecLen = LEN(RealRec)
  PersRecLen = LEN(PersRec)
  
  CLS
  PRINT "Clearing Late Listing Flags. . ."
  PRINT
  PRINT "Realestate. . ."

  OPEN "TaxProp.dat" FOR RANDOM AS #1 LEN = RealRecLen
  RecCnt& = LOF(1) / RealRecLen
  FOR Cnt& = 1 TO RecCnt&
    LOCATE 5, 3
    PRINT "Processing:"; Cnt&; "of"; RecCnt&;
    GET #1, Cnt&, RealRec
    RealRec.LateList = "N"
    PUT #1, Cnt&, RealRec
  NEXT
  CLOSE

  LOCATE 7, 1
  PRINT "Personal Property. . ."
  OPEN "TaxPers.dat" FOR RANDOM AS #1 LEN = PersRecLen
  RecCnt& = LOF(1) / PersRecLen
  FOR Cnt& = 1 TO RecCnt&
    LOCATE 9, 3
    PRINT "Processing:"; Cnt&; "of"; RecCnt&;
    GET #1, Cnt&, PersRec
    PersRec.LateList = "N"
    PUT #1, Cnt&, PersRec
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT " Done."



  

