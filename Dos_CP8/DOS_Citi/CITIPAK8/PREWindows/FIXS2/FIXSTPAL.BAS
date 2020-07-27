  DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Txt$)
  
  '$INCLUDE: 'newcust.BI'
  
  CLS
  PRINT
  PRINT
  PRINT "Processing Customers. . ."
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT
 ' LPRINT "Accounts with a meter type and NO Rate Code"
 ' LPRINT
 ' LPRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    DidOne = 0
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      FOR MCnt = 1 TO 2
        RateCode$ = QPTrim$(UBCustRec(1).Serv(MCnt).RateCode)
        RMTRTYPE$ = QPTrim$(UBCustRec(1).Serv(MCnt).RMTRTYPE)
        IF LEN(RateCode$) = 0 THEN
          IF LEN(RMTRTYPE$) > 0 THEN
            FoundMtr = FoundMtr + 1
           ' LPRINT Cnt&, ;
            DidOne = -1
            UBCustRec(1).Serv(MCnt).RMTRTYPE = ""
          END IF
        END IF
      NEXT
DoneThisOne:
    END IF
    IF DidOne THEN
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
  
  PRINT
  PRINT
  PRINT
  PRINT "Meter no Rate:"; FoundMtr
 ' LPRINT
 ' LPRINT
 ' LPRINT "End of list"
  

