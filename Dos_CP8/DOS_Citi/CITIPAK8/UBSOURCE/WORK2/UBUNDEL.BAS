DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'

  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Customers. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  UNDel = FREEFILE
  OPEN "UNDEL.LST" FOR INPUT AS UNDel
  
  TNumOfRecs& = LOF(UBFile) / UBCustRecLen
  
  DO
    LINE INPUT #UNDel, Acct$
    Account = VAL(Acct$)

    IF Account > 0 AND Account <= TNumOfRecs& THEN
      LOCATE 5, 1
      PRINT "Undeleting:"; Account;
      GET UBFile, Account, UBCustRec(1)
      UBCustRec(1).DelFlag = 0
      Fixed = Fixed + 1
      PUT UBFile, Account, UBCustRec(1)
    END IF
  LOOP UNTIL EOF(UNDel)
  CLOSE

  LOCATE 9, 1
  PRINT "Number of Accounts UNDELETED:"; Fixed
  'KILL "undel.lst"

