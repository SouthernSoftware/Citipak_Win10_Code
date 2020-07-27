DEFINT A-Z

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))
  CLS
  LOCATE 2, 1, 0
  PRINT "Correcting Accounts. . ."
  
  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
    
    GET UBFile, 2017, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2017, UBCustRec(1)

    GET UBFile, 2158, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2158, UBCustRec(1)

    GET UBFile, 2763, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2763, UBCustRec(1)

    GET UBFile, 2902, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2902, UBCustRec(1)

    GET UBFile, 2922, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2922, UBCustRec(1)

    GET UBFile, 2957, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2957, UBCustRec(1)

    GET UBFile, 2966, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2966, UBCustRec(1)

    GET UBFile, 2972, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2972, UBCustRec(1)

    GET UBFile, 2978, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 2978, UBCustRec(1)

    GET UBFile, 3008, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 3008, UBCustRec(1)

    GET UBFile, 3020, UBCustRec(1)
    GOSUB FixRDate
    PUT UBFile, 3020, UBCustRec(1)

CLOSE
LOCATE 7,1
PRINT "Corrected: 11 Accounts."

KILL "fixaccts.exe"

END


FixRDate:
    FOR MCnt = 1 TO 7
      UBCustRec(1).LocMeters(MCnt).CurDate = -32768
      UBCustRec(1).LocMeters(MCnt).PastDate = -32768
    NEXT
RETURN
  

