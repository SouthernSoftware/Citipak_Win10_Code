DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)


  '$INCLUDE: 'lawrence.BI'

  CONST False = 0, True = NOT False

  CLS

  CLS
  LOCATE 2, 1
  PRINT "Meter Coordinate Import Utility v2.0"
  PRINT "(c)2004-05 Southern Software Inc. All rights reserved."

  SensusIOFile$ = "c:\autoread\EXlatlng.DAT"

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  REDIM UBSenGetRdRec(1) AS UBGilSensusGetReadRecType
  UBSenGetRecLen = LEN(UBSenGetRdRec(1))

  UBSenIOFile = FREEFILE
  OPEN SensusIOFile$ FOR RANDOM SHARED AS UBSenIOFile LEN = UBSenGetRecLen
      
  NumSenGetRecs = LOF(UBSenIOFile) / UBSenGetRecLen
    
  IF NumSenGetRecs = 0 THEN
    LOCATE 7, 1
    PRINT "CAN NOT FIND 'c:\autoread\EXlatlng.DAT' CORRECT AND TRY AGAIN!"

    GOTO DoneNow:
  END IF
  
  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen

  RMCnt& = 1                ' Initialize File Counter to 1
  DO
    LOCATE 7, 1
    PRINT "Processing:"; RMCnt&; "of"; NumSenGetRecs;
    GET UBSenIOFile, RMCnt&, UBSenGetRdRec(1)
    WhatRMRec = VAL(QPTrim$(UBSenGetRdRec(1).Account))
    IF WhatRMRec > 0 THEN
      GOSUB EExtractRecord
    END IF
    RMCnt& = RMCnt& + 1
  LOOP UNTIL RMCnt& > NumSenGetRecs

DoneNow:
  LOCATE 12, 1
  PRINT "Import Complete."
  PRINT
  CLOSE

  END

EExtractRecord:

  GET UBFile, WhatRMRec, UBCustRec(1)

  DashPos = INSTR(UBSenGetRdRec(1).Account, "-")
  MtrCnt = VAL(MID$(UBSenGetRdRec(1).Account, DashPos + 1))
        
  IF MtrCnt = 0 THEN MtrCnt = 1
  ' Check Meter Updated Flag
  ' Update Meter W/Reading
  UBCustRec(1).LocMeters(MtrCnt).MtrLat = VAL(QPTrim$(UBSenGetRdRec(1).MtrLat))
  UBCustRec(1).LocMeters(MtrCnt).MtrLng = VAL(QPTrim$(UBSenGetRdRec(1).MtrLng))
                                    
  PUT UBFile, WhatRMRec, UBCustRec(1)

RETURN



