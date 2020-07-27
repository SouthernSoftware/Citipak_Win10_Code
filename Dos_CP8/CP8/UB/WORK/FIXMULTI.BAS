DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  'REDIM UBTransRec(1) AS UBTransRecType
  'UBTransRecLen = LEN(UBTransRec(1))
  'END
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;

    GET UBCFile, Cnt&, UBCustRec(1)
    DidEm = 0
    FOR ZZ = 1 TO 7
      Curr& = UBCustRec(1).LocMeters(ZZ).CurRead
      Prev& = UBCustRec(1).LocMeters(ZZ).PrevRead
      IF Curr& >= 0 OR Prev& >= 0 THEN
        UBCustRec(1).LocMeters(ZZ).MTRMulti = 10
        'UBCustRec(1).LocMeters(ZZ).CurRead = UBCustRec(1).LocMeters(ZZ).CurRead / 100
        'UBCustRec(1).LocMeters(ZZ).PrevRead = UBCustRec(1).LocMeters(ZZ).PrevRead / 100
        DidEm = -1
        PUT UBCFile, Cnt&, UBCustRec(1)
      END IF
    NEXT
    IF DidEm = -1 THEN
      PUT UBCFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE

'  LOCATE 9, 1
'  PRINT "Processing Transactions. . ."
'
'  UBTFile = FREEFILE
'  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBCFile LEN = UBTransRecLen
'  TNumOfRecs& = LOF(UBTFile) \ UBTransRecLen
'  FOR Cnt& = 1 TO TNumOfRecs&
'    LOCATE 12, 1
'    PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
'
'    GET UBTFile, Cnt&, UBTransRec(1)
'    DidEm = 0
'    FOR ZZ = 1 TO 7
'      Curr& = UBTransRec(1).CurRead(ZZ)
'      Prev& = UBTransRec(1).PrevRead(ZZ)
'      IF Curr& >= 100 AND Prev& >= 100 THEN
'        Curr& = Curr& / 100
'        'UBTransRec(1).CurRead(ZZ)
'        UBCustRec(1).LocMeters(ZZ).MTRMulti = 100
'        UBCustRec(1).LocMeters(ZZ).CurRead = UBCustRec(1).LocMeters(ZZ).CurRead / 100
'        UBCustRec(1).LocMeters(ZZ).PrevRead = UBCustRec(1).LocMeters(ZZ).PrevRead / 100
'        DidEm = -1
'        PUT UBCFile, Cnt&, UBCustRec(1)
'      END IF
'    NEXT
'    IF DidEm = -1 THEN
'      PUT UBCFile, Cnt&, UBCustRec(1)
'    END IF
'  NEXT
'  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."
  KillFile "FIXMULTI.EXE"


