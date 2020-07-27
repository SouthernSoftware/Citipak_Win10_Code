DEFINT A-Z

DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)


  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'UBVBcust.bi'
  '$INCLUDE: 'ubschlum.bi'
  
  CONST False = 0, True = NOT False
  
  ImportFileName$ = "PC2HOST.EXP"  'File IN from them.

  REDIM UBSchlumHHRec(1) AS SchlumHHType
  UBSchlumHHRecLen = LEN(UBSchlumHHRec(1))

  UBSchlFile = FREEFILE
  OPEN ImportFileName$ FOR RANDOM SHARED AS UBSchlFile LEN = UBSchlumHHRecLen

  NumTRGetRecs = LOF(UBSchlFile) / UBSchlumHHRecLen

  IF NumTRGetRecs = 0 THEN
    CLOSE
    BEEP
    STOP
  END IF

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  CLS
  PRINT "Importing Readings from PC2Host file. . ."

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust& = LOF(UBFile) / UBCustRecLen

  FOR Cnt& = 1 TO NumTRGetRecs
    Prec& = 0
    LOCATE 5, 1

    PRINT "Processing:"; Cnt&; " of"; NumTRGetRecs;
    GET UBSchlFile, Cnt&, UBSchlumHHRec(1)

    BookSeq$ = LEFT$(UBSchlumHHRec(1).Notes8, 8) '+ MID$(UBSchlumHHRec(1).Notes8, 3, 6)

    FOR CustCnt& = 1 TO NumOfCust&
      GET UBFile, CustCnt&, UBCustRec(1)
      CustSeq$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb

      IF CustSeq$ = BookSeq$ THEN
        Prec& = CustCnt&
        IF Prec& > 0 THEN
          GET UBFile, Prec&, UBCustRec(1)
          GOSUB SchlumExtractRecord
        END IF
        EXIT FOR
      END IF
    NEXT
  NEXT
  
  CLOSE

  LOCATE 10, 1
  PRINT "Done.   U:"; WereRead; " N:"; NewRead

  END

SchlumExtractRecord:
  BadDate = 0

  MeterSlot = VAL(QPTrim$(MID$(UBSchlumHHRec(1).Notes8, 9, 2)))

  CurReading# = VAL(QPTrim$(UBSchlumHHRec(1).MtrRead))
  IF CurReading# < 0 THEN CurReading# = 0

  'UCode1$ = QPTrim$(UBCustRec(1).UserCode1)
  UCode2$ = QPTrim$(UBCustRec(1).USERCODE2)

  IF LEN(UCode2$) = 0 THEN
    Multi# = 1
  ELSE
    Multi# = 10 'UBCustRec(1).LocMeters(MeterSlot).MTRMulti
  END IF

  UCode1 = VAL(QPTrim$(UBCustRec(1).USERCODE1))
  SELECT CASE UCode1
  CASE 1
    Multi# = 10
  CASE 2
    Multi# = 100
  CASE 3
    Multi# = 1000
  END SELECT

  'IF Multi# = 0 THEN Multi# = 1
  'CurReading# = CurReading# * Multi#

  RYear$ = QPTrim$(RIGHT$(UBSchlumHHRec(1).ReadDate, 2))

  IF LEN(RYear$) < 2 THEN
    BadDate = True
  END IF

  WhatYear = VAL(RYear$)
  IF WhatYear < 65 THEN
    ReadYear$ = "-20" + RYear$
  ELSE
    ReadYear$ = "-19" + RYear$
  END IF
  DateRead$ = LEFT$(UBSchlumHHRec(1).ReadDate, 2) + "-" + MID$(UBSchlumHHRec(1).ReadDate, 3, 2) + ReadYear$
  ReadDate = Date2Num(DateRead$)

  IF CurReading# >= 9999999999# THEN
    CurReading# = 999999999#
  END IF

'here dale
'IF UBCustRec(1).LocMeters(MeterSlot).CurRead <> CurReading# THEN
'  CLS
'  PRINT "cust:"; UBCustRec(1).LocMeters(MeterSlot).CurRead
'  PRINT "shlm:"; CurReading#
'  PRINT "Cust Account:"; PRec&
'  PRINT "Shlm ReadRec:"; Cnt&
'  PRINT "   MeterSlot:"; MeterSlot
'  CLOSE
'  END
'END IF


  IF UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y" THEN
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
    WereRead = WereRead + 1
  ELSE
    UBCustRec(1).LocMeters(MeterSlot).PrevRead = UBCustRec(1).LocMeters(MeterSlot).CurRead
    UBCustRec(1).LocMeters(MeterSlot).PastDate = UBCustRec(1).LocMeters(MeterSlot).CurDate
    UBCustRec(1).LocMeters(MeterSlot).ReadFlag = "Y"
    UBCustRec(1).LocMeters(MeterSlot).CurDate = ReadDate
    UBCustRec(1).LocMeters(MeterSlot).CurRead = CurReading#
    NewRead = NewRead + 1
  END IF

  c1$ = QPTrim$(UBCustRec(1).HHMSG1)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes1)
  IF LEN(s1$) > 0 THEN
    IF s1$ <> c1$ THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG1 = s1$
    END IF
  END IF
  c1$ = QPTrim$(UBCustRec(1).HHMSG2)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes2)
  IF LEN(s1$) > 0 THEN
    IF s1$ <> c1$ THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG2 = s1$
    END IF
  END IF
  c1$ = QPTrim$(UBCustRec(1).HHMSG3)
  s1$ = QPTrim$(UBSchlumHHRec(1).Notes3)
  IF LEN(s1$) > 0 THEN
    IF s1$ <> c1$ THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMSG3 = s1$
    END IF
  END IF

  BookSeq$ = LEFT$(UBSchlumHHRec(1).Notes8, 2) + "-" + MID$(UBSchlumHHRec(1).Notes8, 3, 6)

  PUT UBFile, Prec&, UBCustRec(1)

  RETURN

