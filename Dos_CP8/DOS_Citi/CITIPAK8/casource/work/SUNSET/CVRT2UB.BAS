DEFINT A-Z
DECLARE FUNCTION MakeSearchName$ (CUSTNAME$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'CACust.BI'
  '$INCLUDE: 'newCust.BI'
CLS

  MaxIntg% = -32767
  MaxLong& = -2147483647

  DIM CACust  AS CACUSTType
  CALen = LEN(CACust)  'Length of Cust Record Structure
  DIM UBCust      AS NewUBCustRecType
  DIM BlankCust   AS NewUBCustRecType
  UBLen = LEN(UBCust)
  today = Date2Num(DATE$)

  lowAsc = ASC("0")
  HigAsc = ASC("9")
  dshAsc = ASC("-")

  CALL KillFile("ubcust.dat")

  CAFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CAFile LEN = CALen
  NumOfRecs& = LOF(CAFile) \ CALen
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBLen
  FOR Cnt& = 1 TO NumOfRecs&
    GET #CAFile, Cnt&, CACust
    LSET UBCust = BlankCust
    STATE$ = ""
    City$ = ""
    Zip$ = ""
    SEQ& = Cnt& * 10
    Sequence$ = "000000" + QPTrim$(STR$(SEQ&))

    UBCust.BOOK = "01"
    UBCust.SEQNUMB = RIGHT$(Sequence$, 6)

    UBCust.Status = "A"
    UBCust.OPENDATE = today
    Cust$ = QPTrim$(CACust.CUSTNAME)
    IF (LEN(Cust$) = 0) OR (CACust.Deleted <> 0) THEN
      UBCust.DelFlag = -1
      GOTO NoNameSkip
    END IF
    UBCust.CUSTNAME = Cust$

    'Cust$ = QPTrim$(UBCust.CustName)
    UBCust.SEARCH = MakeSearchName$(QPTrim$(UBCust.CUSTNAME))

    UBCust.ADDR1 = QPTrim$(CACust.ADD1)
    ADDR2$ = QPTrim$(CACust.ADD2)
    Add2Len = LEN(ADDR2$)

    FOR zz = Add2Len TO 1 STEP -1
      ThisChar = ASC(MID$(ADDR2$, zz, 1))
      IF (ThisChar < lowAsc OR ThisChar > HigAsc) AND NOT (ThisChar = dshAsc) THEN
        Zip$ = MID$(ADDR2$, zz + 1)
        Dash = INSTR(Zip$, "-")
        IF Dash > 0 THEN
          Zip$ = LEFT$(Zip$, Dash - 1) + MID$(Zip$, Dash + 1)
        END IF
        ADDR2$ = QPTrim$(MID$(ADDR2$, 1, zz))
        Add2Len = LEN(ADDR2$)
        STATE$ = RIGHT$(ADDR2$, 2)
        ADDR2$ = MID$(ADDR2$, 1, Add2Len - 2)
        EXIT FOR
      END IF
    NEXT
    Comma = INSTR(ADDR2$, ",")
    IF Comma > 0 THEN
      City$ = QPTrim$(LEFT$(ADDR2$, Comma - 1))
      'Addr2$=
    ELSE
      City$ = QPTrim$(ADDR2$)
    END IF
    'UBCust.Addr2 = Addr2$
    UBCust.ZIPCODE = Zip$
    UBCust.STATE = STATE$
    UBCust.City = City$
    UBCust.SERVADDR = QPTrim$(CACust.PropDesc)
    UBCust.Addr911 = QPTrim$(CACust.LEGALDES)

    UBCust.CASHONLY = "N"
    UBCust.LATEFEE = "Y"
    UBCust.CUTOFFYN = "Y"
    UBCust.TAXEXPT = "N"
    UBCust.SRCIT = "N"

    UBCust.ProRatePCT = 100
    UBCust.HHMSG1 = UBCust.Addr911
    UBCust.HHMSG2 = UBCust.SERVADDR
    UBCust.HHMSG3 = QPTrim$(CACust.Parcel)
    UBCust.FlatRates(1).FRDESC = "UTILITY"
    UBCust.FlatRates(1).FRAMT = CACust.WATERAMT
    UBCust.FlatRates(1).FRFREQ = "R"
    UBCust.FlatRates(1).REVSRC = 5
    UBCust.FlatRates(1).NumMin = 1

    FOR zz = 1 TO 7
      UBCust.LocMeters(zz).MTRNUM = ""
      UBCust.LocMeters(zz).MTRMulti = MaxIntg
      UBCust.LocMeters(zz).MTRType = ""
      UBCust.LocMeters(zz).MTRUnit = ""
      UBCust.LocMeters(zz).NumUser = MaxIntg
      UBCust.LocMeters(zz).InsDate = MaxIntg
      UBCust.LocMeters(zz).CurRead = MaxLong&
      UBCust.LocMeters(zz).PrevRead = MaxLong&
      UBCust.LocMeters(zz).CurDate = MaxIntg
      UBCust.LocMeters(zz).PastDate = MaxIntg
      UBCust.LocMeters(zz).ReadFlag = ""
      UBCust.LocMeters(zz).AvgUse = 0
      UBCust.LocMeters(zz).UseCnt = 0
    NEXT
    UBCust.BillCycl = 1
    'PRINT UBCust.City, UBCust.State, UBCust.ZipCode
    'FOR zz& = 1 TO 30000
    '  a$ = INKEY$
    '  IF LEN(a$) > 0 THEN EXIT FOR
    'NEXT

  '  UBCust.ADDR2
    IF LEN(a$) > 0 THEN EXIT FOR
NoNameSkip:
    PUT #UBFile, , UBCust
  NEXT

  CLOSE

FUNCTION MakeSearchName$ (CUSTNAME$)

  CLen = LEN(CUSTNAME$)

  GOSUB RemoveSpaces

  Comma = INSTR(CUSTNAME$, ",")
  IF Comma > 0 THEN
    SmallName$ = LEFT$(CUSTNAME$, Comma - 1) + MID$(CUSTNAME$, Comma + 1)
  ELSE
    SmallName$ = CUSTNAME$
  END IF
  MakeSearchName$ = SmallName$

  EXIT FUNCTION

RemoveSpaces:
    Char$ = CUSTNAME$
    SPos = INSTR(Char$, " ")
    DO WHILE SPos > 0
      Char$ = LEFT$(Char$, SPos - 1) + MID$(Char$, SPos + 1)
      SPos = INSTR(Char$, " ")
    LOOP
    CUSTNAME$ = QPTrim$(Char$)

RETURN

END FUNCTION

