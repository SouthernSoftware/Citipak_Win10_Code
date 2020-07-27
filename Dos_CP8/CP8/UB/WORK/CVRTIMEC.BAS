DEFINT A-Z

  '$INCLUDE: 'ubInter.bi'
  '$INCLUDE: 'ubJett.bi'
  CLS

  REDIM UBInterRDRec(1)   AS UBIntermecHHRecType
  REDIM UBJettRDRec(1)    AS UBJettHHRecType

  UBInterRDRecLen = LEN(UBInterRDRec(1))
  UBJettRDRecLen = LEN(UBJettRDRec(1))

  HHFile = FREEFILE
  OPEN "CEMTREAD.DAT" FOR RANDOM SHARED AS HHFile LEN = UBInterRDRecLen
  NumIntermecRdRecs = LOF(HHFile) / UBInterRDRecLen

  JHFile = FREEFILE
  OPEN "CEMTREAD.DAT" FOR RANDOM SHARED AS JHFile LEN = UBJettRDRecLen

  FOR Cnt = NumIntermecRdRecs TO 1 STEP -1
    LOCATE 3, 1
    PRINT "Converting: "; Cnt;
    GET HHFile, Cnt, UBInterRDRec(1)

    UBJettRDRec(1).CEVariant = UBInterRDRec(1).CEVariant
    UBJettRDRec(1).CEStrLen = UBInterRDRec(1).CEStrLen
    UBJettRDRec(1).CustName = UBInterRDRec(1).CustName
    UBJettRDRec(1).ServAddress = UBInterRDRec(1).ServAddress
    UBJettRDRec(1).MeterID = UBInterRDRec(1).MeterID
    UBJettRDRec(1).LowRead = UBInterRDRec(1).LowRead
    UBJettRDRec(1).HighRead = UBInterRDRec(1).HighRead
    UBJettRDRec(1).Account = UBInterRDRec(1).Account
    UBJettRDRec(1).ReadFlag = UBInterRDRec(1).ReadFlag
    UBJettRDRec(1).MeterType = UBInterRDRec(1).MeterType
    UBJettRDRec(1).Book = UBInterRDRec(1).Book
    UBJettRDRec(1).CurRead = UBInterRDRec(1).CurRead
    UBJettRDRec(1).PastRead = UBInterRDRec(1).PastRead
    UBJettRDRec(1).ReadDate = UBInterRDRec(1).ReadDate
    LSET UBJettRDRec(1).ReadTime = ""
    UBJettRDRec(1).Note1 = UBInterRDRec(1).Note1
    UBJettRDRec(1).Note2 = UBInterRDRec(1).Note2
    UBJettRDRec(1).Note3 = UBInterRDRec(1).Note3
    UBJettRDRec(1).NoteStatus = UBInterRDRec(1).NoteStatus

    PUT JHFile, Cnt, UBJettRDRec(1)

  NEXT
  CLOSE

