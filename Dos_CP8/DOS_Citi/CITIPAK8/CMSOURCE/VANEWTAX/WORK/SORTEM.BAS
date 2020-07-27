DECLARE SUB ClearBack ()
DECLARE SUB DisplayDCScrn (ScrnName$)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB SortDCNameIndex

  'SHARED MChoice$

  ClearBack
  ShowProcessingScrn "Creating Customer Name Index"

  REDIM DCCustRec(1) AS DCCustRecType
  REDIM DCCustIdxRec(1) AS DCCustIDXRecType
  
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 14    'size of the key element - coded as follows:

  GoodRecords = 0
  
  'Help$ = "Sorting Customer Index"
  'PrintHelp Help$
'STOP
  DCCustRecLen = LEN(DCCustRec(1))

  DCFile = FREEFILE
  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  REDIM Array(1 TO NumOfDCRecs)  AS Struct
  FOR Cnt = 1 TO NumOfDCRecs
    GET DCFile, Cnt, DCCustRec(1)
    IF (DCCustRec(1).Deleted <> "Y") AND (LEN(QPTrim$(DCCustRec(1).CUSTNUMB)) > 0) THEN
    'IF (DCCustRec(1).Deleted <> "Y") THEN
      GoodRecords = GoodRecords + 1
      LSET Array(GoodRecords).who = DCCustRec(1).SORTNAME
      Array(GoodRecords).RecNum = Cnt
    END IF
    ShowPctCompL Cnt, NumOfDCRecs
  NEXT Cnt
  CLOSE DCFile

  SortT Array(1), GoodRecords, Dir, SSize, MOff, MSize

  KillFile "DCCUST.IDX"
  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  DCIdxFile = FREEFILE
  OPEN "DCCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS DCIdxFile LEN = DCCustIdxRecLen
  FOR Cnt = 1 TO GoodRecords
    DCCustIdxRec(1).IDXNAME = Array(Cnt).who
    DCCustIdxRec(1).IDXRECORD = Array(Cnt).RecNum
    PUT DCIdxFile, Cnt, DCCustIdxRec(1)
    ShowPctCompL Cnt, GoodRecords
  NEXT Cnt
  CLOSE

'  RelinkDCStuff

  ClearBack
  DisplayDCScrn "UPDATEOK"
  WaitForAction

END SUB

