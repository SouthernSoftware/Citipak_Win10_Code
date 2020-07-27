DECLARE SUB ShowCustHistory (CustRec&)
  DEFINT A-Z
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FClose (Handle%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.bi'

  CONST False = 0, True = NOT False

SUB DeleteCust
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  DO
    BlockClear
    
    LookUp RecNo&, "Customer to Delete", 4, True, False
    
    IF RecNo& > 0 THEN
      GOSUB CheckCust
      IF CustOK THEN
        EXIT DO
      END IF
    ELSE
      EXIT DO
    END IF
  LOOP
  
  IF RecNo& > 0 THEN
    DisplayUBScrn "DELACUST"
    DO
      GOSUB DDisplayCustInfo
      Ok = MsgBox%("UB", "DODELBTN")
      SELECT CASE Ok
      CASE 1
        OKFlag = False
        ExitFlag = True
      CASE 2
        ExitFlag = True
        Ok = MsgBox%("UB", "OK2DEL")
        SELECT CASE Ok
        CASE 1
          OKFlag = False
        CASE 2
          OKFlag = True
        END SELECT
      CASE 3
        ShowCustHistory RecNo&
      END SELECT
    LOOP UNTIL ExitFlag
  END IF
  
  IF NOT OKFlag THEN
    GOTO ExitDeleteCust
  END IF
  
  BlockClear
  DisplayUBScrn "UPDATDSK"
  'WaitForAction
  GOSUB DeleteCustomer
  
  '******************************
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitDeleteCust:
  
  EXIT SUB
  
CheckCust:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  IF TotalBalance# <> 0 THEN
    Ok = MsgBox%("UB", "NODELBAL")
    CustOK = False
  ELSEIF UBCustRec(1).DepositAmt <> 0 THEN
    Ok = MsgBox%("UB", "NODELDEP")
    CustOK = False
  ELSEIF UBCustRec(1).Status <> "I" THEN
    Ok = MsgBox%("UB", "NODELNIN")
    CustOK = False
  ELSE
    CustOK = True
  END IF
  
  RETURN
  
DDisplayCustInfo:
  Zip$ = UBCustRec(1).ZIPCODE
  Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  QPrintRC UBCustRec(1).Book, 5, 26, -1
  QPrintRC UBCustRec(1).SEQNUMB, 5, 29, -1
  QPrintRC UBCustRec(1).Status, 5, 47, -1
  QPrintRC Num2Date(UBCustRec(1).OPENDATE), 5, 63, -1
  QPrintRC UBCustRec(1).SEARCH, 7, 26, -1
  QPrintRC STR$(RecNo&), 7, 62, -1
  QPrintRC UBCustRec(1).CUSTNAME, 8, 26, -1
  QPrintRC UBCustRec(1).ADDR1, 9, 26, -1
  QPrintRC UBCustRec(1).ADDR2, 10, 26, -1
  QPrintRC UBCustRec(1).SERVADDR, 11, 26, -1
  QPrintRC UBCustRec(1).CITY, 12, 26, -1
  QPrintRC UBCustRec(1).STATE, 12, 54, -1
  QPrintRC Zip$, 12, 63, -1
  RETURN
  
DeleteCustomer:
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  UBCustRec(1).Book = ""
  UBCustRec(1).SEQNUMB = ""
  UBCustRec(1).DelFlag = True
  PUT UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  RETURN
  
  'Update Book Index ***************************
  '  NumOfRecs = FileSize(BookIndexFile) \ 4
  '
  '  BlockSize = (NumOfRecs * 4)
  '  NewBlockSize = BlockSize + 4
  '
  '  REDIM IndexArray(1 TO NumOfRecs + 1) AS UBCustIndexRecType
  '  FOpenS BookIndexFile, IHandle
  '  FGetRTA IHandle, IndexArray(2), 1, BlockSize
  '
  '  FOR FCnt = 2 TO NumOfRecs + 1
  '    IF IndexArray(FCnt).RecNum = RecNo& THEN
  '      IndexArray(FCnt).RecNum = NextCustRecs&
  '      EXIT FOR
  '    END IF
  '  NEXT
  '  IndexArray(1).RecNum = RecNo&
  '
  '  FPutRTA IHandle, IndexArray(1), 1, NewBlockSize
  '  FClose IHandle
  
  'Update Name Index***************************
  '  NumOfRecs = FileSize(NameIndexFile) \ 4
  '  BlockSize = (NumOfRecs * 4)
  '  NewBlockSize = BlockSize + 4
  '  REDIM IndexArray(1 TO NumOfRecs + 1) AS UBCustIndexRecType
  '  FOpenS BookIndexFile, IHandle
  '  FGetRTA IHandle, IndexArray(2), 1, BlockSize
  '  IndexArray(1).RecNum = NextCustRecs&
  '  FPutRTA IHandle, IndexArray(1), 1, NewBlockSize
  '  FClose IHandle
  '******************************
  '******************************
  
END SUB

SUB ReUseCust

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  DO
    BlockClear

    LookUp RecNo&, "Customer to Reuse", 4, True, False

    IF RecNo& > 0 THEN
      GOSUB RCheckCust
      IF CustOK THEN
        EXIT DO
      END IF
    ELSE
      EXIT DO
    END IF
  LOOP

  IF RecNo& > 0 THEN
    DisplayUBScrn "DELACUST"
    DO
      GOSUB RDisplayCustInfo
      Ok = MsgBox%("UB", "DODELBTN")
      SELECT CASE Ok
      CASE 1
        OKFlag = False
        ExitFlag = True
      CASE 2
        ExitFlag = True
        Ok = MsgBox%("UB", "OK2DEL")
        SELECT CASE Ok
        CASE 1
          OKFlag = False
        CASE 2
          OKFlag = True
        END SELECT
      CASE 3
        ShowCustHistory RecNo&
      END SELECT
    LOOP UNTIL ExitFlag
  END IF

  IF NOT OKFlag THEN
    GOTO ExitReuseCust
  END IF

  BlockClear
  DisplayUBScrn "UPDATDSK"
  GOSUB ReuseCustomer
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitReuseCust:

  EXIT SUB

RCheckCust:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  IF TotalBalance# <> 0 THEN
    Ok = MsgBox%("UB", "NODELBAL")
    CustOK = False
  ELSEIF UBCustRec(1).DepositAmt <> 0 THEN
    Ok = MsgBox%("UB", "NODELDEP")
    CustOK = False
  ELSEIF UBCustRec(1).Status <> "I" THEN
    Ok = MsgBox%("UB", "NODELNIN")
    CustOK = False
  ELSE
    CustOK = True
  END IF

  RETURN

RDisplayCustInfo:
  Zip$ = UBCustRec(1).ZIPCODE
  Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  QPrintRC UBCustRec(1).Book, 5, 26, -1
  QPrintRC UBCustRec(1).SEQNUMB, 5, 29, -1
  QPrintRC UBCustRec(1).Status, 5, 47, -1
  QPrintRC Num2Date(UBCustRec(1).OPENDATE), 5, 63, -1
  QPrintRC UBCustRec(1).SEARCH, 7, 26, -1
  QPrintRC STR$(RecNo&), 7, 62, -1
  QPrintRC UBCustRec(1).CUSTNAME, 8, 26, -1
  QPrintRC UBCustRec(1).ADDR1, 9, 26, -1
  QPrintRC UBCustRec(1).ADDR2, 10, 26, -1
  QPrintRC UBCustRec(1).SERVADDR, 11, 26, -1
  QPrintRC UBCustRec(1).CITY, 12, 26, -1
  QPrintRC UBCustRec(1).STATE, 12, 54, -1
  QPrintRC Zip$, 12, 63, -1
  RETURN

ReuseCustomer:

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  UBCustRec(1).Book = ""
  UBCustRec(1).SEQNUMB = ""
  UBCustRec(1).DelFlag = True
  PUT UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  RETURN

  'Update Book Index ***************************
  '  NumOfRecs = FileSize(BookIndexFile) \ 4
  '
  '  BlockSize = (NumOfRecs * 4)
  '  NewBlockSize = BlockSize + 4
  '
  '  REDIM IndexArray(1 TO NumOfRecs + 1) AS UBCustIndexRecType
  '  FOpenS BookIndexFile, IHandle
  '  FGetRTA IHandle, IndexArray(2), 1, BlockSize
  '
  '  FOR FCnt = 2 TO NumOfRecs + 1
  '    IF IndexArray(FCnt).RecNum = RecNo& THEN
  '      IndexArray(FCnt).RecNum = NextCustRecs&
  '      EXIT FOR
  '    END IF
  '  NEXT
  '  IndexArray(1).RecNum = RecNo&
  '
  '  FPutRTA IHandle, IndexArray(1), 1, NewBlockSize
  '  FClose IHandle

  'Update Name Index***************************
  '  NumOfRecs = FileSize(NameIndexFile) \ 4
  '  BlockSize = (NumOfRecs * 4)
  '  NewBlockSize = BlockSize + 4
  '  REDIM IndexArray(1 TO NumOfRecs + 1) AS UBCustIndexRecType
  '  FOpenS BookIndexFile, IHandle
  '  FGetRTA IHandle, IndexArray(2), 1, BlockSize
  '  IndexArray(1).RecNum = NextCustRecs&
  '  FPutRTA IHandle, IndexArray(1), 1, NewBlockSize
  '  FClose IHandle
  '******************************
  '******************************
  
END SUB

SUB SetCustFinal
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  BlockClear
  LookUp RecNo&, "Customer to Final", 2, True, True
  
  IF RecNo& > 0 THEN
    GOSUB GetCustInfo
    DisplayUBScrn "FINACUST"
    GOSUB DisplayCustInfo
    Ok = MsgBox%("UB", "DOFINBTN")
    
    SELECT CASE Ok
    CASE 1
      OKFlag = False
    CASE 2
      OKFlag = True
    END SELECT
  END IF
  
  IF NOT OKFlag THEN
    GOTO ExitSetFinal
  END IF
  
  '******************************
  BlockClear
  DisplayUBScrn "UPDATDSK"
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  GET UBCust, RecNo&, UBCustRec(1)
  LSET UBCustRec(2) = UBCustRec(1)
  UBCustRec(1).Status = "F"
  UBCustRec(1).Book = ""
  UBCustRec(1).SEQNUMB = ""
  PUT UBCust, RecNo&, UBCustRec(1)
  
  UBCustRec(2).Status = "I"
  UBCustRec(2).OPENDATE = -32767
  UBCustRec(2).SEARCH = ""
  UBCustRec(2).CUSTNAME = "VACANT"
  UBCustRec(2).ADDR1 = ""
  UBCustRec(2).ADDR2 = ""
  UBCustRec(2).HPHONE = ""
  UBCustRec(2).WPHONE = ""
  UBCustRec(2).SOSEC = ""
  UBCustRec(2).DRVLIC = ""

  UBCustRec(2).USEDRAFT = ""
  UBCustRec(2).BANKNAME = ""
  UBCustRec(2).BANKLOC = ""
  UBCustRec(2).TRANSIT = ""
  UBCustRec(2).BANKACCT = ""

  UBCustRec(2).CustPIN = 0
  UBCustRec(2).LastTrans = 0
  UBCustRec(2).CurrBalance = 0
  UBCustRec(2).PrevBalance = 0
  FOR Cnt = 1 TO 15
    UBCustRec(2).CurrRevAmts(Cnt) = 0
    UBCustRec(2).PrevRevAmts(Cnt) = 0
  NEXT
  UBCustRec(2).DepositAmt = 0
  UBCustRec(2).DelFlag = 0
  UBCustRec(2).PreNoteFlag = 0
  UBCustRec(2).WOLastTrans = 0
  UBCustRec(2).EstFlag = ""
  UBCustRec(2).MessageRec = 0
  
  NextCustRecs& = (LOF(UBCust) \ UBCustRecLen) + 1
  
  PUT UBCust, NextCustRecs&, UBCustRec(2)
  
  CLOSE UBCust

'***************************************************
' 08-18-97 reworked to completely reindex

  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure

  ShowProcessingScrn "Reading Customer Names"
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  REDIM IdxBuff(1 TO NumOfRecs&) AS UBCustReIndexRecType

  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
    IdxBuff(Cnt).First = QPTrim$(UBCustRec(1).CUSTNAME)
    IdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT

  CLOSE UBFile

  QPrintRC " Sorting Customer Names", 11, 29, -1

  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 10, 2
  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 10

  GOSUB ClearBlock
  QPrintRC "Writing Customer Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1

  KillFile "UBCUSTNM.IDX"

  UBFile = FREEFILE
  OPEN "UBCUSTNM.IDX" FOR RANDOM SHARED AS UBFile LEN = 4

  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, IdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT

  CLOSE UBFile
  ERASE IdxBuff
  GOSUB ClearBlock

  QPrintRC "Reading Location Information", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  REDIM LIdxBuff(1 TO NumOfRecs&) AS UBLocaReIndexRecType

  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    LIdxBuff(Cnt).Book = UBCustRec(1).Book
    LIdxBuff(Cnt).SEQNUMB = UBCustRec(1).SEQNUMB
    LIdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT

  CLOSE UBFile

  QPrintRC " Sorting Locations Names", 11, 29, -1

  SortT LIdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 8

  GOSUB ClearBlock
  QPrintRC "Writing Location Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1

  KillFile "UBCUSTBK.IDX"

  UBFile = FREEFILE
  OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS UBFile LEN = 4

  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, LIdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile
  REDIM BookSeq(1) AS BookSeqRecType

  KillFile "UBOOKSEQ.DAT"
  BookHand = FREEFILE
  OPEN "UBOOKSEQ.DAT" FOR RANDOM SHARED AS BookHand LEN = 4
  FOR Cnt = 1 TO NumOfRecs&
    BookSeq(1).BookSeq = QPValL(LIdxBuff(Cnt).Book + LIdxBuff(Cnt).SEQNUMB)
    PUT BookHand, Cnt, BookSeq(1)
  NEXT
  CLOSE BookHand
  ERASE LIdxBuff, BookSeq

  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitSetFinal:
  
  EXIT SUB

ClearBlock:
  HideCursor
  Blank$ = SPACE$(40)
  FOR Cnt = 8 TO 15
    QPrintRC Blank$, Cnt, 21, -1
  NEXT
  ShowCursor
RETURN

  
  
GetCustInfo:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  RETURN
  
DisplayCustInfo:
  Zip$ = UBCustRec(1).ZIPCODE
  Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  QPrintRC UBCustRec(1).Book, 5, 26, -1
  QPrintRC UBCustRec(1).SEQNUMB, 5, 29, -1
  QPrintRC UBCustRec(1).Status, 5, 47, -1
  QPrintRC Num2Date(UBCustRec(1).OPENDATE), 5, 63, -1
  QPrintRC UBCustRec(1).SEARCH, 7, 26, -1
  QPrintRC UBCustRec(1).CUSTNAME, 8, 26, -1
  QPrintRC UBCustRec(1).ADDR1, 9, 26, -1
  QPrintRC UBCustRec(1).ADDR2, 10, 26, -1
  QPrintRC UBCustRec(1).SERVADDR, 11, 26, -1
  QPrintRC UBCustRec(1).CITY, 12, 26, -1
  QPrintRC UBCustRec(1).STATE, 12, 54, -1
  QPrintRC Zip$, 12, 63, -1
  RETURN
  
  
  
  
END SUB

