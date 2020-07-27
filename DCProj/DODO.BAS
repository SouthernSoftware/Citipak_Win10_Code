DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB ReIndexSystem (PromptFlag%)
  
  UBLog " IN: Reindex Utility Files"
  
  BlockClear
  IF PromptFlag% THEN
    Ok = MsgBox%("UB", "MUSTEXIT")
    SELECT CASE Ok
    CASE 2
      GOTO ExitReindex
    END SELECT
  END IF
  
  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Reading Customer Names"
  UBLog "BEGIN: Customer Name Reindex"
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs&) AS nUBCustReIndexRecType
  
  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
    IF UBCustRec(1).DelFlag THEN
      IdxBuff(Cnt).DelFlag = "Y"
    ELSE
      IdxBuff(Cnt).DelFlag = ""
    END IF
    IdxBuff(Cnt).Status = UBCustRec(1).Status
    IdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  
  CLOSE UBFile
  
  QPrintRC " Sorting Customer Names", 11, 29, -1
  
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
  
  GOSUB ClearBlock
  QPrintRC "Writing Customer Search Data", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  KillFile "UBCUSTSN.DAT"
  UBFile = FREEFILE
  OPEN "UBCUSTSN.DAT" FOR RANDOM SHARED AS UBFile LEN = LEN(IdxBuff(1))
  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, IdxBuff(Cnt)
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile
  
  ERASE IdxBuff
  UBLog "FINISH: Customer Name Reindex"
  GOSUB ClearBlock
  
  QPrintRC "Reading Location Information", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Book\Sequence Reindex"
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  REDIM LIdxBuff(1 TO NumOfRecs&) AS UBLocaReIndexRecType
  
  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    LIdxBuff(Cnt).Book = UBCustRec(1).Book
    LIdxBuff(Cnt).SeqNumb = UBCustRec(1).SeqNumb
    LIdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  
  CLOSE UBFile
  
  QPrintRC " Sorting Locations Names", 11, 29, -1
  
  SortT LIdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 8
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  
  GOSUB ClearBlock
  QPrintRC "Writing Location Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  'here
  KillFile "UBCUSTBK.IDX"
  
  UBFile = FREEFILE
  OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS UBFile LEN = 4
  
  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, LIdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile
  
  UBLog "FINISH: Book\Sequence Reindex"
  REDIM BookSeq(1) AS BookSeqRecType
  
  KillFile "UBOOKSEQ.DAT"
  UBLog "BEGIN: Rebuild Book\Sequence List"
  BookHand = FREEFILE
  OPEN "UBOOKSEQ.DAT" FOR RANDOM SHARED AS BookHand LEN = 4
  FOR Cnt = 1 TO NumOfRecs&
    BookSeq(1).BookSeq = QPValL(LIdxBuff(Cnt).Book + LIdxBuff(Cnt).SeqNumb)
    PUT BookHand, Cnt, BookSeq(1)
  NEXT
  CLOSE BookHand
  UBLog "FINISH: Rebuild Book\Sequence List"

  ERASE LIdxBuff, BookSeq, IdxBuff
  ERASE UBCustRec, UBTransRec
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitReindex:
  UBLog "OUT: Reindex Utility Files" + CRLF$
  EXIT SUB
  
ClearBlock:
  HideCursor
  Blank$ = SPACE$(40)
  FOR Cnt = 8 TO 15
    QPrintRC Blank$, Cnt, 21, -1
  NEXT
  ShowCursor
RETURN
  
END SUB

