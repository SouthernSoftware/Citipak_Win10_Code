DECLARE SUB ReadTransPins ()
DECLARE SUB SumEmpYTD ()

DEFINT A-Z
SUB RelinkTransHistory (YearReLink)
  
  IF YearReLink > 0 THEN
    RelinkYear$ = QPStrI(YearReLink)
  ELSE
    RelinkYear$ = RIGHT$(COMMAND$, 2)
  END IF
  
  IF NOT INSTR("949596979899", RelinkYear$) > 0 THEN
    BEEP
    PRINT
    PRINT
    PRINT "Invalid year spec for relink!"
    EXIT SUB
  END IF
  
  LowDate = Date2Num("01-01-19" + RelinkYear$)
  HiDate = Date2Num("12-31-19" + RelinkYear$)
  
  REDIM TPntr(0 TO 200)
  
  REDIM TPins(1 TO 1)
  
  DIM ToScrn AS STRING * 60
  
  Emp2RecLen = LEN(Emp2Rec)
  Emp3RecLen = LEN(Emp3Rec)
  
  TranRecLen = LEN(TranRec)
  
  ENumOfRec = FileSize(EmpData2Name) \ Emp2RecLen
  TNumOfRec = FileSize(TransHistFileName) \ TranRecLen
  
  '  REDIM IdxBuff(1 TO ENumOfRec)
  '  FGetAH EmpIdxLName, IdxBuff(1), 2, ENumOfRec
  
  NewRecCnt = 1
  
  '  COLOR 15, 1
  '  CLS
  ReadTransPins
  
  NAME TransHistFileName AS OldHistFileName
  
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2RecLen
  OPEN EmpData3Name FOR RANDOM AS #4 LEN = Emp3RecLen
  OPEN OldHistFileName FOR RANDOM AS #2 LEN = TranRecLen
  OPEN TransHistFileName FOR RANDOM AS #3 LEN = TranRecLen
  
  '  OPEN EmpData3Name FOR RANDOM AS #3 LEN = Emp3RecLen
  Row = CSRLIN
  FOR ECnt = 1 TO ENumOfRec
    GET #1, ECnt, Emp2Rec
    LOCATE Row, 1
    LSET ToScrn = "Relinking Transactions: " + FUsing(STR$((ECnt / ENumOfRec) * 100), "###") + "% Completed."
    PRINT ToScrn;
    GOSUB GetTransRecNums
    IF TPntr(0) THEN
      GOSUB RebuildTransHistory
    ELSE
      Emp2Rec.LastTransRec = 0
    END IF
    PUT #1, ECnt, Emp2Rec
    PUT #4, ECnt, Emp3Rec
    
    Emp3Rec = Emp3RecB
    
  NEXT
  CLOSE
  
  LOCATE Row, 1
  LSET ToScrn = "Transaction Relink Complete."
  PRINT ToScrn;
  KILL OldHistFileName
  
  EXIT SUB
  
  
GetTransRecNums:
  REDIM TPntr(0 TO 200)
  TotalTransRecs = 0
  FOR TCnt = 1 TO TNumOfRec
    IF TPins(TCnt) = Emp2Rec.EmpPin THEN
      TotalTransRecs = TotalTransRecs + 1
      TPntr(TotalTransRecs) = TCnt
    END IF
    TPntr(0) = TotalTransRecs
  NEXT
  RETURN
  
RebuildTransHistory:
  FirstEmpHRec = NewRecCnt
  '   PRINT TPntr(0): STOP
  
  
  FOR Cnt = 1 TO TPntr(0)
    GET #2, TPntr(Cnt), TranRec
    IF Cnt = 1 THEN
      TranRec.PrevTransRec = 0
    ELSE
      TranRec.PrevTransRec = NewRecCnt - 1
    END IF
    TranRec.RetGrossPay = TranRec.GrossPay
    PUT #3, NewRecCnt, TranRec
    NewRecCnt = NewRecCnt + 1
    Emp2Rec.LastTransRec = NewRecCnt - 1
    SELECT CASE TranRec.CheckDate
    CASE LowDate TO HiDate
      SumEmpYTD
    END SELECT
  NEXT
  RETURN
  
  
END SUB

