DEFINT A-Z
FUNCTION GetRevCharge# (RateTbl AS UBRateTblRecType, TMeterConsp&, MeterMulti&) STATIC
  
  MinBillAmt# = RateTbl.MINAMT
  
  IF MinBillAmt# < -1000000 THEN
    MinBillAmt# = 0
    TAmt# = -1
    GOTO GotTAmt
  END IF
  
  LastTblCnt = 10
  FOR BCnt = 1 TO 10
    IF RateTbl.TblBreaks(BCnt).UNITAMT <= 0 THEN
      LastTblCnt = BCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  MeterConsump& = TMeterConsp&
  
  TAmt# = 0
  
  IF LastTblCnt >= 2 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(1).UNITS AND MeterConsump& < RateTbl.TblBreaks(2).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    END IF
  ELSE          'no other rate breaks
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 2
  IF LastTblCnt >= 3 THEN
    IF MeterConsump& > RateTbl.TblBreaks(2).UNITS AND MeterConsump& < RateTbl.TblBreaks(3).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 3
  IF LastTblCnt >= 4 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(3).UNITS AND MeterConsump& < RateTbl.TblBreaks(4).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'Break 4
  IF LastTblCnt >= 5 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(4).UNITS AND MeterConsump& < RateTbl.TblBreaks(5).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 5
  IF LastTblCnt >= 6 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(5).UNITS AND MeterConsump& < RateTbl.TblBreaks(6).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 6
  IF LastTblCnt >= 7 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(6).UNITS AND MeterConsump& < RateTbl.TblBreaks(7).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 7
  IF LastTblCnt >= 8 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(7).UNITS AND MeterConsump& < RateTbl.TblBreaks(8).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 8
  IF LastTblCnt >= 9 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(8).UNITS AND MeterConsump& < RateTbl.TblBreaks(9).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GOTO GotTAmt
  END IF
  
  'break 9
  IF LastTblCnt >= 10 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(9).UNITS AND MeterConsump& < RateTbl.TblBreaks(10).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    GOTO GotTAmt
  END IF
  
GotTAmt:
  GetRevCharge# = Round#(MinBillAmt# + TAmt#)
  
END FUNCTION

