
  'REDIM SHARED Items(0) AS FLen
  
  'PRINT LEN(EmpRec1), LEN(EmpRec3), LEN(EMP2Rec(1))
  
  'ON TIMER(1) GOSUB ShowStack
  'TIMER ON
  
  'ShowStack:
  '  QPrintRC STR$(FRE(-2)), 25, 1, 31
  '  RETURN
  
  '  DIM Emp3 AS EmpData3Type
  '  FOpenS EmpData3Name, Handle3
  '  RecLen = LEN(Emp3)
  '  NumOfRec = FileSize(EmpData3Name) \ RecLen
  '  FOR Cnt = 1 TO NumOfRec
  '    FGetRTA Handle3, Emp3, CLNG(Cnt), RecLen
  '    Emp3.Data1RecNum = Cnt
  '    FPutRTA Handle3, Emp3, CLNG(Cnt), RecLen
  '  NEXT
  '  FClose Handle3
  'END
  
  '  DIM Emp2 AS EmpData2Type
  '  RecLen = LEN(Emp2)
  '
  '  OPEN EmpData2Name FOR RANDOM AS #1 LEN = RecLen
  '
  '  NumOfRec = FileSize(EmpData2Name) \ RecLen
  '  FOR Cnt = 1 TO NumOfRec
  '    GET #1, Cnt, Emp2
  '     'IF Emp2.EmpPin = 171 THEN PRINT Emp2.EmpLName; Emp2.EmpFName
  '     IF UCASE$(QPTrim$(Emp2.EmpLName)) = "WILLIS" THEN STOP
  ''    PRINT Emp2.EmpPin
  '    z! = TIMER
  '    DO: LOOP WHILE z! + .1 > TIMER
  '    'z$ = INPUT$(1)
  '    IF INSTR(UCASE$(Emp2.EmpPType), "HOUR") THEN LPRINT Emp2.EmpNo
  '
  '  NEXT
  '  CLOSE
  ''STOP
  'END
  '
'  TDate = Date2Num("03-29-1996")
'  DIM TR AS TransRecType
'  RecLen = LEN(TR)
  '
  'KILL "prdata\newtranh.dat"
  
'  NumOfRec = FileSize(TransHistFileName) \ RecLen
'  OPEN TransHistFileName FOR RANDOM AS #1 LEN = RecLen
  
'  OPEN "prdata\newtranh.dat" FOR RANDOM AS #2 LEN = RecLen
  '
  'RecCnt = 1
  
'  FOR Cnt = 1 TO NumOfRec
'    GET #1, Cnt, TR
'    '
'    IF TR.EmpPin = 47 THEN      'AND TR.CheckDate = TDate THEN
'
'      'IF TR.CheckNum = 20651 OR TR.CheckNum = 17 THEN
'      Removed = Removed + 1
'    ELSE
'      PUT #2, , TR
'    END IF
    'ELSE
    '  PUT #2, , TR
    'END IF
    '     PRINT TR.CheckNum, TR.TotRegWage, TR.GrossWage, TR.NETPAY
    'STOP
    '      IF TR.CheckNum > 0 THEN
    '        'PUT #2, RecCnt, TR
    '        RecCnt = RecCnt + 1
    '        Removed = Removed + 1
    '      END IF
    '    ELSE
    '      'PUT #2, RecCnt, TR
    '      RecCnt = RecCnt + 1
    '    END IF
'  NEXT
'  CLOSE
'  PRINT "Total removed:"; Removed
'  END
  '
  'PRINT TotalTra, Net#
  
  '  DIM TR AS TransRecType
  '  LastEmpPin = -100
  
  '  RecLen = LEN(TR)
  '  Date1 = Date2Num("06-07-1995")
  'Date2 = Date2Num("08-31-1995")
  
  '  OPEN TransHistFileName FOR RANDOM AS #1 LEN = RecLen
  '  OPEN "prdata\Temphist.dat" FOR RANDOM AS #2 LEN = RecLen
  
  '  NumOfRec = FileSize(TransHistFileName) \ RecLen
  '  FOR Cnt = 1 TO NumOfRec
  '    GET #1, Cnt, TR
  '    IF TR.RetireAmt = 0 THEN '    LastEmpPin <> TR.EmpPin THEN  'TR.CheckDate <= Date2 THEN
  '      TR.MatchRetAmt = 0
  '      DidCnt = DidCnt + 1
  '      PUT #1, Cnt, TR
  '    END IF
  
  '      PRINT TR.GrossPay
  '      STOP
  '      EXIT FOR
  '    END IF
  
  '     IF TR.DAMT(2) > 0 THEN
  '       TGrossPay# = Round(TR.GROSSPAY - TR.DAMT(2))
  '     ELSE
  '       TGrossPay# = TR.GROSSPAY
  '     END IF
  '    TR.FedGrossPay = TGrossPay#
  '    TR.StaGrossPay = TGrossPay#
  '    TR.SocGrossPay = TR.GROSSPAY
  '    TR.MedGrossPay = TR.GROSSPAY
  '    TR.RetGrossPay = TR.GROSSPAY
  '    PUT #1, Cnt, TR
  '  NEXT
  '  CLOSE
  '  CLS
  '  PRINT "Number of Trans:"; NumOfRec; " Changed:"; DidCnt
