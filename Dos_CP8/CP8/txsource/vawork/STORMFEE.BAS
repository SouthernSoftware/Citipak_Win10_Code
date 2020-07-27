DEFINT A-Z

DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)


  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'


  CONST False = 0, True = NOT False


  CrLf$ = CHR$(13) + CHR$(10)

  DIM RRec AS PropertyRecType
  DIM RLen AS INTEGER
  DIM RCnt AS LONG

  RLen = LEN(RRec)
  CLS
  DIM TCnt AS LONG
  DIM PCnt AS INTEGER
  DIM v$
  v$ = "|"
  PCnt = 0

  OPEN "TAXPROP.DAT" FOR RANDOM AS #1 LEN = RLen
  OPEN "Discount.TXT" FOR OUTPUT AS #2

  RCnt = LOF(1) / RLen
  FOR TCnt = 1 TO RCnt
   GET #1, TCnt, RRec
   IF RRec.Deleted = 0 THEN
     IF RRec.EXMPOTHR > 0 THEN
       PCnt = PCnt + 1
       PRINT #2, QPTrim(STR$(RRec.EXMPOTHR)); v$; QPTrim(RRec.RealPin); v$; QPTrim(STR$(RRec.CustPin)); v$
       'IF PCnt MOD 10 = 0 THEN
       '  SLEEP
       'END IF
     END IF
   END IF
  NEXT
  CLOSE
  PRINT "Total:"; PCnt
  'END

'StormFee

  OPEN "TAXPROP.DAT" FOR RANDOM AS #1 LEN = RLen
  OPEN "StormFee.TXT" FOR OUTPUT AS #2

  RCnt = LOF(1) / RLen
  FOR TCnt = 1 TO RCnt
   GET #1, TCnt, RRec
   IF RRec.Deleted = 0 THEN
     IF RRec.StormFee > 0 THEN
       PCnt = PCnt + 1
       PRINT #2, QPTrim(STR$(RRec.StormFee)); v$; QPTrim(RRec.RealPin); v$; QPTrim(STR$(RRec.CustPin)); v$
       'IF PCnt MOD 10 = 0 THEN
       '  SLEEP
       'END IF
     END IF
   END IF
  NEXT
  CLOSE
  PRINT "Total:"; PCnt




