DEFINT A-Z

  '$INCLUDE: 'newcust.bi'

  CLS
  PRINT
  PRINT
  PRINT
  PRINT "Resequencing Read Sequence Numbers. . ."

  DIM UBCust AS NewUBCustRecType
  CustLen = LEN(UBCust)

  IdxRecLen = 4 'we are using a long integer

  OPEN "UBCUSTBK.IDX" FOR RANDOM AS #1 LEN = 4
  NumOfRecs = LOF(1) \ 4
  REDIM IdxBuff(1 TO NumOfRecs) AS LONG
  FOR Cnt = 1 TO NumOfRecs
    GET #1, Cnt, IdxBuff(Cnt)
  NEXT
  CLOSE

  OPEN "ubcust.dat" FOR RANDOM AS #1 LEN = CustLen
  NumOfRecs = LOF(1) \ CustLen

  FOR Cnt = 1 TO NumOfRecs
    LOCATE 6, 1
    PRINT "Processing:"; Cnt; " of"; NumOfRecs;
    ThisCust& = IdxBuff(Cnt)
    GET #1, ThisCust&, UBCust
    IF UBCust.DelFlag = 0 THEN
      SeqCnt& = SeqCnt& + 1
      Seq& = SeqCnt& * 50
      UBCust.Seq = Seq&
      PUT #1, ThisCust&, UBCust
    END IF
  NEXT
  CLOSE


