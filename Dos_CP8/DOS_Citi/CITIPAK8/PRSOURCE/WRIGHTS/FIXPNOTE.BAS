DEFINT A-Z

'$INCLUDE: 'PREmpRec.biw'

'$INCLUDE: 'PRFiles.bi'

  DIM Emp2Rec    AS EmpData2Type
  Emp2Len = LEN(Emp2Rec)

  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2Len

  NumOfRec = LOF(1) \ Emp2Len

  FOR Cnt = 1 TO NumOfRec
    GET #1, Cnt, Emp2Rec
    TBank$ = Emp2Rec.PreNoted + Emp2Rec.BankName
    LSET Emp2Rec.BankName = TBank$
    IF Emp2Rec.PreNoteFlag <> 0 THEN
      LSET Emp2Rec.PreNoted = "Y"
    ELSE
      LSET Emp2Rec.PreNoted = "N"
    END IF
    PUT #1, Cnt, Emp2Rec
  NEXT
  CLOSE

