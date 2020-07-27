DEFINT A-Z
DECLARE FUNCTION QPValI% (Numb$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)

  '$INCLUDE: 'newcust.bi'

  TYPE SeqIdxBuffType
    RSeq     AS LONG
    RecNo    AS LONG
    Book     AS INTEGER
    Fill     AS STRING * 6
  END TYPE

  CLS
  PRINT
  PRINT
  PRINT
  PRINT "Resequencing Read Sequence Numbers. . ."

  REDIM UBCust(1 TO 2) AS NewUBCustRecType
  CustLen = LEN(UBCust(1))

  IdxRecLen = 4 'we are using a long integer

  IFile = FREEFILE
  OPEN "UBCUSTBK.IDX" FOR RANDOM AS #IFile LEN = 4
  NFile = FREEFILE
  OPEN "ubcust.new" FOR RANDOM AS #NFile LEN = CustLen

  OFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM AS #OFile LEN = CustLen

  PRINT
  PRINT "Restoring Original Read Sequence. . ."
  PRINT

  NumOfRecs = LOF(IFile) \ 4
  REDIM IdxBuff(1 TO NumOfRecs) AS SeqIdxBuffType
  FOR Cnt = 1 TO NumOfRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; " of"; NumOfRecs;
    GET #IFile, Cnt, IdxBuff(Cnt).RecNo
    GET #NFile, IdxBuff(Cnt).RecNo, UBCust(1)
    GET #OFile, IdxBuff(Cnt).RecNo, UBCust(2)
    UBCust(2).Seq = UBCust(1).Seq
    IdxBuff(Cnt).Book = QPValI(UBCust(2).Book)
    IdxBuff(Cnt).RSeq = UBCust(2).Seq
    PUT #OFile, IdxBuff(Cnt).RecNo, UBCust(2)
  NEXT

  PRINT
  PRINT
  PRINT "Resequencing. . ."

  NewSeq& = 0
  Start = 1
  WhatBook = IdxBuff(1).Book
  FOR Cnt = 2 TO NumOfRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; " of"; NumOfRecs;
    ThisBook = IdxBuff(Cnt).Book
    IF ThisBook <> WhatBook THEN
      Last = Cnt - 1
      NumElements = Last - Start
      IF NumElements > 1 THEN
        SortT IdxBuff(Start), NumElements, 0, 16, 0, -2
      END IF
      FOR SCnt = Start TO Last
        NewSeq& = NewSeq& + 50
        GET #OFile, IdxBuff(SCnt).RecNo, UBCust(1)
        UBCust(1).FillPad = STR$(UBCust(1).Seq)
        UBCust(1).Seq = NewSeq&
        PUT #OFile, IdxBuff(SCnt).RecNo, UBCust(1)
      NEXT
      Start = Last
      WhatBook = ThisBook
    END IF
  NEXT
  Last = NumOfRecs
  NumElements = Last - Start
'  PRINT
'  PRINT " Last:"; Last
'  PRINT "Start:"; Start
'  PRINT "NumEl:"; NumElements
'  a$ = INPUT$(1)

  IF NumElements > 1 THEN
    SortT IdxBuff(Start), NumElements, 0, 16, 0, -2
  END IF

  FOR Cnt = Start TO Last
    LOCATE , 1
    PRINT "Processing:"; Cnt; " of"; NumOfRecs;
    IF IdxBuff(Cnt).RecNo > 0 THEN
      NewSeq& = NewSeq& + 50
      GET #OFile, IdxBuff(Cnt).RecNo, UBCust(1)
      UBCust(1).FillPad = STR$(UBCust(1).Seq)
      UBCust(1).Seq = NewSeq&
      PUT #OFile, IdxBuff(Cnt).RecNo, UBCust(1)
    END IF
  NEXT

  CLOSE

