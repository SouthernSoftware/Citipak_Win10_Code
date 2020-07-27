DEFINT A-Z
DECLARE FUNCTION MakeSeqNumb$ (WhatSeq&)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

  '$INCLUDE: 'newcust.BI'
  
  CLS
  PRINT
  PRINT
  PRINT "Resequencing Customer Locations. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

  PRINT
  PRINT
  
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  WhatSeq& = 0

  FOR Cnt& = 1 TO IdxNumOfRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    ThisCust& = IdxBuff(Cnt&).RecNum
    GET UBFile, ThisCust&, UBCustRec(1)
    IF UBCustRec(1).DelFlag THEN
      UBCustRec(1).Book = ""
      UBCustRec(1).SeqNumb = ""
      PUT UBFile, ThisCust&, UBCustRec(1)
    ELSE
       ThisSeqNumb$ = QPTrim$(UBCustRec(1).SeqNumb)
       IF LEN(ThisSeqNumb$) > 0 THEN
         Seq$ = RIGHT$(ThisSeqNumb$, 5)
         Seq$ = Seq$ + "0"
         UBCustRec(1).SeqNumb = Seq$
         PUT UBFile, ThisCust&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE
PRINT
PRINT
PRINT
PRINT "Done."

