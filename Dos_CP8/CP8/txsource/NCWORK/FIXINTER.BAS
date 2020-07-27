DEFINT A-Z
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2num (thedate$)

  '$INCLUDE: 'TaxCust.BI'

  CLS
  PRINT
  PRINT "Removing Transactions. . ."

  TDate = Date2num("09-13-2000")
  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumTR& = LOF(TaxTran) / TaxTranRecLen
  FOR Cnt& = 1 TO NumTR&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; NumTR&;
    GET TaxTran, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = TDate THEN
      IF TaxTran(1).TranType = 4 THEN
        IF TaxTran(1).BelongTo > 0 THEN
          GET TaxTran, TaxTran(1).BelongTo, TaxTran(2)
          IF TaxTran(2).TAXYEAR = 2000 THEN
            Found = Found + 1
            TaxTran(2).Revenue.interest = Round#(TaxTran(2).Revenue.interest - TaxTran(1).Amount)
            PUT TaxTran, TaxTran(1).BelongTo, TaxTran(2)
            TaxTran(1).CustomerRec = 0
            TaxTran(1).Amount = 0
            TaxTran(1).BelongTo = 0
            PUT TaxTran, Cnt&, TaxTran(1)
          END IF
        END IF
      END IF
    END IF
  NEXT

  CLOSE
  CALL KILLFile("FIXINTER.EXE")
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete."



'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

