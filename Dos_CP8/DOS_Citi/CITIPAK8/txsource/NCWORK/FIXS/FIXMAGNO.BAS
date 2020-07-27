DEFINT A-Z
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2num (thedate$)

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Processing. . ."
  TDate = Date2num("02-04-2001")
  TDate2 = Date2num("02-04-2002")
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  FOR Cnt& = 1 TO LOF(TaxTran) / TaxTranRecLen
    GET TaxTran, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = TDate AND TaxTran(1).TranType = 2 THEN
      Fixed = Fixed + 1
      TaxTran(1).TransDate = TDate2
      PUT TaxTran, Cnt&, TaxTran(1)
    END IF

  NEXT

  CLOSE
  CALL KILLFile("FIXMAGNO.EXE")
  PRINT
  PRINT
  PRINT "Fixed:"; Fixed

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

