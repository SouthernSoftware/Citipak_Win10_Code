DEFINT A-Z
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2num (thedate$)

  '$INCLUDE: 'TaxCust.BI'

  TDate = Date2num("08-21-2000")
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  FOR Cnt& = 1 TO LOF(TaxTran) / TaxTranRecLen
    GET TaxTran, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate >= TDate THEN
      TaxTran(1).Posted2GL = ""
    ELSE
      TaxTran(1).Posted2GL = "Y"
    END IF
    PUT TaxTran, Cnt&, TaxTran(1)

  NEXT

  CLOSE
  CALL KILLFile("FIXMAGNO.EXE")
  PRINT "Done...."

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

