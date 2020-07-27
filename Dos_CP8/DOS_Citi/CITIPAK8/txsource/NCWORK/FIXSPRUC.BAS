DECLARE FUNCTION TRound# (N#)
DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'

  CLS
  PRINT "Scanning Tax Accounts. . ."
  REDIM TaxCust(1 TO 2) AS TaxCustType
  TaxCustRecLen = LEN(TaxCust(1))

  OPEN "TaxCust.DAT" FOR RANDOM SHARED AS #1 LEN = TaxCustRecLen

  NumOfRec& = LOF(1) / TaxCustRecLen

  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;

    GET #1, Cnt&, TaxCust(1)

    GOSUB CheckBalance
    IF Balance# = 0 THEN
      IF TaxCust(1).Active <> "N" THEN
        TaxCust(1).Active = "N"    'update the active flag now.
        PUT #1, Cnt&, TaxCust(1)  'put it back to disk
        Fixed = Fixed + 1
      END IF
    END IF
 
DoneWithThisOne:

  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "   Changed:"; Fixed
  CALL KILLFile("FIXSPRUC.EXE")
  END

CheckBalance:
  REDIM TaxTrans(1) AS TaxTransactionType
  IF TaxCust(1).LastTrans = 0 THEN Balance# = 0: RETURN
  'Open the File and Look For A Balance
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  TransRecord& = TaxCust(1).LastTrans
  WHILE TransRecord& <> 0
    GET TransFile, TransRecord&, TaxTrans(1)
    IF TaxTrans(1).TranType = 1 THEN
      Balance# = TaxTrans(1).Revenue.Principle1 + TaxTrans(1).Revenue.Principle2 + TaxTrans(1).Revenue.Principle3 + TaxTrans(1).Revenue.Principle4 + TaxTrans(1).Revenue.Principle5
      Balance# = Balance# + TaxTrans(1).Revenue.Interest + TaxTrans(1).Revenue.PENALTY + TaxTrans(1).Revenue.Collection
      Balance# = Balance# - (TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd)
      Balance# = Balance# - (TaxTrans(1).Revenue.InterestPD + TaxTrans(1).Revenue.PenaltyPd + TaxTrans(1).Revenue.CollectionPd)
      Balance# = TRound#(Balance#)
      IF Balance# > 0 THEN
        CLOSE TransFile
        RETURN
      END IF
    END IF
    TransRecord& = TaxTrans(1).LastTrans
  WEND
  Balance# = 0
  CLOSE TransFile
RETURN

FUNCTION TRound# (N#)
  TRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

