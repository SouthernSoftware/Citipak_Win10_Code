DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'TaxCust.BI'

  'BadDate = Date2Num("11-21-2005")
'  FixDate = Date2Num("12-03-2003")
'END

  CLS
  REDIM TaxTran(1 TO 2) AS TaxTransactionType

  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumTR& = LOF(TaxTran) / TaxTranRecLen

  LOCATE 3, 1
  PRINT "Correcting Transaction. . ."
  'FOR TRCnt& = 1 TO NumTR&
  TRCnt& = 143271
  GET #TaxTran, TRCnt&, TaxTran(1)
  GET #TaxTran, TaxTran(1).belongto, TaxTran(2)

  'STOP

  'TaxTran(2).Revenue.Principle1 = TaxTran(2).Revenue.Principle5
  'TaxTran(2).Revenue.Principle5 = 0
  TaxTran(2).Revenue.Principle1PD = 0
  TaxTran(2).Revenue.Principle2PD = 0
  TaxTran(2).Revenue.Principle3PD = 0
  TaxTran(2).Revenue.Principle4PD = 0
  TaxTran(2).Revenue.Principle5PD = 0
  TaxTran(2).Revenue.InterestPD = 0
  TaxTran(2).Revenue.PenaltyPD = 0
  SWAP TaxTran(2).Revenue.Principle5, TaxTran(2).Revenue.Principle1
  'TaxTran(2).Revenue.Principle2 = TaxTran(2).Revenue.Principle2 + TaxTran(1).Revenue.Principle2
  'TaxTran(2).Revenue.Principle2PD = TaxTran(2).Revenue.Principle2PD + TaxTran(1).Revenue.Principle2PD
  'TaxTran(2).Revenue.Principle3 = TaxTran(2).Revenue.Principle3 + TaxTran(1).Revenue.Principle3
  'TaxTran(2).Revenue.Principle3PD = TaxTran(2).Revenue.Principle3PD + TaxTran(1).Revenue.Principle3PD
  'TaxTran(2).Revenue.Principle4 = TaxTran(2).Revenue.Principle4 + TaxTran(1).Revenue.Principle4
  'TaxTran(2).Revenue.Principle4PD = TaxTran(2).Revenue.Principle4 + TaxTran(1).Revenue.Principle4

  'TaxTran(2).Revenue.Principle5 = TaxTran(2).Revenue.Principle5 - TaxTran(1).Revenue.Principle5
  'TaxTran(2).Revenue.Principle5PD = TaxTran(2).Revenue.Principle5PD + TaxTran(1).Revenue.Principle5PD

  PUT #TaxTran, TaxTran(1).belongto, TaxTran(2)

      TaxTran(1).Amount = 0
      TaxTran(1).Revenue.Interest = 0
      TaxTran(1).Revenue.Penalty = 0
      TaxTran(1).Revenue.Collection = 0
      TaxTran(1).Revenue.InterestPD = 0
      TaxTran(1).Revenue.PenaltyPD = 0
      TaxTran(1).Revenue.CollectionPD = 0
      TaxTran(1).Revenue.Principle1 = 0
      TaxTran(1).Revenue.Principle1PD = 0
      TaxTran(1).Revenue.Principle2 = 0
      TaxTran(1).Revenue.Principle2PD = 0
      TaxTran(1).Revenue.Principle3 = 0
      TaxTran(1).Revenue.Principle3PD = 0
      TaxTran(1).Revenue.Principle4 = 0
      TaxTran(1).Revenue.Principle4PD = 0
      TaxTran(1).Revenue.Principle5 = 0
      TaxTran(1).Revenue.Principle5PD = 0
      TaxTran(1).Description = "VOID VOID VOID VOID"
      PUT #TaxTran, TRCnt&, TaxTran(1)
  '  END IF
  'NEXT
  
  CLOSE
  LOCATE 6, 1
  PRINT "Done..."
  CALL killfile("fixhill.exe")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5000001) / 100
END FUNCTION

