
  TotalCustBalance# = 0
  FOR RevAmts = 1 TO MaxRevsCnt
    TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec(1).CurrRevAmts(RevAmts))
  NEXT


  PRINT #RHandle, "  Acct Bal: "; USING "$$####,#.##"; TotalCustBalance#
