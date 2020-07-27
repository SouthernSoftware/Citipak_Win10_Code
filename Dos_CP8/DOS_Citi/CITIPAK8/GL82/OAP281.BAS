DEFINT A-Z

  '$INCLUDE: 'AP.bi'                        'AP File Types

  REDIM APLedgerRec(1) AS APLedger81RecType
  REDIM OAPLedgerRec(1) AS APLedgerRecType
  OLdRecLen = LEN(OAPLedgerRec(1))
  LdRecLen = LEN(APLedgerRec(1))

  NAME "apledger.dat" AS "oledger.dat"

  OPEN "oLEDGER.DAT" FOR RANDOM AS #1 LEN = OLdRecLen
  OPEN "apledger.dat" FOR RANDOM AS #2 LEN = LdRecLen
  NumOfRec& = LOF(1) \ OLdRecLen
  FOR Cnt = 1 TO NumOfRec&
  GET #1, Cnt, OAPLedgerRec(1)

 APLedgerRec(1).VIN = OAPLedgerRec(1).VIN
 APLedgerRec(1).VendorCode = OAPLedgerRec(1).VendorCode
 APLedgerRec(1).VRecNum = OAPLedgerRec(1).VRecNum
 APLedgerRec(1).TrDate = OAPLedgerRec(1).TrDate
 APLedgerRec(1).GLDistDate = OAPLedgerRec(1).GLDistDate
 APLedgerRec(1).DUEDATE = OAPLedgerRec(1).DUEDATE

 APLedgerRec(1).TrCode = OAPLedgerRec(1).TrCode
 APLedgerRec(1).DOCNum = OAPLedgerRec(1).DOCNum
 APLedgerRec(1).PONUM = OAPLedgerRec(1).PONUM
 APLedgerRec(1).PAYCODE = OAPLedgerRec(1).PAYCODE
 APLedgerRec(1).PrintCode = OAPLedgerRec(1).PrintCode
 APLedgerRec(1).PDCheckNum = OAPLedgerRec(1).PDCheckNum
 APLedgerRec(1).PDCheckDate = OAPLedgerRec(1).PDCheckDate
 APLedgerRec(1).Comment = OAPLedgerRec(1).MiscCode
 APLedgerRec(1).DeptNumb = 0
 APLedgerRec(1).PSLFlag = ""
 APLedgerRec(1).Get1099 = ""
 APLedgerRec(1).Amt = OAPLedgerRec(1).Amt
 APLedgerRec(1).FrstDist = OAPLedgerRec(1).FrstDist
 APLedgerRec(1).LastDist = OAPLedgerRec(1).LastDist
 APLedgerRec(1).NextTrans = OAPLedgerRec(1).NextTrans
 APLedgerRec(1).TaxAmt = 0
 APLedgerRec(1).Pad = ""
 PUT #2, Cnt, APLedgerRec(1)
 NEXT
 CLOSE

