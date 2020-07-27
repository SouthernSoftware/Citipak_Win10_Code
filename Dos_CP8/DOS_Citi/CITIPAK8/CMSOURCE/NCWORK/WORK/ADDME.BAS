

PrintReceiptFunction:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  RecNumber = (LOF(CHandle) \ CMTrRecLen) + 1
  CLOSE CHandle

  'Set Printer Port for Receipts Printer if none set to 1

  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))

  OPEN "O", 15, "RECPT.PRN": WIDTH #15, 255

  PRINT #15, ""
  PRINT #15, TownName$
  PRINT #15, "LICENSE PAYMENT"
  PRINT #15, "Date: "; PostDate$
  PRINT #15,
  PRINT #15, "Account #"; Acct$
  PRINT #15, CustName$
  PRINT #15, Address$
  PRINT #15, Desc$
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; TotalOwed#
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; TotalCash#
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; TotalCheck#
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; TotalPaid#
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; TotalChange#
  PRINT #15,
  PRINT #15,
  PRINT #15, "Operator: "; OperRecNum%
  PRINT #15, "Receipt#: "; USING "######"; RecNumber
  PRINT #15,
  PRINT #15, TAB(7); "T H A N K   Y O U !"
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  CLOSE
  PrintRptFile Header$, "RECPT.PRN", GetDEFPort%, RetCode%, 5
RETURN


  IF Posted <> True THEN RETURN
  PostDate = Date2Num%(PostDate$)

  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  CMTRRec(1).TransDate = PostDate
  CMTRRec(1).TransAmount = TotalPaid#
  CMTRRec(1).TransCash = TotalCash#
  CMTRRec(1).TransCheck = TotalCheck#
  CMTRRec(1).TransAmtOwed = TotalOwed#
  CMTRRec(1).TransDesc = Desc$
  CMTRRec(1).TransSource = 41
  CMTRRec(1).TransName = CustName$
  CMTRRec(1).TransAcctNum = AcctRecord
  CMTRRec(1).TransDetNum = DetailRecord!
  CMTRRec(1).TransOperNum = Operator
  CMTRRec(1).TransPad = ""
  FldFactor = 0
  FOR Cnt = 1 TO 15
    CMTRRec(1).TransRevAmt(Cnt) = 0
  NEXT Cnt
  PUT CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTRRec(1)
  CLOSE CHandle

