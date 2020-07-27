DECLARE SUB ShowProcessingScrn (RptTitle$)

DEFINT A-Z
SUB ZeroOffBalances

  UBLog " IN: Zero Offseting Balances"
  BlockClear
  Ok = MsgBox%("UBSETUP", "MUSTEXI3")

  SELECT CASE Ok
  CASE 2
    UBLog "ABORTED: Zero Offseting Balances"
    'GOTO ExitRelink
  END SELECT

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure

  ShowProcessingScrn "Zero Offseting Balances"

  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen





  CLOSE

END SUB

