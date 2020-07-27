DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  Date1 = Date2Num%("12-06-1999")
'  Date2 = Date2Num%("02-28-1997")
'  ActNo = 3822
  
  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCust(1)    AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  CustLen = LEN(UBCust(1))

  KillFile "UBTRANS.bak"
  NAME "UBTRANS.DAT" AS "UBTRANS.bak"

  UBFile1 = FREEFILE
  OPEN "UBTRANS.bak" FOR RANDOM SHARED AS UBFile1 LEN = UBTranRecLen

  UBFile2 = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile2 LEN = UBTranRecLen

  UBFile3 = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile3 LEN = CustLen

  TNumOfRecs& = LOF(UBFile1) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile1, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransDate = Date1 AND UBTranRec(1).TransType = TranUtilityBill THEN
      GET UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
      FOR zz = 1 TO 15
        UBCust(1).CurrRevAmts(zz) = uRound(UBCust(1).CurrRevAmts(zz) - UBTranRec(1).RevAmt(zz))
      NEXT
      CustRev# = 0
      FOR zz = 1 TO 15
        CustRev# = uRound#(CustRev# + UBCust(1).CurrRevAmts(zz))
      NEXT
      UBCust(1).CurrBalance = CustRev#
      UBCust(1).PrevBalance = 0
      PUT UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
      Removed = Removed + 1
    ELSE
      PUT UBFile2, , UBTranRec(1)
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
'PRINT " Copied:"; TTCnt&;
PRINT "Removed:"; Removed
END
KILL "UBTRANS.bak"
KILL "striptrn.exe"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

