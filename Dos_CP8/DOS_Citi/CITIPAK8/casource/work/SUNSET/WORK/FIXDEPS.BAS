DEFINT A-Z

  '$INCLUDE: 'ubpaymnt.BI'

  REDIM UBPaymentRec(1)     AS UBPaymentRecType
  REDIM oUBPaymentRec(1)    AS oUBPaymentRecType

  UBPayRecLen = LEN(UBPaymentRec(1))
  oUBPayRecLen = LEN(UBPaymentRec(1))

  NAME "UBDEP1.DAT" AS "UBDEP1.old"

  OldFileName$ = "UBDEP1.OLD"
  PayFileName$ = "UBDEP1.DAT"

  OldD = FREEFILE
  OPEN OldFileName$ FOR RANDOM SHARED AS OldD LEN = oUBPayRecLen

  NewD = FREEFILE
  OPEN PayFileName$ FOR RANDOM SHARED AS NewD LEN = UBPayRecLen
  NumOfRec = LOF(OldD) / oUBPayRecLen

  FOR Cnt = 1 TO NumOfRec
    GET OldD, Cnt, oUBPaymentRec(1)
    UBPaymentRec(1).OPERNUM = oUBPaymentRec(1).OPERNUM
    UBPaymentRec(1).PAYDATE = oUBPaymentRec(1).PAYDATE
    UBPaymentRec(1).CUSTACCT = oUBPaymentRec(1).CUSTACCT
    UBPaymentRec(1).CUSTNAME = oUBPaymentRec(1).CUSTNAME
    UBPaymentRec(1).CUSTADDR = oUBPaymentRec(1).CUSTADDR
    UBPaymentRec(1).CUSTCMNT = oUBPaymentRec(1).CUSTCMNT
    UBPaymentRec(1).TaxExempt = oUBPaymentRec(1).TaxExempt
    UBPaymentRec(1).AMTOWED = oUBPaymentRec(1).AMTOWED
    UBPaymentRec(1).TENDERTY = oUBPaymentRec(1).TENDERTY
    UBPaymentRec(1).CASHAMT = oUBPaymentRec(1).CASHAMT
    UBPaymentRec(1).CHKAMT = oUBPaymentRec(1).CHKAMT
    UBPaymentRec(1).AMTRECD = oUBPaymentRec(1).AMTRECD
    UBPaymentRec(1).CHANGE = oUBPaymentRec(1).CHANGE
    UBPaymentRec(1).DESC = oUBPaymentRec(1).DESC
    FOR zz = 1 TO 15
      UBPaymentRec(1).PaidOwed(zz).AmtOwe1 = oUBPaymentRec(1).PaidOwed(zz).AmtOwe1
      UBPaymentRec(1).PaidOwed(zz).AmtPd1 = oUBPaymentRec(1).PaidOwed(zz).AmtPd1
    NEXT
    UBPaymentRec(1).TOTOWED = oUBPaymentRec(1).TOTOWED
    UBPaymentRec(1).AMTPAID = oUBPaymentRec(1).AMTPAID
    UBPaymentRec(1).Status = ""
    PUT NewD, Cnt, UBPaymentRec(1)
  NEXT
  CLOSE
  CALL KillFile(OldFileName$)
  CALL KillFile("FixDeps.exe")

