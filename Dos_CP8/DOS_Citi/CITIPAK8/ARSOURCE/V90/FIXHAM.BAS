DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS

  DIM SHARED ARCust(1) AS ARCustRecType
  DIM SHARED oARCust(1) AS ARCustRecType90

  ARCustLen = LEN(ARCust(1))
  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = ARCustLen
  FOR cnt = 1 TO (LOF(1) / ARCustLen)
  GET #1, cnt, oARCust(1)
  GOSUB SetData
  ARCust(1).CUSTNUMB = QPTrim$(STR$(cnt))

  PUT #1, cnt, ARCust(1)
  NEXT
  CLOSE
  END

SetData:

    ARCust(1).CUSTNUMB = oARCust(1).CUSTNUMB
    ARCust(1).SORTNAME = oARCust(1).SORTNAME
    ARCust(1).BILLNAME = oARCust(1).BILLNAME
    ARCust(1).ADDRESS1 = oARCust(1).ADDRESS1
    ARCust(1).ADDRESS2 = oARCust(1).ADDRESS2
    ARCust(1).CITY = oARCust(1).CITY
    ARCust(1).STATE = oARCust(1).STATE
    ARCust(1).ZIPCODE = oARCust(1).ZIPCODE
    ARCust(1).CUSTNAME = oARCust(1).CUSTNAME
    ARCust(1).CONTACT = oARCust(1).CONTACT
    ARCust(1).BILLCAT1 = oARCust(1).BILLCAT1
    ARCust(1).DESC1 = oARCust(1).DESC1
    ARCust(1).REV1 = oARCust(1).REV1
    ARCust(1).Fee1 = oARCust(1).Fee1
    ARCust(1).BILLCAT2 = oARCust(1).BILLCAT2
    ARCust(1).DESC2 = oARCust(1).DESC2
    ARCust(1).REV2 = oARCust(1).REV2
    ARCust(1).Fee2 = oARCust(1).Fee2
    ARCust(1).BILLCAT3 = oARCust(1).BILLCAT3
    ARCust(1).DESC3 = oARCust(1).DESC3
    ARCust(1).REV3 = oARCust(1).REV3
    ARCust(1).Fee3 = oARCust(1).Fee3
    ARCust(1).BILLCAT4 = oARCust(1).BILLCAT4
    ARCust(1).DESC4 = oARCust(1).DESC4
    ARCust(1).REV4 = oARCust(1).REV4
    ARCust(1).Fee4 = oARCust(1).Fee4
    ARCust(1).BILLCAT5 = oARCust(1).BILLCAT5
    ARCust(1).DESC5 = oARCust(1).DESC5
    ARCust(1).REV5 = oARCust(1).REV5
    ARCust(1).Fee5 = oARCust(1).Fee5
    ARCust(1).IssuanceFee = oARCust(1).IssuanceFee
    ARCust(1).CustLocation = oARCust(1).CustLocation
    ARCust(1).WPHONE = oARCust(1).WPHONE
    ARCust(1).FeeAmt = oARCust(1).FeeAmt
    ARCust(1).LICENSE = oARCust(1).LICENSE
    ARCust(1).VALID = oARCust(1).VALID
    ARCust(1).Inactive = "N"
    ARCust(1).PRORATE = 100
    ARCust(1).AcctBal = oARCust(1).AcctBal
    ARCust(1).IssueLicense = "Y" 'oARCust(1).IssueLicense
    ARCust(1).Deleted = oARCust(1).Deleted
    ARCust(1).FirstTrans = oARCust(1).FirstTrans
    ARCust(1).LastTrans = oARCust(1).LastTrans
    ARCust(1).RoomtoGrow = oARCust(1).RoomtoGrow

RETURN

