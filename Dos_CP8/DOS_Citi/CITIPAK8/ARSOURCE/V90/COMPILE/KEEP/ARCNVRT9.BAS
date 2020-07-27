DEFINT A-Z

  ''$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS

  REDIM ARCust85(1) AS ARCustRecType85
  REDIM ARCust(1) AS ARCustRecType
  ARCust85Len = LEN(ARCust85(1))
  ARCustLen = LEN(ARCust(1))

  CLS
  PRINT
  PRINT " AR Conversion Utility v8.5 to v9.0"
  CALL KillFile("ARCSTBAK.DAT")
  NAME "arcust.dat" AS "ARCSTBAK.DAT"

  ARCust = FREEFILE
  OPEN "arcust.dat" FOR RANDOM AS ARCust LEN = ARCustLen
  ARCust85 = FREEFILE
  OPEN "ARCSTBAK.DAT" FOR RANDOM AS ARCust85 LEN = ARCust85Len
  NumOfRec& = LOF(ARCust85) / ARCust85Len
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumOfRec&;
    GET ARCust85, Cnt&, ARCust85(1)
    GOSUB DoARData
    PUT ARCust, Cnt&, ARCust(1)
  NEXT
  CLOSE
  LOCATE 9, 1
  PRINT "Conversion Complete."
  PRINT
  CALL KillFile("ARCNVRT9.EXE")
  END


DoARData:
    REDIM ARCust(1) AS ARCustRecType

    ARCust(1).CUSTNUMB = ARCust85(1).CUSTNUMB
    ARCust(1).SORTNAME = ARCust85(1).SORTNAME
    ARCust(1).BILLNAME = ARCust85(1).BILLNAME
    ARCust(1).ADDRESS1 = ARCust85(1).ADDRESS1
    ARCust(1).ADDRESS2 = ARCust85(1).ADDRESS2
    ARCust(1).CITY = ARCust85(1).CITY
    ARCust(1).STATE = ARCust85(1).STATE
    ARCust(1).ZIPCODE = ARCust85(1).ZIPCODE
    ARCust(1).CustName = ARCust85(1).CustName
    ARCust(1).Contact = ARCust85(1).Contact
    ARCust(1).BILLCAT1 = ARCust85(1).BILLCAT1
    ARCust(1).DESC1 = ARCust85(1).DESC1
    ARCust(1).REV1 = ARCust85(1).REV1
    ARCust(1).Fee1 = ARCust85(1).Fee1
    ARCust(1).BILLCAT2 = ARCust85(1).BILLCAT2
    ARCust(1).DESC2 = ARCust85(1).DESC2
    ARCust(1).REV2 = ARCust85(1).REV2
    ARCust(1).Fee2 = ARCust85(1).Fee2
    ARCust(1).BILLCAT3 = ARCust85(1).BILLCAT3
    ARCust(1).DESC3 = ARCust85(1).DESC3
    ARCust(1).REV3 = ARCust85(1).REV3
    ARCust(1).Fee3 = ARCust85(1).Fee3
    ARCust(1).BILLCAT4 = ARCust85(1).BILLCAT4
    ARCust(1).DESC4 = ARCust85(1).DESC4
    ARCust(1).REV4 = ARCust85(1).REV4
    ARCust(1).Fee4 = ARCust85(1).Fee4
    ARCust(1).BILLCAT5 = ARCust85(1).BILLCAT5
    ARCust(1).DESC5 = ARCust85(1).DESC5
    ARCust(1).REV5 = ARCust85(1).REV5
    ARCust(1).Fee5 = ARCust85(1).Fee5
    ARCust(1).IssuanceFee = ARCust85(1).IssuanceFee
    ARCust(1).CustLocation = ARCust85(1).CustLocation
    ARCust(1).WPHONE = ARCust85(1).WPHONE
    ARCust(1).FeeAmt = ARCust85(1).FeeAmt
    ARCust(1).LICENSE = ARCust85(1).LICENSE
    ARCust(1).VALID = ARCust85(1).VALID

    ARCust(1).Inactive = "N"

    ARCust(1).AcctBal = ARCust85(1).AcctBal
    ARCust(1).IssueLicense = ARCust85(1).IssueLicense
    ARCust(1).Deleted = ARCust85(1).Deleted
    ARCust(1).FirstTrans = ARCust85(1).FirstTrans
    ARCust(1).LastTrans = ARCust85(1).LastTrans
    ARCust(1).RoomtoGrow = ARCust85(1).RoomtoGrow


RETURN

