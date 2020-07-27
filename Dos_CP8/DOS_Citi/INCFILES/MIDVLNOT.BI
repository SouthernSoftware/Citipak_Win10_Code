  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                                CITY OF MIDVILLE"
  PRINT #UBRpt, "                                  P.O. Box 234"
  PRINT #UBRpt, "                               MIDVILLE, GA 30441"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    "; UBCustRec(1).CustName; TAB(60); "ACCOUNT #"; Acct$
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                                REMINDER NOTICE"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Our records indicate that your account is past due. Your past due"
  PRINT #UBRpt, USING "    amount is $$#####,#.## If your bill is not paid by 5:00 p.m. on "; CustBal#
  PRINT #UBRpt, "    "; LongPDate$; ", your water service will be discontinued."
  PRINT #UBRpt, "    A $25.00 reconnect fee will be added to your account."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    The past due bill and $25.00 fee must be paid in full before service"
  PRINT #UBRpt, "    will be reconnected."
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    If you need to make payment arrangements, Please call the City"
  PRINT #UBRpt, "    Office before the cutoff date, so I can setup a payment schedule"
  PRINT #UBRpt, "    for you."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Thank you,"
  PRINT #UBRpt, "    CITY OF MIDVILLE"
  PRINT #UBRpt, CHR$(12)

