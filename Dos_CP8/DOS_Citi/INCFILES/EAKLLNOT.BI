  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); NDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    "; UBCustRec(1).CustName
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    RE:  UTILITY CUT OFF NOTICE"
  PRINT #UBRpt, "          AMOUNT:"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt, "          ACCOUNT: #"; Acct$;
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Our records indicate that you have failed to make the payment on"
  PRINT #UBRpt, "    your utility bill for the Town of Eakly."
  PRINT #UBRpt, "    To prevent service from being disconnected, payment of the total"
  PRINT #UBRpt, "    amount must be received by this office by no later than 1:00 P.M."
  PRINT #UBRpt, "    on "; PDate$; ".  Once the services have been disconnected,"
  PRINT #UBRpt, "    you will be required to pay the total amount due, and a $75.00"
  PRINT #UBRpt, "    hook-up fee, before your services will be continued."
  PRINT #UBRpt, "    Please have payment in our office by your due date.  Please allow"
  PRINT #UBRpt, "    enough time for your payment to reach our office by your due date,"
  PRINT #UBRpt, "    we are not responsible for late mail.  This notice takes priority"
  PRINT #UBRpt, "    over any bill received."
  PRINT #UBRpt, "    If you believe the amount is not correct, you may meet or call us"
  PRINT #UBRpt, "    between 9:00-4:00, Monday thru Friday, before your due date.  The"
  PRINT #UBRpt, "    number we can be reached at is 405-797-3252.  We do have the "
  PRINT #UBRpt, "    authority to adjust your bill in case of error."
  PRINT #UBRpt,
  PRINT #UBRpt, "    THIS IS YOUR FINAL NOTICE,  OUR NEXT ACTION WILL BE TO POST A DOOR"
  PRINT #UBRpt, "    KNOCKER AT ADDRESS AND TURN OFF THE UTILITY SERVICES."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Hopefully this matter will be resolved before such action is necessary."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Town of Eakly"
  PRINT #UBRpt, "    Utility Department"
  PRINT #UBRpt, "    Credit & Collections"
  PRINT #UBRpt, CHR$(12)

