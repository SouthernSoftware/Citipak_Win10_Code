'Whitakers Late notice
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 Town of Whitakers"
  PRINT #UBRpt, "                                    P.O. Box 727"
  PRINT #UBRpt, "                                 Whitakers NC 27891"
  PRINT #UBRpt, "                                      437-4011"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "RE: Account# "; Acct$
  PRINT #UBRpt, TAB(10); "   Location# "; Location$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     We have not received payment for the above referenced"
  PRINT #UBRpt, TAB(10); "water/sewer account.  Please note that if this bill is not"
  PRINT #UBRpt, TAB(10); "paid by the 19th, a penalty of $25 will be charged on the 20th."
  PRINT #UBRpt, TAB(10); "Water will be subject to cut-off without further notice!  Prior"
  PRINT #UBRpt, TAB(10); "to reconnection of this service, the town must receive payment"
  PRINT #UBRpt, TAB(10); "for the delinquent bill, and all penalties!"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact the Town Office"
  PRINT #UBRpt, TAB(10); "during normal working hours."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "please disregard this notice."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "PLEASE NOTE, THAT PAYMENT MUST BE RECEIVED PRIOR TO 3:30 P.M. TO"
  PRINT #UBRpt, TAB(10); "HAVE SERVICE RECONNECTED THE SAME DAY.  IF PAYMENT IS RECEIVED"
  PRINT #UBRpt, TAB(10); "AFTER 3:30, SERVICE WILL BE RECONNECTED THE NEXT REGULAR WORKING"
  PRINT #UBRpt, TAB(10); "DAY."
  'PRINT #UBRpt, TAB(10); ""
  'PRINT #UBRpt, TAB(10); "Town of Polkton Water Department."
  PRINT #UBRpt, CHR$(12)


  'PRINT #UBRpt, TAB(10); "     Please pay the amount of: "; FUsing$(STR$(TotalBal#), "$$,######.##"); " to prevent this"
  'PRINT #UBRpt, TAB(10); "scheduled service interruption."

