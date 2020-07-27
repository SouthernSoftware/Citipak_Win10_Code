PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "Russell County Water & Sewage Authority"
PRINT #UBRpt, TAB(20); "             P.O. Box 655"
PRINT #UBRpt, TAB(20); "        Castlewood, VA  24224"
PRINT #UBRpt, TAB(20); "         PHONE: (276)762-5411"
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); UBCustRec(1).CustName; "    "; LongNDate$
PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode; TAB(55); "Loc. No. "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Acct # "; Acct$
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Dear Customer:"
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Our records indicate that your water bill is PAST DUE.  All bills"
PRINT #UBRpt, TAB(10); "not paid by the due date specified on the bill, will be penalized"
PRINT #UBRpt, TAB(10); "10% of the current bill.  Since your bill is PAST DUE, immediate"
PRINT #UBRpt, TAB(10); "payment is required to avoid disconnection of your water service."
PRINT #UBRpt, TAB(10); "If payment is not received in our office by the end of the work day"
PRINT #UBRpt, TAB(10); "on "; PDate$; ", your water will be disconnected at our convenience."
PRINT #UBRpt, TAB(10); "Please be reminded that payments made on disconnection date at New"
PRINT #UBRpt, TAB(10); "Peoples Bank or Miners Exchange Bank without notification to Russell"
PRINT #UBRpt, TAB(10); "County Water and Sewage Authority are still subject to disconnection."
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "If disconnection is made for non-payment, there is a $100.00 "
PRINT #UBRpt, TAB(10); "non-refundable reconnection charge that must be paid in addition"
PRINT #UBRpt, TAB(10); "to the total amount due before service is restored."
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "If payment has been made, or you feel you have received this notice"
PRINT #UBRpt, TAB(10); "in error, please contact the office between the hours of 8:00 a.m."
PRINT #UBRpt, TAB(10); "and 4:30 p.m."
PRINT #UBRpt,
PRINT #UBRpt, TAB(35); "TOTAL DUE $"; USING "$$####.##"; CustBal#
PRINT #UBRpt,
PRINT #UBRpt, TAB(25); "NO PARTIAL PAYMENTS WILL BE ACCEPTED"
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Please understand that due to repeated requests, we will no longer"
PRINT #UBRpt, TAB(10); "be able to extend additional time nor will we be able to reduce the"
PRINT #UBRpt, TAB(10); "amount due on your account."
PRINT #UBRpt, CHR$(12);



