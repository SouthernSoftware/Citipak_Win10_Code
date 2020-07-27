PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "Russell County Water & Sewage Authority"
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "             P.O. Box 655"
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "        Castlewood, VA  24224"
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); UBCustRec(1).CustName; "    "; LongNDate$
PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode; TAB(55); "Loc. No. "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Acct # "; Acct$
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Dear Customer:"
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Our records indicate that your water bill is PAST DUE.   All "
PRINT #UBRpt, TAB(10); "bills not paid by the 10th of the month are penalized 10% of the  "
PRINT #UBRpt, TAB(10); "current bill.  Since your bill is PAST DUE, immediate payment is"
PRINT #UBRpt, TAB(10); "required to avoid disconnection of your water service.  If payment"
PRINT #UBRpt, TAB(10); "is not received in our office by the end of the work day on"
PRINT #UBRpt, TAB(10); PDate$; ", your water will be disconnected at our convenience."
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "If disconnection is made for non-payment, there is a $20.00 "
PRINT #UBRpt, TAB(10); "non-refundable reconnection charge that must be paid in addition"
PRINT #UBRpt, TAB(10); "to the total amount due before service is restored."
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Please be advised that we will no longer send a final notice"
PRINT #UBRpt, TAB(10); "for disconnection.  The disconnection notice will be printed"
PRINT #UBRpt, TAB(10); "on your card if you have a past due balance.  Please refer to"
PRINT #UBRpt, TAB(10); "your monthly statement for amount due and disconnection date."
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "If payment has been made, or you feel you have received this notice"
PRINT #UBRpt, TAB(10); "in error, please contact the office between the hours of 8:00 a.m."
PRINT #UBRpt, TAB(10); "and 4:00 p.m."
PRINT #UBRpt,
PRINT #UBRpt, TAB(35); "TOTAL DUE $"; USING "$$####.##"; CustBal#
PRINT #UBRpt,
PRINT #UBRpt, TAB(25); "NO PARTIAL PAYMENTS WILL BE ACCEPTED"
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Please understand that due to repeated requests, we will no longer"
PRINT #UBRpt, TAB(10); "be able to extend additional time nor will we be able to reduce the"
PRINT #UBRpt, TAB(10); "amount due on your account."
PRINT #UBRpt, CHR$(12);



