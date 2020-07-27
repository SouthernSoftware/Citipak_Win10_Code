PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "     Taloga Public Works Authority"
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "             P.O. Box 307"
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "          Taloga OK,  73667"
PRINT #UBRpt,
PRINT #UBRpt, TAB(20); "        PHONE: (580)328-5444"
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); UBCustRec(1).CustName; "    "; LongNDate$
PRINT #UBRpt, TAB(8); UBCustRec(1).ADDR1
PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).CITY); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode; TAB(55); "Loc. No. "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "Acct # "; Acct$
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "Dear Customer:"
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "Our records indicate that your utility bill is PAST DUE. All bills"
PRINT #UBRpt, TAB(8); "not paid by the due date specified on the bill, will be penalized"
PRINT #UBRpt, TAB(8); "10% of the current bill.  Since your bill is PAST DUE, Immediate"
PRINT #UBRpt, TAB(8); "payment is required to avoid disconnection of your utility service."
PRINT #UBRpt, TAB(8); "If payment is not received in our office by 12:00 noon on "; PDate$
PRINT #UBRpt, TAB(8); "your utilities will be disconnected at our convenience."
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "If disconnection is made for non-payment, there is a $30.00 "
PRINT #UBRpt, TAB(8); "non-refundable reconnection charge that must be paid in addition"
PRINT #UBRpt, TAB(8); "to the total amount due before service is restored."
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "Please be advised that we will not send a final disconnect notice."
PRINT #UBRpt, TAB(8); "Please refer to your monthly utility bill card for amount due and"
PRINT #UBRpt, TAB(8); "disconnect date."
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "If payment has been made, or you feel you have received this notice"
PRINT #UBRpt, TAB(8); "in error, please contact the office between the hours of 8:00 a.m."
PRINT #UBRpt, TAB(8); "and 4:00 p.m."
PRINT #UBRpt,
PRINT #UBRpt, TAB(35); "TOTAL DUE $"; USING "$$####.##"; CustBal#
PRINT #UBRpt,
PRINT #UBRpt, TAB(25); "NO PARTIAL PAYMENTS WILL BE ACCEPTED"
PRINT #UBRpt,
PRINT #UBRpt, TAB(8); "IF YOU BELIEVE THE AMOUNT DUE IS NOT CORRECT, YOU MAY CONTACT THE"
PRINT #UBRpt, TAB(8); "TOWN ADMINISTRATOR, BETWEEN 8:00 A.M. AND 4:00 P.M MONDAY THROUGH"
PRINT #UBRpt, TAB(8); "FRIDAY AT THE TOWN HALL. THE TOWN ADMINISTRATOR HAS THE AUTHORITY TO "
PRINT #UBRpt, TAB(8); "ADJUST A CUSTOMER'S BILL IN CASE OF AN ERROR OR, APPROVE A PAYMENT "
PRINT #UBRpt, TAB(8); "SCHEDULE FOR ELIGIBLE CUSTOMERS."
PRINT #UBRpt, CHR$(12);




