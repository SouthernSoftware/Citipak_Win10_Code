LPRINT : LPRINT : LPRINT
LPRINT TAB(20); "Russell County Water & Sewage Authority"
LPRINT
LPRINT TAB(20); "             P.O. Box 655"
LPRINT
LPRINT TAB(20); "        Castlewood, VA  24224"
LPRINT
LPRINT
LPRINT TAB(10); nme$; TAB(55); DATE$
LPRINT TAB(10); add1$
LPRINT TAB(10); city$; " "; st$; "  "; ZP$; TAB(55); "Loc. No. "; M$
LPRINT
LPRINT TAB(10); "Acct # "; 1234  'CVS(mrec$)
LPRINT
LPRINT TAB(10); "Dear Customer:"
LPRINT
LPRINT TAB(10); "Our records indicate that your water bill is in arrears.  All "
LPRINT TAB(10); "bills not paid by the 10th of the month are penalized 10% of the  "
LPRINT TAB(10); "current bill.  Since your bill is past due, immediate payment is"
LPRINT TAB(10); "due to avoid disconnection of your water service.  If payment is"
LPRINT TAB(10); "not received in our office by the end of the work day on "
LPRINT TAB(10); dueby$; ", your water shall be disconnected at our convenience."
LPRINT
LPRINT TAB(10); "If disconnection is made for non-payment, there is a $20.00 "
LPRINT TAB(10); "non-refundable reconnection charge that must be paid in addition"
LPRINT TAB(10); "to the total amount due before service is restored."
LPRINT
LPRINT TAB(10); "Please be advised that we will no longer send a final notice"
LPRINT TAB(10); "for disconnection.  The disconnection notice will be printed"
LPRINT TAB(10); "on your card if you have a past due balance.  Please refer to"
LPRINT TAB(10); "your monthly statement for amount due and disconnection date."
LPRINT
LPRINT TAB(10); "If payment has been made or you feel you have received this notice"
LPRINT TAB(10); "in error, please contact the office between the hours of 8:00 a.m."
LPRINT TAB(10); "and 4:00 p.m."
LPRINT
LPRINT TAB(40); "TOTAL DUE $"; USING "$$####.##"; pastdue# + curbal#
LPRINT
LPRINT TAB(30); "NO PARTIAL PAYMENTS WILL BE ACCEPTED"
LPRINT
LPRINT TAB(10); "Please understand that due to repeated requests, we will no longer"
LPRINT TAB(10); "be able to extend additional time nor will we be able to reduce the"
LPRINT TAB(10); "amount due on your account."
LPRINT CHR$(12);

