DEFINT A-Z
LPRINT
LPRINT
LPRINT
LPRINT TAB(10); "Town of Lawrenceville"
LPRINT TAB(10); "400 N. Main Street"
LPRINT TAB(10); "Lawrenceville, VA  23868";
REM LPRINT TAB(36); "(804)-848-2414"
LPRINT TAB(55); DATE$
LPRINT
LPRINT
LPRINT
LPRINT
LPRINT
LPRINT TAB(10); nme$; TAB(55); "Acct # "; m$
LPRINT TAB(10); add1$; TAB(55); "Pastdue: $"; USING "#####.##"; pastdue# + curbal#
       chrin = INSTR(city$, " ")
       IF chrin = 0 OR chrin > 20 THEN chrin = 20
LPRINT TAB(10); LEFT$(city$, chrin); " "; st$; "  "; ZP$; TAB(55);
LPRINT
LPRINT
LPRINT
LPRINT
LPRINT
LPRINT TAB(15); "Your water service will be discontinued for lack of "
LPRINT TAB(10); "payment on "; LEFT$(offdate$, 2); "-"; MID$(offdate$, 3, 2); "-"; RIGHT$(offdate$, 2); ", at 12 o'clock unless arrangements"
LPRINT TAB(10); "have been made with this office before this date. If service"
LPRINT TAB(10); "is dicontinued a $10.00 cut on fee will be due plus the "
LPRINT TAB(10); "past due amount shown above."
LPRINT
LPRINT TAB(15); "The town office hours are from 9:00 A.M. until 5:00 P.M."
LPRINT TAB(10); "Monday through Friday. Phone number is (804)-848-2414."
LPRINT
LPRINT TAB(15); "If you have already mailed your payment please disregard"
LPRINT TAB(10); "this notice."
LPRINT
LPRINT
LPRINT
LPRINT TAB(40); "Water Department"
LPRINT TAB(40); "Town of Lawrenceville"
LPRINT
LPRINT CHR$(12); : RETURN

