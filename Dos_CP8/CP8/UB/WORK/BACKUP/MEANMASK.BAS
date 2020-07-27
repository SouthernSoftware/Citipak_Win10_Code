DEFINT A-Z
'**********************************
2800 REM alignment mask
2810 FOR LOP = 1 TO 5: LPRINT "": NEXT LOP
     LPRINT TAB(27); STRING$(8, "X");
     LPRINT TAB(36); STRING$(8, "X");
     LPRINT TAB(45); STRING$(8, "X");
     LPRINT TAB(60); "Acct#  XXXXX"
     LPRINT TAB(60); "Meter# XX-XXXXXXXXX"
     LPRINT TAB(42); "$XXXX.XX";
     LPRINT TAB(60); STRING$(20, "X")
     LPRINT TAB(42); "$XXXX.XX";
     LPRINT TAB(60); STRING$(20, "X")
     LPRINT TAB(42); "$XXXX.XX";
     LPRINT TAB(60); "XXXXXXXXX XX"
     LPRINT TAB(42); "XX/XX/XX";
     LPRINT TAB(60); "XXXXX-XXXX"
     LPRINT TAB(30); "Meter : XX-XXXXXXX";
     LPRINT TAB(60); "Net   : $XXXX.XX"
     LPRINT TAB(42); "$XXXX.XX";
     LPRINT TAB(60); "Gross : $XXXX.XX"
     LPRINT TAB(42); "$XXXX.XX"
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     LPRINT
'********************************

2510 FOR LOP = 1 TO 5: LPRINT "": NEXT LOP
     lc = 0
2520 LPRINT TAB(26); USING "#########"; past#;
      LPRINT TAB(35); USING "#########"; current#;
      printusage# = calcusage# / 100
       LPRINT TAB(44); USING "#########"; printusage#;
       LPRINT TAB(60); "Acct#"; TAB(66); acctrecord
       LPRINT TAB(60); "Meter#"; TAB(66); M$;
2530 PASTDUE# = INT((CVD(PASTDUE$) * 100) + .5) / 100
     IF PASTDUE# <> 0 THEN
      LPRINT TAB(42); USING "$####.##"; PASTDUE#;
     END IF
       IF LEFT$(firstname$, 2) = "  " THEN nme$ = lname$: GOTO 2540
     b = INSTR(firstname$, "  "): IF b = 0 THEN b = 20
     nme$ = LEFT$(firstname$, b) + lname$
2540 LPRINT TAB(60); LEFT$(nme$, 20)
     amtdue# = wateramount# + MISCAMOUNT# + memberamount# + tapfeeamount#
     LPRINT TAB(42); USING "$####.##"; amtdue#;
     LPRINT TAB(60); LEFT$(add1$, 20)
     amount# = wateramount# + PENALTYAMOUNT# + PASTDUE# + MISCAMOUNT# + memberamount# + tapfeeamount#
     LPRINT TAB(42); USING "$####.##"; amount#;
     LPRINT TAB(60); LEFT$(city$, 9); " "; st$
     LPRINT TAB(42); pastduedate$;
     LPRINT TAB(60); LEFT$(ZP$, 5); "-"; RIGHT$(ZP$, 4)
     LPRINT TAB(30); "Meter : "; TAB(39); M$;
     LPRINT TAB(60); "Net   :"; TAB(67); USING "$####.##"; amount#
     IF penflag = 1 THEN FEE# = percent * wateramount# / 100:
    LPRINT TAB(42); USING "$####.##"; FEE#;
     finalamount# = amount# + FEE#
     LPRINT TAB(60); "Gross :"; TAB(67); USING "$####.##"; finalamount#
     LPRINT TAB(42); USING "$####.##"; finalamount#
2690 LPRINT
     LPRINT TAB(30); MESS1$
     LPRINT TAB(30); MESS2$
     LPRINT TAB(30); MESS3$
     LPRINT
     LPRINT
     LPRINT
     RETURN

