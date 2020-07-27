2510 FOR LOP = 1 TO 6: LPRINT "": NEXT LOP
     lc = 0
2530 IF LEFT$(firstname$, 2) = "  " THEN nme$ = lname$: GOTO 2540
     B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
     nme$ = LEFT$(firstname$, B) + lname$

2540 IF PAST# > 9999999 OR CURRENT# > 9999999 OR CALCUSAGE# > 9999999 THEN
     LPRINT TAB(2); USING "#######"; PAST# / 100;
      LPRINT TAB(10); USING "#######"; CURRENT# / 100;
       LPRINT TAB(19); USING "#######"; CALCUSAGE#;
        LPRINT TAB(29); USING "#####.##"; amount1#
      ELSE
     LPRINT TAB(2); USING "#######"; PAST#;
      LPRINT TAB(10); USING "#######"; CURRENT#;
       LPRINT TAB(19); USING "#######"; CALCUSAGE#;
        LPRINT TAB(29); USING "#####.##"; amount1#
     END IF

     IF amount2# <> 0 THEN
       LPRINT TAB(2); LEFT$(rev2name$, 14);
       LPRINT TAB(29); USING "#####.##"; amount2#;
     END IF
       LPRINT TAB(40); LEFT$(nme$, 20)

2550 IF AMOUNT3# <> 0 THEN
      LPRINT TAB(2); LEFT$(rev3name$, 14);
       LPRINT TAB(29); USING "#####.##"; AMOUNT3#;
     END IF
      LPRINT TAB(40); LEFT$(add1$, 20)

2560 IF AMOUNT4# <> 0 THEN
      LPRINT TAB(2); LEFT$(rev4name$, 14);
       LPRINT TAB(29); USING "#####.##"; AMOUNT4#;
     END IF
       LPRINT TAB(40); LEFT$(add3$, 20)

2570 IF AMOUNT5# <> 0 THEN
      LPRINT TAB(2); LEFT$(rev5name$, 14);
       LPRINT TAB(29); USING "#####.##"; AMOUNT5#;
     END IF
     LPRINT TAB(40); LEFT$(CITY$, 16); " "; ST$

     IF amount6# <> 0 THEN
      LPRINT TAB(2); LEFT$(rev6name$, 14);
       LPRINT TAB(29); USING "#####.##"; amount6#;
     END IF
     LPRINT TAB(50); LEFT$(ZP$, 5); "-"; MID$(ZP$, 6, 4)

LPRINT
LPRINT
LPRINT TAB(40); LEFT$(add2$, 20)
LPRINT
LPRINT TAB(9); LEFT$(meterreaddate$, 2); "/"; MID$(meterreaddate$, 3, 2); "/"; MID$(meterreaddate$, 5, 2);
 LPRINT TAB(27); billdate$
LPRINT
LPRINT
LPRINT TAB(2); USING "#####"; ACCTRECORD;
 LPRINT TAB(10); USING "#####.##"; pastdue#;
  LPRINT TAB(19); USING "#####.##"; amount# - pastdue#;
   LPRINT TAB(29); USING "#####.##"; amount#;
   LPRINT TAB(43); USING "#####"; ACCTRECORD;
    LPRINT TAB(51); USING "#####.##"; amount#
LPRINT
LPRINT TAB(2); mess1$

