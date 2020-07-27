2800 REM alignment mask
     LPRINT CHR$(27); CHR$(48);
2810 FOR LL = 1 TO 9
     LPRINT "!"
2815 NEXT LL
2816 LPRINT TAB(3); "########";
     LPRINT TAB(15); "########";
     LPRINT TAB(25); "########"
     LPRINT
     LPRINT TAB(13); serfrom$; TAB(27); serto$
     LPRINT : LPRINT : LPRINT
     LPRINT TAB(3); "XX"; TAB(25); "#####.##"
     LPRINT
     LPRINT TAB(3); "XX"; TAB(25); "#####.##";
     LPRINT
     LPRINT TAB(52); "#####.##"
     LPRINT
     LPRINT TAB(52); "XXXXX"
     LPRINT TAB(25); "#####.##"
     LPRINT TAB(52); duedate$
     LPRINT
     LPRINT TAB(25); "#####.##"
     LPRINT TAB(37); "XXXXXXXXXXXXXXXXXXX"
     LPRINT TAB(25); "XXXXX";
     LPRINT TAB(37); "XXXXXXXXXXXXXXXXXXX"
     LPRINT TAB(37); "XXXXXXXXXXXXXXXXXXX"
     LPRINT TAB(25); duedate$;
     LPRINT TAB(37); "XXXXXXXX, XX  #####"
     LPRINT
     LPRINT TAB(2); message1$
     LPRINT TAB(2); message2$
     LPRINT
     LPRINT "~"

