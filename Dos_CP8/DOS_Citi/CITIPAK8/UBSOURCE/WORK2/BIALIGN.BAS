REM  alignment mask
2800 FOR lop = 1 TO 9:
       LPRINT lop
     NEXT lop
2811 LPRINT TAB(30); "XXXXXXXXXXX XXXXXXXXXXXXXXX"
     LPRINT TAB(30); "XXXXXXXXXXXXXXXXXXXXXXXXXXX"
     LPRINT TAB(30); "XXXXXXXXXXXXXXXXXXXXXXXXXXX"
     LPRINT TAB(30); "XXXXXXXXXXXXXXXXXXXX, XX XXXXX-XXXX"
     FOR lop = 1 TO 3
     LPRINT lop
     NEXT lop
2820 LPRINT TAB(6); "XX/XX/XX"; TAB(18); "XX/XX/XX";
     LPRINT TAB(44); "$#####.##"
     LPRINT 1
     LPRINT 2

2830 LPRINT TAB(14); "######"

2835 FOR lop = 1 TO 5:
     LPRINT lop
     NEXT lop
     LPRINT TAB(1); "XXXXXXXXXX XXXXXXXXXXXXX"; TAB(26); "XXXXXXXXXXXXXXXXXXXXXXX"; TAB(54); "######"
     LPRINT 1
     LPRINT 2
     LPRINT TAB(1); "########";
     LPRINT TAB(14); "########";
     LPRINT TAB(24); "########";
     LPRINT TAB(33); "W"; TAB(37); "$#####.##"
     LPRINT TAB(33); "S"; TAB(37); "$#####.##"
     LPRINT TAB(1); "########";
     LPRINT TAB(14); "########";
     LPRINT TAB(24); "########";
     LPRINT TAB(33); "IRR"; TAB(37); "$#####.##"
     LPRINT TAB(33); "LP"; TAB(37); "$#####.##"
     LPRINT TAB(33); "RC"; TAB(37); "$#####.##"
     LPRINT TAB(33); "SC"; TAB(37); "$#####.##"
     LPRINT TAB(33); "MV"; TAB(37); "$#####.##"
     LPRINT TAB(33); "NF"; TAB(37); "$#####.##"
2840 FOR lop = 1 TO 3
       LPRINT lop
     NEXT lop
2850 LPRINT TAB(1); "XX/XX/XX"; TAB(10); "XX/XX/XX"; TAB(19); "XX/XX/XX"; TAB(28); "XX";
     LPRINT TAB(30); "###.##";
     LPRINT TAB(36); "$$#####.##"
     LPRINT 1
     LPRINT "end"

