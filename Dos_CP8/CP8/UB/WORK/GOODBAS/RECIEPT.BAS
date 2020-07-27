DEFINT A-Z
OPEN "com1:" FOR OUTPUT AS #1
PRINT #1, CHR$(&H1B); "@"; 'Initializes the printer (ESC @)
PRINT #1, CHR$(&H1B); "a"; CHR$(1); 'Specifies a centered printing position (ESC a)
PRINT #1, CHR$(&H1B); "!"; CHR$(0); 'Specifies font A (ESC !)
PRINT #1, "January 14, 2002 15:00";
PRINT #1, CHR$(&H1B); "d"; CHR$(3); 'Prints and 3 line feeding (ESC d)
PRINT #1, CHR$(&H1B); "a"; CHR$(0); 'Selects the left print position (ESC a)
PRINT #1, CHR$(&H1B); "!"; CHR$(1); 'Selects font B
PRINT #1, "TM-U210B $20.00"; CHR$(&HA);
PRINT #1, "TM-U210D $21.00"; CHR$(&HA);
PRINT #1, "PS-170 $17.00"; CHR$(&HA);
PRINT #1, CHR$(&HA); 'Line feeding (LF)
PRINT #1, CHR$(&H1B); "!"; CHR$(17); 'Selects double-height mode
PRINT #1, "TOTAL $58.00"; CHR$(&HA);
PRINT #1, CHR$(&H1B); "!"; CHR$(0); 'Cancels double-height mode
PRINT #1, "------------------------------"; CHR$(&HA);
PRINT #1, "PAID $60.00"; CHR$(&HA);
PRINT #1, "CHANGE $ 2.00"; CHR$(&HA);
PRINT #1, CHR$(&H1D); "V"; CHR$(66); CHR$(0); 'Feeds paper & cut
'Drawer Kick (ESC p)
PRINT #1, CHR$(&H1B); CHR$(&H70); CHR$(&H0); CHR$(60); CHR$(120);
CLOSE

