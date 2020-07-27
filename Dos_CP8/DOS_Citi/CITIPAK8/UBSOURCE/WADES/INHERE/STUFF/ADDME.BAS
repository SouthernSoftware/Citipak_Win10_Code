    FOR lll = 1 TO copies

    FOR ll = 1 TO 6: LPRINT " ": NEXT ll
    LPRINT TAB(14); "W O R K   O R D E R   :   U T I L I T Y   D E P T ."
    LPRINT
    LPRINT "    Date Issued: "; DATE$
    LPRINT "  Work Order # : "; USING "######"; nextorder
    LPRINT "    Location # : "; M$
    LPRINT "       Meter # : "; METERNUM$
    LPRINT "  Customer Name: "; nme$
    LPRINT "Service Address: "; add2$
    LPRINT STRING$(79, "-")
    LPRINT "Instruction or Description of Work Needed"
    LPRINT STRING$(79, "-")
    LPRINT line1$
    LPRINT line2$
    LPRINT line3$
    LPRINT line4$
    LPRINT line5$
    LPRINT line6$
    LPRINT line7$
    LPRINT line8$
    LPRINT
    LPRINT
    LPRINT "Remarks Noted by Worker"
    LPRINT
    LPRINT STRING$(79, "_")
    LPRINT
    LPRINT STRING$(79, "_")
    LPRINT
    LPRINT STRING$(79, "_")
    LPRINT
    LPRINT STRING$(79, "_")
    LPRINT
    LPRINT STRING$(79, "_")
    LPRINT
    LPRINT
    LPRINT "BY: ______________________________   DATE: ____________________"
    LPRINT CHR$(12);
   NEXT lll

