DEFINT A-Z
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB MPaintBox (TRow, TCol, Colr, BRow, RCol)
DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB ClrBG ()

SUB CheckData
'St1# = TIMER
'St1# = St1# + .01
'DO
'  st2# = TIMER
'LOOP UNTIL st2# > St1#
END SUB

SUB ClrBG

  LOCATE , , 0
  FOR Row = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Row, 1, 3
  NEXT
  QPrintRC SPACE$(80), 25, 1, 112
  QPrintRC "Wait", 25, 2, 112

END SUB

SUB MnuTitle (Title$, x, Y, Wdth, TitColor)
 
 TRow = x - 1
 W = Wdth + 4
 SCol = Y + W
 'T$ = "- " + Title$ + " -"

 TitleCol = (Y + ((W / 2) - (LEN(Title$) / 2)))
 IF NOT TitleCol MOD 2 THEN
   TitleCol = TitleCol + 1
 END IF
 'TitleCol = y + 1

 QPrintRC SPACE$(W), TRow, Y, TitColor
 QPrintRC Title$, TRow, TitleCol, TitColor
 MPaintBox x, SCol, x, SCol + 1, 8


END SUB

