DEFINT A-Z
DECLARE FUNCTION Date2Num% (WhatDate$)
DECLARE FUNCTION Num2Date$ (WhatDate%)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newfa.BI'

'  REDIM test(1) AS FAItemRecTypeV1
'  PRINT LEN(test(1))
'  REDIM test2(1) AS FAItemRecType
'  PRINT LEN(test2(1))
'  END

  CLS
  PRINT "   FA FileChecking Utility V1.00"
  PRINT
  PRINT
  PRINT "   Checking Items File. . . ";

  REDIM Item1(1) AS FAItemRecTypeV1
  REDIM Item2(1) AS FAItemRecType

  ItemRecLen = LEN(Item2(1))
  ItemFile = FREEFILE

  OPEN "FAITEMS.DAT" FOR RANDOM AS ItemFile LEN = ItemRecLen
  NumOfRec& = LOF(ItemFile) / ItemRecLen
  IF NumOfRec& > 0 THEN
    GET ItemFile, 1, Item2(1)
    IF Item2(1).FileVer <> 2 THEN
      Col = POS(1)
      Row = CSRLIN
      LOCATE Row, Col
      PRINT "Converting.";
      FOR Cnt& = 1 TO NumOfRec&
        LOCATE Row + 1, Col
        PRINT "Processing:"; Cnt&; "of"; NumOfRec&;
        GET ItemFile, Cnt&, Item1(1)
        REDIM Item2(1) AS FAItemRecType
        Item2(1).ITEMTAG = Item1(1).ITEMTAG
        Item2(1).ISTATUS = Item1(1).ISTATUS
        Item2(1).DEPYN = " "
        Item2(1).AQURDATE = Item1(1).AQURDATE
        Item2(1).IDESC1 = Item1(1).IDESC1
        Item2(1).IDESC2 = Item1(1).IDESC2
        Item2(1).GLACCT = Item1(1).GLACCT
        Item2(1).IDEPT = Item1(1).IDEPT
        Item2(1).ASSETCODE = Item1(1).ASSETCODE
        Item2(1).CODEREC = Item1(1).CODEREC
        Item2(1).ILIFE = Item1(1).ILIFE
        Item2(1).ORGCOST = Item1(1).ORGCOST
        Item2(1).DEP2DATE = Item1(1).DEP2DATE
        Item2(1).CDEPDATE = Item1(1).CDEPDATE
        Item2(1).DISPDATE = Item1(1).DISPDATE
        Item2(1).VENDOR = Item1(1).VENDOR
        Item2(1).SERIALNO = Item1(1).SERIALNO
        Item2(1).ITEMMFG = Item1(1).ITEMMFG
        Item2(1).CONTACT = Item1(1).CONTACT
        LSET Item2(1).Fill1 = ""
        Item2(1).FileVer = 2
        IF Item1(1).AQURDATE > -32000 AND Item1(1).ILIFE > 0 THEN
          AQDate$ = Num2Date(Item1(1).AQURDATE)
          AQYear = VAL(RIGHT$(AQDate$, 4))
          AQYear = AQYear + Item1(1).ILIFE
          Item2(1).EOLDATE = Date2Num(LEFT$(AQDate$, 6) + QPTrim$(STR$(AQYear)))
        END IF
        PUT ItemFile, Cnt&, Item2(1)
      NEXT
    ELSE
      PRINT "OK."
    END IF
  ELSE
    PRINT "   NO RECORDS!"
  END IF
  PRINT
  PRINT
  PRINT
  PRINT "   Processing Complete"
  PRINT
  PRINT
  PRINT
  PRINT "Press Any Key. . . ";
  CLOSE
  DO
    zz$ = INKEY$
  LOOP UNTIL LEN(zz$)

