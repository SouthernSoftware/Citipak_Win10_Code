DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'PROPAbst.BI'

  CLS
  PRINT
  PRINT "Correcting MAP/BLOCK/LOT. . ."
  REDIM TaxCustRec(1)  AS TaxCustType
  REDIM PropertyRec(1)  AS PropertyRecType

  CustLen = LEN(TaxCustRec(1))
  PropLen = LEN(PropertyRec(1))
  
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM SHARED AS #PropTaxFile LEN = PropLen
  NumOfPropRecs = LOF(PropTaxFile) / PropLen

  FOR Cnt& = 1 TO NumOfPropRecs
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfPropRecs;
    GET PropTaxFile, Cnt&, PropertyRec(1)
    ThisPin$ = PropertyRec(1).REALPIN
    Map$ = LEFT$(ThisPin$, 5)
    Block$ = MID$(ThisPin$, 6, 3)
    Lot$ = MID$(ThisPin$, 9)
    LSET PropertyRec(1).Map = QPTrim$(Map$)
    LSET PropertyRec(1).Block = QPTrim$(Block$)
    LSET PropertyRec(1).LOTNUMB = QPTrim$(Lot$)
    PUT PropTaxFile, Cnt&, PropertyRec(1)
  NEXT

  CLOSE
  LOCATE 9, 1
  PRINT "Correcting City/State. . ."
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = CustLen
  NumOfTaxRecs = LOF(TaxFile) / CustLen
  FOR Cnt& = 1 TO NumOfTaxRecs
    LOCATE 12, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfTaxRecs;
    IF Cnt& = 4 THEN GOTO SkipThisOne
    GET TaxFile, Cnt&, TaxCustRec(1)
    Temp$ = QPTrim$(TaxCustRec(1).LName)
    IF RIGHT$(Temp$, 1) = "," THEN
      LSET TaxCustRec(1).LName = MID$(Temp$, 1, (LEN(Temp$) - 1))
    END IF
    Temp$ = QPTrim$(TaxCustRec(1).SName)
    IF RIGHT$(Temp$, 1) = "," THEN
      LSET TaxCustRec(1).SName = MID$(Temp$, 1, (LEN(Temp$) - 1))
    END IF
    Temp$ = QPTrim$(TaxCustRec(1).City)
    CPos = INSTR(Temp$, ",")
    IF CPos > 0 THEN
      City$ = LEFT$(Temp$, CPos - 1)
      City$ = QPTrim$(City$)
      State$ = MID$(Temp$, CPos + 1)
      State$ = QPTrim$(State$)
      LSET TaxCustRec(1).City = City$
      LSET TaxCustRec(1).State = State$
    END IF
    PUT TaxFile, Cnt&, TaxCustRec(1)
SkipThisOne:
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "Processing Complete."

  CALL KillFile("Fixmontr.exe")

