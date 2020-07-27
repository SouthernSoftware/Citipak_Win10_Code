DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'

  REDIM FileName$(2)

  REDIM TaxCust(1) AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure

  FileName$(1) = RETaxCustFile
  FileName$(2) = PPTaxCustFile

  FOR WhatFile = 1 TO 2
    SELECT CASE WhatFile
    CASE 1
      LPRINT "Real Estate Customers"
    CASE 2
      LPRINT "Personal Property Customers"
    END SELECT
    TaxFile = FREEFILE
    OPEN FileName$(WhatFile) FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
    NumOfRecs& = LOF(TaxFile) \ TaxCustLen

    FOR Cnt& = 1 TO NumOfRecs&
      GET TaxFile, Cnt&, TaxCust(1)
      City$ = QPTrim$(TaxCust(1).City)
      CityLen = LEN(City$)
      IF CityLen > 0 THEN
        IF RIGHT$(City$, 1) = "." THEN
          CityLen = CityLen - 1              'reduce length by one
          City$ = MID$(City$, 1, CityLen)    'strip off the period
        END IF
        IF MID$(City$, CityLen - 2, 1) <> " " THEN
          LPRINT Cnt&, QPTrim$(TaxCust(1).FName); " "; QPTrim$(TaxCust(1).LName)
          GOTO SkipEm
        END IF
        State$ = RIGHT$(City$, 2)
        City$ = QPTrim$(MID$(City$, 1, CityLen - 2))  'strip off the state
        IF RIGHT$(City$, 1) = "," THEN
          City$ = LEFT$(City$, LEN(City$) - 1)
        END IF
        LSET TaxCust(1).City = City$
        LSET TaxCust(1).State = State$
        PRINT City$, State$
      END IF
      PUT TaxFile, Cnt&, TaxCust(1)
SkipEm:
    NEXT
    CLOSE
    LPRINT CHR$(12)
  NEXT

KillFile "fixtxcus.exe"

