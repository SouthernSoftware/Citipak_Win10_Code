DEFINT A-Z


 ' Now Real Estate
  FOR Snt& = 1 TO NumOfPropRecs
  GET PropTaxFile, Snt&, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).CustPin = 0
  PUT PropTaxFile, Snt&, PropertyRec(1)
  NEXT Snt&

