DEFINT A-Z

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'


  DIM SHARED TaxCustRec(1) AS TaxCustType
  'DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType

  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))

  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  GET TaxFile, CustRecNo, TaxCustRec(1)

  PValue# = 0
  IF NOT TaxCustRec(1).Deleted THEN
    'Figure Values
    'Real Value First
    PersValue# = 0
    ExmpDiscI# = 0
    ExmpDiscO# = 0
    DidCnt = 0
    BeenDone = 0
    IF TaxCustRec(1).FirstPersRec > 0 THEN
      PersRecord! = TaxCustRec(1).FirstPersRec
      WHILE PersRecord! > 0
        GET PersTaxFile, PersRecord!, PersRec(1)
        DidCnt = DidCnt + 1
        PERSVAL
        MHVALUE
        MCVALUE
        CVALUE
        MTVALUE

        IF PropertyRec(1).PROPVALU > 0 THEN
          PersValue# = Round#(PersValue# + PropertyRec(1).PROPVALU)
        END IF
        IF PropertyRec(1).EXMPSENI > 0 THEN
          ExmpDiscI# = Round#(ExmpDiscI# + PropertyRec(1).EXMPSENI)
        END IF
        IF PropertyRec(1).EXMPOTHR > 0 THEN
          ExmpDiscO# = Round#(ExmpDiscO# + PropertyRec(1).EXMPOTHR)
        END IF

        PropertyRecord! = PropertyRec(1).NextRec

'            IF PropertyRecord! > 0 THEN
'              Multi = Multi + 1
'              IF BeenDone = 0 THEN
'                LPRINT TaxCustRec(1).Acct
'                BeenDone = -1
'              END IF
'              'STOP
'            END IF
          WEND
        END IF

