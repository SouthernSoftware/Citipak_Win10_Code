DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)

DEFINT A-Z
SUB DMVTest
  REDIM DMVHeader(1) AS DMVHeader
  REDIM DMVRecord(1) AS DMVRecord
  REDIM TaxSetUp(1) AS TaxMasterType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM PersRec(1) AS PersonalRecType
  
  DIM SSN1(50), LastName1(50), FirstName1(50), Init1(50), Addr1(50), Addr2(50), City(50), State(50), Zip(50), VIN(50), VehValue(50), PPTaxPd(50), PPTaxReimb(50)
  
  ' Date Calculations Here
  Today = Date2Num%(DATE$)
  Today$ = RIGHT$(DATE$, 4) + MID$(DATE$, 4, 2) + LEFT$(DATE$, 2)
  Cdate = Date2Num%("12-31-1997")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$
  PayJourName$ = "TAAND" + JulianDate$ + ".001"
  Header$ = "Creating DMV Test Data"
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR RANDOM AS RptHandle LEN = LEN(DMVHeader(1))
  
  ShowProcessingScrn Header$
  
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
  
  
  FOR Cnt = 1 TO NumOfTaxRecs
    CustRecNo = Cnt
    GET TaxFile, CustRecNo, TaxCustRec(1)
    
    IF NOT TaxCustRec(1).Deleted THEN
      
      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name
      
      
      'Personal Property Here
      PersValue# = 0
      IF TaxCustRec(1).FirstPersRec > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstPersRec
        WHILE PropertyRecord! <> 0
          GET PersTaxFile, PropertyRecord!, PersRec(1)
          IF LEFT$(Ptype$, 1) <> "R" THEN
            Discnt# = Discnt# + PersRec(1).EXMPSENI
          END IF
          PropertyRecord! = PersRec(1).NextRec
        WEND
      END IF
      
      ' SSN1 (50), LastName1(50), FirstName1(50), Init1(50), Addr1(50), Addr2(50), City(50), State(50), Zip(50), VIN(50), VehValue(50), PPTaxPd(50), PPTaxReimb(50)
    END IF
    ShowPctComp DoneCnt, TotalRecs&
  NEXT
  
  
  
  CLOSE
  
  LPTPORT = 1
  PrintRptFile Header$, PayJourName$, LPTPORT%, RetCode%, EntryPoint
  
  KillFile PayJourName$
  
DMVExitJournal:
  EXIT SUB
  
DMVPrintRptHeader:
  DMVHeader(1).RecordType = "H"
  DMVHeader(1).Batch = "1"
  DMVHeader(1).Jury = "AAND"
  DMVHeader(1).TaxYear = "1998"
  DMVHeader(1).SubDate = Today$
  DMVHeader(1).TotalVeh = Vehicles$
  DMVHeader(1).TotalAmt = TotalAmt$
  DMVHeader(1).Filler = ""
  PUT RptHandle, 1, DMVHeader(1)
  RETURN
  
DMVPrintRptEnding:
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, FF$
  RETURN
  
END SUB

