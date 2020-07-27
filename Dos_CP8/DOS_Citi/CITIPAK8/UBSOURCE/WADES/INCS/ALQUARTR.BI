TYPE ATransmitterType
  RecID      AS STRING * 1    'Must equal "A"
  RptPeriod  AS STRING * 4    'Must equal 0397,0697,0997,1297
  FedEIN     AS STRING * 9    '690640046 ??
  TranName   AS STRING * 35   'Brewton City Clerk LSET,FILLED SPACES
  Filler     AS STRING * 31   '
END TYPE

TYPE BBasAuthRecType
  RecID      AS STRING * 1    'Must equal "B"
  RptPeriod  AS STRING * 4    'Must equal 0397,0697,0997,1297
  FedEIN     AS STRING * 9    '690640046 ??
  Computer   AS STRING * 8    '"IBM"
  RecCode    AS STRING * 6    'Must be "ASCII"
  Filler     AS STRING * 52
END TYPE

TYPE EEmployerRecType
  RecID      AS STRING * 1    'Must equal "E"
  RptPeriod  AS STRING * 4    'Must equal 0397,0697,0997,1297
  FedEIN     AS STRING * 9    '690640046 ??
  Employer   AS STRING * 24   'Employer Name
  DIRAcctNo  AS STRING * 10   'DIR Acount No.
  NameCode   AS STRING * 1    'must be "F"
  StateCode  AS STRING * 2    '"01"
  UnitNum    AS STRING * 5    '"00000"
  Filler     AS STRING * 24
END TYPE

TYPE SEmpWageRecType
  RecID      AS STRING * 1    'Must equal "S"
  EmpSOC     AS STRING * 9    'Employee soc
  EmpName    AS STRING * 27   'Employee Name   LSET,FILLED SPACES
  StateCode  AS STRING * 2    '"01"
  RptPeriod  AS STRING * 4    'Must equal 0397,0697,0997,1297
  Wages      AS STRING * 9    'RIGHT JUSTIFY,ZERO FILL,NO NO NUMERIC
  DirNumber  AS STRING * 10   '"0213057600"
  Filler     AS STRING * 18
END TYPE

TYPE IInterMedRecType
  RecID      AS STRING * 1    'Must equal "I"
  TotWages   AS STRING * 10   'RIGHT JUSTIFY,ZERO FILL,NO NO NUMERIC
                            'I think it's the total Wages in the last segment
                            'of 41 or less recs.
  Filler     AS STRING * 69
END TYPE

TYPE TTotalRecType
  RecID      AS STRING * 1    'Must equal "T"
  TotalSRecs AS STRING * 7    'Total Count of SEmpWageRecs
  TotAllWage AS STRING * 13   'Total Wages of SEmpWageRecs
  Filler     AS STRING * 59   '
END TYPE

TYPE TFinalRecType
  RecID      AS STRING * 1    'Must equal "F"
  TotalSRecs AS STRING * 7    'Total Count of SEmpWageRecs
  Filler     AS STRING * 72   '
END TYPE

