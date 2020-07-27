DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB ARFixMess
  
  CLS
  
  REDIM ARTransRec(1) AS ARTransRecType
  
  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))
  
  
  FOR Cnt = 1 TO 104
    GET ARTRFile, Cnt, ARTransRec(1)
    
    'Get Customer
    OpenARCustFile NumOfArRecs, ARFile
    GET ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
    ARCustRec(1).IssueLicense = "Y"
    ARCustRec(1).FeeAmt = ARTransRec(1).TransAmount
    PUT ARFile, VAL(ARTransRec(1).CustomerNumber), ARCustRec(1)
    
    
    IF ARTransRec(1).TransType = 1 THEN PRINT "Charge"; Cnt: STOP
    IF ARTransRec(1).TransType = 2 THEN PRINT "Payment"; Cnt
    
  NEXT Cnt
  CLOSE         'Close all open files now
  
  
END SUB

