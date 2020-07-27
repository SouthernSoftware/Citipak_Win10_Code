DEFINT A-Z

'$INCLUDE: 'ap.bi'
'$INCLUDE: 'cp65ap.bi'

DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION GetVendorRec (VendorCode$)
DECLARE SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
DECLARE SUB OpenVendorFile (VendorFile, NumVRecs)
DECLARE SUB OpenAPLedgerFile (APLedgerFile, NumTrans&, RecLen)

DIM SHARED VendorIdx AS VendorIdxRecType
DIM SHARED VENDOR AS VendorRecType

CONST False = 0, True = NOT False

FUNCTION GetVendorRec (VendorCode$)
  
  SHARED VendorIdx AS VendorIdxRecType
  
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  
  IF NumActiveVendors = 0 THEN
    FindVendorRec = 0
    CLOSE VendorIdxFile
    EXIT FUNCTION
  END IF
  
  Match = False
  FirstRec = 1
  LastRec = NumActiveVendors
  Lookfor$ = QPTrim$(UCASE$(VendorCode$))
  
  DO UNTIL LastRec < FirstRec
    
    MiddleRec = (LastRec + FirstRec) \ 2
    
    GET VendorIdxFile, MiddleRec, VendorIdx
    
    IF Lookfor$ = QPTrim$(VendorIdx.VendorCode) THEN
      Match = True
      EXIT DO
    ELSEIF Lookfor$ < VendorIdx.VendorCode THEN
      LastRec = MiddleRec - 1
    ELSE
      FirstRec = MiddleRec + 1
    END IF
    
  LOOP
  
  IF Match THEN
    GetVendorRec = VendorIdx.RecNum
  ELSE
    GetVendorRec = 0
  END IF
  
  CLOSE VendorIdxFile
  
END FUNCTION

SUB OpenAP7VIdx (Vendor7IdxFile, NumActive7Vendors, VIdx AS Vendor7IdxRecType)

   Vendor7IdxFile = FREEFILE
   VIdxRecLen = LEN(VIdx)
   OPEN "apvenddx.dat" FOR RANDOM SHARED AS Vendor7IdxFile LEN = VIdxRecLen
   NumActive7Vendors = LOF(Vendor7IdxFile) \ VIdxRecLen

END SUB

SUB OpenAPDistFile (APDistFile, NumDistRecs&, RecLen)
  
  APDistFile = FREEFILE
  OPEN "APDIST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APDistFile LEN = RecLen
  NumDistRecs& = LOF(APDistFile) \ RecLen
  
END SUB

SUB OpenAPLedgerFile (APLedgerFile, NumTran&, RecLen)
  
  APLedgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APLedgerFile LEN = RecLen

  NumTran& = LOF(APLedgerFile) \ RecLen
  
END SUB

SUB OpenDefDistFile

   'APDefDistFile = FREEFILE
   'OPEN "apinvdis.dat" FOR RANDOM AS APDefDistFile

END SUB

SUB OpenVendorFile (VendorFile, NumVRecs)

   SHARED VENDOR AS VendorRecType
   VRecLen = LEN(VENDOR)
   VendorFile = FREEFILE
   OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile LEN = VRecLen
   NumVRecs = LOF(VendorFile) \ VRecLen
   
END SUB

SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
   
   SHARED VendorIdx AS VendorIdxRecType
   VendorIdxLen = LEN(VendorIdx)
   VendorIdxFile = FREEFILE
   OPEN "apvendor.idx" FOR RANDOM ACCESS READ WRITE SHARED AS VendorIdxFile LEN = VendorIdxLen
   NumActiveVendors = LOF(VendorIdxFile) \ VendorIdxLen
   
END SUB

