  'apaux.bas
  'DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, ky$, Action%, Cnf AS ANY)
  DEFINT A-Z
DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTNo, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
  
  '$INCLUDE: 'DefCnf.bi'
  '$INCLUDE: 'ap.bi'
  '$INCLUDE: 'formedit.BI'                    'Report printer routines
  '$INCLUDE: 'fieldinf.BI'                    'Report printer routines
  ''$INCLUDE: 'RPTPRN.BI'                    'Report printer routines
  '$INCLUDE: 'SCRAUX.BI'
  
  'TYPE FLen
  '   V AS STRING * 64
  'END TYPE
  
  'DECLARE SUB VertMenuT (Item() AS FLen, Choice%, MaxLen%, BoxBot%, ky$, Action%, Cnf AS ANY)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE FUNCTION FindVendorRec (VendorCode$)
DECLARE FUNCTION VendorPickList (Str2Find$, Abort%)
DECLARE SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
DECLARE SUB OpenVendorFile (VendorFile, NumVRecs)
DECLARE SUB OpenDefDistFile (DefRecLen, DefDistFileNum, NumDefRecs)
  
  DIM SHARED VendorIdx AS VendorIdxRecType
  DIM SHARED Vendor AS VendorRecType
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'SetCnf.bi'

SUB DisplayAPScrn (ScrnName$)
  LibFile2Scrn "AP.QSL", ScrnName$, MonoCode, -1, ErrorCode
END SUB

SUB IndexVendorFile
  
  SHARED VendorIdx AS VendorIdxRecType
  
  IF Exist("apvendor.idx") THEN 'Delete index file if it exits
    KILL "apvendor.idx"
  END IF
  
  '  '--Open the Index file
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  
  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file
  
  IF LOF(VendorFile) = 0 THEN   'get out if nothing to do
    CLOSE
    EXIT SUB
  END IF
  
  REDIM IdxBuff(1 TO NumVRecs) AS VendorIdxRecType              'dim the buffer array
  
  FOR Cnt = 1 TO NumVRecs       'Load the buffer
    GET VendorFile, Cnt, Vendor
    IF Vendor.VIN > 99999 THEN
      Vendor.DELFLAG = -1
      PUT VendorFile, Cnt, Vendor
    END IF

    VendorName$ = QPTrim$(Vendor.VName)
    IF LEN(VendorName$) = 0 THEN
      Vendor.DELFLAG = -1
      PUT VendorFile, Cnt, Vendor
    END IF

    IF LEN(VendorName$) <> 0 THEN
      IF ASC(LEFT$(VendorName$, 1)) < 48 OR ASC(LEFT$(VendorName$, 1)) > 90 THEN
        Vendor.DELFLAG = -1
        PUT VendorFile, Cnt, Vendor
      END IF
    END IF
    
    VendorNumber$ = RTRIM$(Vendor.VNum)
    IF LEN(VendorNumber$) = 0 THEN
      Vendor.DELFLAG = -1
      PUT VendorFile, Cnt, Vendor
    END IF

    IF LEN(VendorNumber$) <> 0 THEN
      IF ASC(LEFT$(VendorNumber$, 1)) < 48 OR ASC(LEFT$(VendorNumber$, 1)) > 90 THEN
        Vendor.DELFLAG = -1
        PUT VendorFile, Cnt, Vendor
      END IF
    END IF
    
    IF Vendor.DELFLAG = 0 THEN  'Get only active records
      GoodAccts = GoodAccts + 1
      LSET IdxBuff(GoodAccts).VendorCode = QPTrim$(Vendor.VNum)
      'RSET IdxBuff(GoodAccts).VendorCode = QPTrim$(Vendor.VNum)
      IdxBuff(GoodAccts).RecNum = Cnt
    END IF
  NEXT Cnt
  
  CLOSE VendorFile              'Close the file
  
  '--redim with just good accts
  REDIM PRESERVE IdxBuff(1 TO GoodAccts) AS VendorIdxRecType
  
  SortT IdxBuff(1), GoodAccts, 0, 12, 0, 10
  
  'FPutAH "apvendor.idx", IdxBuff(1), 12, GoodAccts
  
  '--write out to the index file
  FOR Cnt = 1 TO GoodAccts
    GET VendorIdxFile, Cnt, AcctIdx
    'RSET VendorIdx.VendorCode = QPTrim$(IdxBuff(Cnt).VendorCode)
    LSET VendorIdx.VendorCode = QPTrim$(IdxBuff(Cnt).VendorCode)
    VendorIdx.RecNum = IdxBuff(Cnt).RecNum
    PUT VendorIdxFile, Cnt, VendorIdx
  NEXT
  '
  CLOSE VendorIdxFile           'close index
  
END SUB

SUB OpenAPDistFile (APDistFile, NumDistRecs&, RecLen)
  
  APDistFile = FREEFILE
  OPEN "APDIST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APDistFile LEN = RecLen
  NumDistRecs& = LOF(APDistFile) \ RecLen
  
END SUB

SUB OpenAPLedgerFile (APLedgerFile, NumTrans&, RecLen)
  
  APLedgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APLedgerFile LEN = RecLen
  NumTrans& = LOF(APLedgerFile) \ RecLen
  
END SUB

SUB OpenDefDistFile (DefRecLen, APDefDistFile, NumDefRecs)
  
  APDefDistFile = FREEFILE
  OPEN "APVDIST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APDefDistFile LEN = DefRecLen
  NumDefRecs = LOF(APDefDistFile) \ DefRecLen
  
END SUB

SUB OpenVendorFile (VendorFile, NumVRecs)
  
  'SHARED VENDOR AS VendorRecType
  VRecLen = LEN(Vendor)
  VendorFile = FREEFILE
  OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile LEN = VRecLen
  'VendorFile1 = FREEFILE
  'OPEN "apvend1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile1 LEN = VRecLen
  NumVRecs = LOF(VendorFile) \ VRecLen
  
  'FOR Cnt = 1 TO NumVRecs
  'GET VendorFile, Cnt, Vendor
  'IF Vendor.DELFLAG = 0 THEN
  'IF Cnt < 8226 THEN
  ' PUT VendorFile1, Cnt, Vendor
  'END IF
  'IF Cnt > 8226 THEN
  ' IF Cnt = 8226 THEN PUT VendorFile1, 632, Vendor
  ' IF Cnt = 8227 THEN PUT VendorFile1, 633, Vendor
  ' IF Cnt = 8228 THEN PUT VendorFile1, 634, Vendor
  ' IF Cnt = 8229 THEN PUT VendorFile1, 635, Vendor
  'END IF
  '
  '
  'END IF
  '
  '
  'NEXT Cnt
  'STOP
  
  
  
  
  
  
END SUB

SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
  
  SHARED VendorIdx AS VendorIdxRecType
  
  'OPEN "apvendor.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 14
  
  VendorIdxLen = LEN(VendorIdx)
  VendorIdxFile = FREEFILE
  OPEN "apvendor.idx" FOR RANDOM ACCESS READ WRITE SHARED AS VendorIdxFile LEN = VendorIdxLen
  NumActiveVendors = LOF(VendorIdxFile) \ VendorIdxLen
  
END SUB

