DEFINT A-Z
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB OpenPOTransFile (TransFileNum%, NumTrans&)
'apaux.bas
'DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, ky$, Action%, Cnf AS ANY)
DECLARE FUNCTION FindFund% (FundNum$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, ky, Action, FVI AS ANY, SEG Array)
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
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE FUNCTION Round# (N#)

'$INCLUDE: 'DefCnf.bi'
'$INCLUDE: 'ap.bi'
'$INCLUDE: 'gl.bi'
'$INCLUDE: 'formedit.BI'                    'Report printer routines
'$INCLUDE: 'fieldinf.BI'                    'Report printer routines
'$INCLUDE: 'SCRAUX.BI'

'TYPE FLen
'   V AS STRING * 64
'END TYPE

'DECLARE SUB VertMenuT (Item() AS FLen, Choice%, MaxLen%, BoxBot%, ky$, Action%, Cnf AS ANY)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, ky$, Action%, Cnf AS ANY)
DECLARE FUNCTION FindVendorRec (VendorCode$)
DECLARE FUNCTION VendorPickList (Str2Find$, Abort%)
DECLARE SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
DECLARE SUB OpenVendorFile (VendorFile, NumVRecs)

DIM SHARED VendorIdx AS VendorIdxRecType
DIM SHARED VENDOR AS VendorRecType

CONST False = 0, True = NOT False

'$INCLUDE: 'SetCnf.bi'

SUB OpenAPDistFile (APDistFile, NumDistRecs&, RecLen)
  
  APDistFile = FREEFILE
  OPEN "APDIST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APDistFile LEN = RecLen
  NumDistRecs& = LOF(APDistFile) \ RecLen
  
END SUB

SUB OpenAPLedgerFile (APLedgerFile, NumTrans, RecLen)
  
  APLedgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APLedgerFile LEN = RecLen
  NumTrans = LOF(APLedgerFile) \ RecLen
  
END SUB

SUB OpenDefDistFile

   'APDefDistFile = FREEFILE
   'OPEN "apinvdis.dat" FOR RANDOM AS APDefDistFile

END SUB

SUB OpenVendorFile (VendorFile, NumVRecs)

'   SHARED VENDOR AS VendorRecType
   VRecLen = LEN(VENDOR)
   VendorFile = FREEFILE
   OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile LEN = VRecLen
   NumVRecs = LOF(VendorFile) \ VRecLen
   
END SUB

SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)
   
   SHARED VendorIdx AS VendorIdxRecType
   
   'OPEN "apvendor.idx" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 14

   VendorIdxLen = LEN(VendorIdx)
   VendorIdxFile = FREEFILE
   OPEN "apvendor.idx" FOR RANDOM ACCESS READ WRITE SHARED AS VendorIdxFile LEN = VendorIdxLen
   NumActiveVendors = LOF(VendorIdxFile) \ VendorIdxLen
   
END SUB

