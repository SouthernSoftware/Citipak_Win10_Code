DEFINT A-Z
DECLARE SUB OldUtilPayEntry (OperRecNumber%, PostDate$)
DECLARE SUB OpenSetupFile (SetUpFileNum%)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintJournal (OPERNUMB)
DECLARE SUB PrintReceipt ()
DECLARE SUB PostAndPrint (Posted)
DECLARE SUB InvalidCodeNumber ()
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
DECLARE SUB GetMiscCodeRecord (MCFile, RecNo, MiscCodeValue$)
DECLARE SUB AddMiscCode ()
DECLARE SUB MiscPayEntry (OperRecNumber, PostDate$)
DECLARE SUB UtilPayEntry (OperRecNumber, PostDate$)
DECLARE SUB SourceOfReceipt (OperRecNumber, source$)
DECLARE SUB AuthorizeOperator (oktoadd)
DECLARE SUB AddEditOperator ()
DECLARE SUB ValidateOperator (OperNum, OperPassword$, Results)
DECLARE SUB OperatorEntry (Results, OperNum, PostDate$)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE SUB FileView (FileName$, Ky, action, FVI AS ANY, SEG array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION Num2Date$ (Dat%)
'DECLARE FUNCTION Peek1% (Segment%, Address%)
'DECLARE FUNCTION WEnvTest ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
' DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()

  TYPE Struct
   who AS STRING * 14
   RecNum AS INTEGER
  END TYPE

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'

'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'

'*** Citipak Include Files Needed ***
'$INCLUDE: 'CMFILES.BI'
'$INCLUDE: 'ubtrans.BI'

  CONST False = 0, True = NOT False

  HideCursor

  OperatorEntry Results, OperNum, PostDate$

  IF Results < 0 THEN GOTO ExitProgram
  IF Results = 0 THEN RUN

  ' Next Find Out Source of Receipt 1=Misc : 2=Util : 3=Util Dep : 4=Tax : 5=Lic
  ' Only 1 & 2 Are Valid Now

  source = 0    ' Set to 0

  OperRecNumber = Results       ' Set Operator Record Number to be pulled when needed

  DO

    'SourceOfReceipt OperRecNumber, source$

    IF source$ = "INVALID" THEN RUN

    'IF source$ = "M" THEN
    '  MiscPayEntry OperRecNumber, PostDate$
    'END IF

   ' IF source$ = "U" THEN
   '   OldUtilPayEntry OperRecNumber, PostDate$
   ' END IF

  LOOP UNTIL source$ = "INVALID"


ExitProgram:
  'RUN "cpmenu"
  HideCursor
  CLS
  END

SUB oSourceOfReceipt (OperRecNumber, source$)
startmainsub:


  SHARED Choice$()
  
  LibName$ = "CM"
  ScrnName$ = "CMSOURCE"
  help$ = "Cash Management System V8.0"
  LOCATE 1, 1, 0
  
  
  ShowCursor
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  PrintHelp help$

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Rem Set Choice
  REDIM Choice$(2, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Utility Payment"
  Choice$(2, 0) = "Misc Payment"
  'Choice$(3, 0) = "Utility Deposit"
  'Choice$(4, 0) = "Tax Payment"
  'Choice$(5, 0) = "License Payment"
  
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET form$(F, 0) = ""
  NEXT F
  
  
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 3, 55: COLOR 15: PRINT LEFT$(CMOperRec(1).operatorname, 19)
  OPERNUMB = CMOperRec(1).OperatorNumber
  CLOSE CMFile
  
  DO
    EditForm form$(), Fld(), frm(1), Cnf, action
    
    source$ = LEFT$(form$(1, 0), 1)
    IF ASC(LEFT$(source$, 1)) > 32 THEN EXIT SUB
    
    SELECT CASE frm(1).KeyCode
    CASE F7KEY
      'PrintJournal (OperRecNumber)
      action = 1
      GOTO startmainsub
    CASE ESC
      source$ = "INVALID"
      Done = True
    CASE ELSE
      Done = False
    END SELECT

  LOOP UNTIL Done
  
END SUB

