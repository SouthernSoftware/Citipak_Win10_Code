  DEFINT A-Z
  'Update Personal Prop Detail Records from Txt file WARSAW 2003 LINEITEM.TXT
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TAXFILE%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB CustomerListing ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB ClearScrn ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (b#)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
  CONST False = 0, True = NOT False
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  CLS
  ON ERROR RESUME NEXT
  DIM A$(30)
  CALL KillFile("taxpers.dat")
  OpenTaxCustFile NUMOFTAXRECS, TAXFILE
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  CLS
  OPEN "LINEITEm.TXT" FOR INPUT AS #11

' Initialize Existing Properties
'  FOR Cnt& = 1 TO NumOfPersRecs
'    GET PersTaxFile, Cnt&, PersRec(1)
'    PersRec(1).MHVALUE = 0
'    PersRec(1).MCVALUE = 0
'    PersRec(1).CVALUE = 0
'    PersRec(1).MTVALUE = 0
'    PersRec(1).EXMPSENI = 0
'    PersRec(1).EXMPOTHR = 0
'    IF LEFT$(PersRec(1).Desc5, 1) = "N" THEN
'      PersRec(1).PersVal = 0
'    END IF
'    PUT PersTaxFile, Cnt&, PersRec(1)
'  NEXT Cnt&

Top:
  FOR CC = 1 TO 30
    INPUT #11, A$(CC)
    IF ERR = 62 THEN
      CLOSE
      PRINT "Done"
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Boom! Error Code:"; ERR

      END
    END IF
    PRINT A$(CC)
  NEXT CC
  PRINT A$(3)
  PRINT STRING$(79, "-")
  LineCnt = LineCnt + 1

  'IF LEN(A$(1)) <> 1 THEN STOP
  IF SkipFirst = 0 THEN
    SkipFirst = 1
    GOTO Top
  END IF
  
  FOR Cnt& = 1 TO NUMOFTAXRECS
    GET TAXFILE, Cnt&, TaxCustRec(1)
    NF = 1
    IF TaxCustRec(1).CountyAcct = VAL(A$(3)) THEN  'was 2
      GOSUB updatepersonal
      NF = 0
      EXIT FOR
    END IF
  NEXT Cnt&
  IF NF = 1 THEN
    STOP
    NoFind = NoFind + 1
  END IF
Nextone:
  GOTO Top
  CLOSE
  EXIT SUB
  
  
updatepersonal:
  
  VTYPE = 0
  
  IF VAL(A$(23)) > 0 THEN VTYPE = 13
  IF VAL(A$(25)) > 0 THEN VTYPE = 14
  IF VAL(A$(26)) > 0 THEN VTYPE = 15
  
  PersRec(1).MCVALUE = 0
  PersRec(1).MTVALUE = 0
  PersRec(1).MHVALUE = 0
  PersRec(1).PersVal = 0
  
  IF VTYPE = 14 THEN
    PersRec(1).MTVALUE = VAL(A$(25))
    GOTO upd1
  END IF
  IF VTYPE = 15 THEN
    PersRec(1).MHVALUE = VAL(A$(26))
    GOTO upd1
  END IF
  
  PersRec(1).PersVal = VAL(A$(23))

upd1:
  IF A$(4) = "Y" THEN QUALIFY$ = "Y/" ELSE QUALIFY$ = "N/"   'was 3
  PersRec(1).PROPPIN = "1"
  PersRec(1).PROPDATE = Date2Num%("11-09-2005")
  PersRec(1).CVALUE = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = A$(6)  'was 5
  PersRec(1).DESC2 = A$(5)  'was 4
  PersRec(1).DESC3 = LEFT$(A$(5), 4)   'was 4
  PersRec(1).DESC4 = ""
  PersRec(1).Desc5 = QUALIFY$ + STRING$(24, " ") + "2005"
  PersRec(1).CustPin = Cnt&
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 2004
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
  PUT PersTaxFile, PRecord&, PersRec(1)
  RETURN
  
  
  
END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

