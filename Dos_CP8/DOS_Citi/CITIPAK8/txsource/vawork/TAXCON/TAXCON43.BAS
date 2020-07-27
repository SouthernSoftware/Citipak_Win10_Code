DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  New Market Va  2002
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
  '$INCLUDE: 'TaxConv.bi'
  
  DIM SHARED zip(1) AS ZipType
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(87)

  OPEN "ZipCode.txt" FOR INPUT AS #11
1 INPUT #11, a$
  PRINT a$
  a$ = RTRIM$(a$)

  Zipc$ = LEFT$(a$, 5)
  kk = INSTR(a$, Zipc$)
  IF LEN(a$) > 5 THEN
    j = LEN(a$) - (kk + 10)
    City$ = MID$(a$, 6, j)
   ELSE
    City$ = ""
  END IF
  OPEN "ZipCode1.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #10 LEN = LEN(zip(1))
   zip(1).ZipCode = Zipc$
   zip(1).City = City$
   NexRecord! = (LOF(10)) \ LEN(zip(1)) + 1
  PUT 10, NexRecord!, zip(1)
  CLOSE 10
  GOTO 1


Nextone:
   GOTO 1
   CLOSE
   EXIT SUB



END SUB

