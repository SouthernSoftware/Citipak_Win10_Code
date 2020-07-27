DEFINT A-Z
'Convert Owner Records to Master PP Records from TXT File (Independence 2001)
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
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


  'New Market Va Tax PP Conversion of Master (Customers)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'TaxConv.bi'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED taxcustrec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED IOwn(1) AS IOwner

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(55)
   OPEN "indepo~1.txt" FOR INPUT AS #11
   CLS
   b = 0
1  FOR cc = 1 TO 15
    INPUT #11, a$(cc)
    PRINT a$(cc), cc
   NEXT cc
   PRINT STRING$(79, "-")
   'SLEEP 1
   
   TaxFile = FREEFILE
   OPEN "PPOWNER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = 8
    IOwn(1).Owner = VAL(a$(1))
    IOwn(1).NameID = VAL(a$(2))
   PUT TaxFile, LOF(TaxFile) / 8 + 1, IOwn(1)
   CLOSE TaxFile
   GOTO 1
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(taxcustrec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(taxcustrec(1))
  
END SUB

