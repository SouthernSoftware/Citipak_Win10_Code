DEFINT A-Z
  'Shenandoah Va Tax  Update of Master (Customers) 2001
'Convert Master PP Records from TXT File
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




  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(35)
  OpenTaxCustFile NumOfTaxRecs, TaxFile

   OPEN "PPMASTER.DAT" FOR INPUT AS #11
   CLS
   b = 0
1  FOR cc = 1 TO 26
    INPUT #11, a$(cc)
   NEXT cc
   PRINT a$(1), a$(2)
                  
    TaxFile = FREEFILE
    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    Record! = NumOfTaxRecs + 1
    NME$ = a$(6)
    kk = INSTR(NME$, " ")
    IF a$(5) = "I" THEN
     L$ = LEFT$(NME$, kk - 1)
     F$ = MID$(NME$, kk + 1, LEN(NME$) - kk)
     S$ = L$
    ELSE
     L$ = NME$
     F$ = ""
     S$ = L$
    END IF
    IF LEN(LTRIM$(a$(10))) = 4 THEN a$(10) = "0" + a$(10)
    
    TaxCustRec(1).FNAME = F$
    TaxCustRec(1).LName = L$
    TaxCustRec(1).SName = S$
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = a$(8)
    TaxCustRec(1).Addr2 = a$(9)
    TaxCustRec(1).City = ""
    TaxCustRec(1).STATE = ""
    TaxCustRec(1).Zip = a$(10)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = a$(3)
    TaxCustRec(1).SSSN = a$(4)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(a$(2))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = x
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("11-01-2001")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
    CLOSE TaxFile
   
NextOne:
   GOTO 1
   CLOSE
   EXIT SUB

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

