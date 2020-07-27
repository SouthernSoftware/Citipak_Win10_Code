DEFINT A-Z
'Convert MASTER PP Records from WARSAW.TXT File YEAR=2003
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


  'Stephens City Va Tax PP Conversion of Master (Customers)

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
  DIM a$(38)
  OpenTaxCustFile NumOfTaxRecs, TaxFile


   OPEN "TAXPAYer.TXT" FOR INPUT AS #11
   CLS
   b = 0
   
1  c = c + 1
   
  
FOR cc = 1 TO 14

    INPUT #11, a$(cc)
   PRINT cc, a$(cc)
   NEXT cc
   PRINT STRING$(79, "-")

   
   acct$ = a$(1)
   acct! = VAL(acct$)
   PRINT acct!, a$(2)

   FOR Cnt& = 1 TO NumOfTaxRecs
   GET TaxFile, Cnt&, TaxCustRec(1)

   IF acct! = TaxCustRec(1).CountyAcct THEN
    Record! = Cnt&
    GOSUB UpdateAcct
    GOTO 1
   END IF
   NEXT Cnt&

   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
   IF a$(14) = "I" THEN
      TaxCustRec(1).FNAME = a$(4)
      TaxCustRec(1).LName = a$(3)
      TaxCustRec(1).SName = a$(3)
     ELSE
      TaxCustRec(1).FNAME = ""
      TaxCustRec(1).LName = a$(3)
      TaxCustRec(1).SName = a$(3)
    END IF

    IF a$(14) = "P" OR a$(2) = "B" THEN
     TaxCustRec(1).FNAME = ""
     TaxCustRec(1).LName = a$(3)
     TaxCustRec(1).SName = a$(3)
    END IF

    TaxCustRec(1).acct = Record!
    IF LEN(RTRIM$(a$(5))) = 0 THEN
    TaxCustRec(1).Addr1 = a$(6)
    TaxCustRec(1).Addr2 = a$(7)
    ELSE
    TaxCustRec(1).Addr1 = a$(5)
    TaxCustRec(1).Addr2 = a$(6)
    END IF
    TaxCustRec(1).City = a$(8)
    TaxCustRec(1).STATE = a$(9)
    TaxCustRec(1).Zip = LEFT$(a$(10), 5)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = a$(11)
    TaxCustRec(1).SSSN = a$(12)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(a$(1))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("11-15-2003")
    PUT TaxFile, Record!, TaxCustRec(1)

   
   GOTO 1

UpdateAcct:
   IF a$(14) = "I" THEN
      TaxCustRec(1).FNAME = a$(4)
      TaxCustRec(1).LName = a$(3)
      TaxCustRec(1).SName = a$(3)
     ELSE
      TaxCustRec(1).FNAME = ""
      TaxCustRec(1).LName = a$(3)
      TaxCustRec(1).SName = a$(3)
    END IF

    IF a$(14) = "P" OR a$(2) = "B" THEN
     TaxCustRec(1).FNAME = ""
     TaxCustRec(1).LName = a$(3)
     TaxCustRec(1).SName = a$(3)
    END IF
    IF LEN(RTRIM$(a$(5))) = 0 THEN
    TaxCustRec(1).Addr1 = a$(6)
    TaxCustRec(1).Addr2 = a$(7)
    ELSE
    TaxCustRec(1).Addr1 = a$(5)
    TaxCustRec(1).Addr2 = a$(6)
    END IF
    TaxCustRec(1).City = a$(8)
    TaxCustRec(1).STATE = a$(9)
    TaxCustRec(1).Zip = LEFT$(a$(10), 5)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = a$(11)
    TaxCustRec(1).SSSN = a$(12)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).OPENDATE = Date2Num%("11-15-2003")
    PUT TaxFile, Record!, TaxCustRec(1)
    RETURN

   

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

