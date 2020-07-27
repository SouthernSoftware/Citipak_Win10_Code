DEFINT A-Z
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
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBCustType
  DIM SHARED TBCustRec1(1) AS TBCust1Type
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "STPCTYPM.TXT" FOR INPUT AS #11
   CLS
   b = 0
1  LINE INPUT #11, a$
   PRINT a$
   b = b + 1
   PRINT STRING$(79, "-")


    TaxFile = FREEFILE
    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    Record! = NumOfTaxRecs + 1


    NM$ = MID$(a$, 18, 35)
    kk = INSTR(NM$, ",")
    IF kk = 0 THEN
     FstName$ = ""
     LName$ = NM$
      ELSE
     LName$ = LTRIM$(LEFT$(NM$, kk - 1))
     FstName$ = LTRIM$(MID$(NM$, kk + 1, LEN(NM$) - (kk)))
    END IF

    
    PRINT "Add 1 "; MID$(a$, 53, 30)
    PRINT "Add 2 "; MID$(a$, 83, 30)
    PRINT "Add 3 "; MID$(a$, 113, 30)
    
    TaxCustRec(1).FNAME = UCASE$(LTRIM$(FstName$))
    TaxCustRec(1).LName = UCASE$(LName$)
    TaxCustRec(1).SName = UCASE$(LName$)
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).CountyAcct = 0


    IF LEN(QPTrim$(MID$(a$, 113, 30))) = 0 THEN
     TaxCustRec(1).Addr1 = UCASE$(MID$(a$, 54, 30))
     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = UCASE$(MID$(a$, 84, 30))
     TaxCustRec(1).STATE = ""
    ELSE
     TaxCustRec(1).Addr1 = UCASE$(MID$(a$, 54, 30))
     TaxCustRec(1).Addr2 = UCASE$(MID$(a$, 84, 30))
     TaxCustRec(1).City = UCASE$(MID$(a$, 114, 30))
     TaxCustRec(1).STATE = ""
    END IF
    TaxCustRec(1).Zip = MID$(a$, 145, 5)
    TaxCustRec(1).HPHONE = MID$(a$, 277, 10)
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = MID$(a$, 174, 9)
    TaxCustRec(1).SSSN = MID$(a$, 184, 9)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(MID$(a$, 6, 8))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = x
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("08-01-1999")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
    CLOSE TaxFile
   
NextOne:
   GOTO 1
   CLOSE
   EXIT SUB


Updatepersonal:
   PersRec(1).PERSVAL = TBCustRec(1).op
   PersRec(1).MCVALUE = TBCustRec(1).fm
   PersRec(1).CVALUE = TBCustRec(1).ls
   PersRec(1).MTVALUE = TBCustRec(1).ME
   PersRec(1).MHVALUE = TBCustRec(1).mh
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("03-01-1999")
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = TBCustRec(1).pdesc
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
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

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

