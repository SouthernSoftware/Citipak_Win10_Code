' White Lake 2003 Master Names (WLMSTR.TXT) followed by WLLBILL for Discovery's
DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
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
DECLARE FUNCTION Round# (B#)
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

'White Lake NC Version
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  CLS
   OPEN "wlmstr.dat" FOR INPUT AS #11
   Record = NumOfTaxRecs
1  LINE INPUT #11, a$
    PRINT a$
   a$(1) = LEFT$(a$, 7)
   a$(2) = MID$(a$, 8, 35)
   a$(3) = MID$(a$, 43, 30)
   a$(4) = RTRIM$(MID$(a$, 73, 20))
   a$(5) = RTRIM$(MID$(a$, 93, 5))
   a$(6) = MID$(a$, 104, 15)

   



' search for existing account first
 FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
 GET TaxFile, Cnt&, TaxCustRec(1)
 IF TaxCustRec(1).CountyAcct = VAL(a$(1)) THEN
 'update name and address and return
   NM$ = RTRIM$(a$(6))
   FRNM$ = ""
   NLegth = LEN(NM$)
   kk = INSTR(NM$, " ")
    IF kk > 0 THEN
     FRNM$ = MID$(NM$, kk + 1, NLegth - kk)
     LNM$ = LEFT$(NM$, kk - 1)
     ELSE
     LNM$ = NM$
     FRNM$ = ""
    END IF
   TaxCustRec(1).FName = FRNM$
   TaxCustRec(1).LName = LNM$
   TaxCustRec(1).SName = a$(6)
   TaxCustRec(1).Addr1 = a$(3)
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).City = LEFT$(a$(4), LEN(a$(4)) - 2)
   TaxCustRec(1).State = RIGHT$(RTRIM$(a$(4)), 2)
   TaxCustRec(1).Zip = LEFT$(a$(5), 5)
   PUT TaxFile, Cnt&, TaxCustRec(1)

   GOTO 1
 END IF


 NEXT Cnt&
 'Must Add Account


   PRINT STRING$(50, "-")
   'Split Name Here
   Record = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1

   NM$ = RTRIM$(a$(6))
   FRNM$ = ""
   NLegth = LEN(NM$)
   kk = INSTR(NM$, " ")
    IF kk > 0 THEN
     FRNM$ = MID$(NM$, kk + 1, NLegth - kk)
     LNM$ = LEFT$(NM$, kk - 1)
     ELSE
     LNM$ = NM$
     FRNM$ = ""
    END IF
   TaxCustRec(1).Acct = Record
   TaxCustRec(1).FName = FRNM$
   TaxCustRec(1).LName = LNM$
   TaxCustRec(1).SName = a$(6)
   TaxCustRec(1).Addr1 = a$(3)
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).City = LEFT$(a$(4), LEN(a$(4)) - 2)
   TaxCustRec(1).State = RIGHT$(RTRIM$(a$(4)), 2)
   TaxCustRec(1).Zip = LEFT$(a$(5), 5)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcct = VAL(a$(1))
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   GOTO 1
   
   CLOSE
   EXIT SUB

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

