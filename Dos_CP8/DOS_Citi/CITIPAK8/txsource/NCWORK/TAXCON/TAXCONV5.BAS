' White Lake 2001 Master Names (WLNAMES) followed by wlbills
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
  DIM A$(120)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  CLS
   OPEN "WLNAMES" FOR INPUT AS #11
   Record = NumOfTaxRecs
1  LINE INPUT #11, A$
    PRINT A$
   A$(1) = LEFT$(A$, 7)
   A$(2) = MID$(A$, 9, 40)
   A$(3) = MID$(A$, 45, 25)
   A$(4) = RTRIM$(MID$(A$, 76, 20))
   A$(5) = MID$(A$, 97, 5)
   A$(6) = MID$(A$, 103, 9)
   A$(7) = MID$(A$, 113, 7)
   A$(8) = MID$(A$, 121, 9)
   A$(9) = MID$(A$, 131, 15)

   
   PRINT STRING$(50, "-")


   'Find Acct
   FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(A$(1)) THEN


   IF VAL(A$(5)) > 0 THEN
    LNM$ = RTRIM$(A$(2))
    ELSE
   NM$ = RTRIM$(A$(9))
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

   END IF
     TaxCustRec(1).OPENDATE = Date2Num%("07-01-2002")

     TaxCustRec(1).FName = FRNM$

     TaxCustRec(1).LNAME = LNM$
     TaxCustRec(1).SName = LNM$
     TaxCustRec(1).Addr1 = A$(3)
     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = LEFT$(A$(4), LEN(A$(4)) - 2)
     TaxCustRec(1).State = RIGHT$(RTRIM$(A$(4)), 2)
     TaxCustRec(1).Zip = LEFT$(A$(8), 5)
     TaxCustRec(1).HPHONE = "   " + A$(7)
     TaxCustRec(1).CSSN = A$(6)
     TaxCustRec(1).Deleted = 0
     TaxCustRec(1).Active = "Y"
     PUT TaxFile, Cnt&, TaxCustRec(1)
     GOTO 1
    END IF
  NEXT Cnt&

'Add Record
   'Split Name Here
   Record = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1

   NM$ = RTRIM$(A$(9))
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
   TaxCustRec(1).LNAME = LNM$
   TaxCustRec(1).SName = A$(1)
   TaxCustRec(1).Addr1 = A$(3)
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).City = A$(4)
   TaxCustRec(1).State = RIGHT$(RTRIM$(A$(4)), 2)
   TaxCustRec(1).Zip = LEFT$(A$(8), 5)
   TaxCustRec(1).HPHONE = A$(7)
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = A$(6)
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcct = VAL(A$(1))
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   'GOSUB UpdateReal
   'GOSUB UpdatePers
   GOTO 1
   
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = A$(22)
    PropertyRec(1).PROPDATE = -32767
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = A$(23)
    PropertyRec(1).BLOCK = A$(25)
    PropertyRec(1).LOTNUMB = A$(26)
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(A$(11))
    PropertyRec(1).EXMPSENI = VAL(A$(14))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = A$(28)
    PropertyRec(1).PROPNOT2 = A$(29)
    PropertyRec(1).PROPNOT3 = A$(30)
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN
UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = -32767
   PersRec(1).PersVal = VAL(A$(12))
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = VAL(A$(15))
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = ""
   PersRec(1).Desc5 = ""
   PersRec(1).CustPin = Record
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN

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

