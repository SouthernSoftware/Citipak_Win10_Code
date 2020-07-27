DEFINT A-Z
'Update Personal Prop Records from Txt file Stuart VA 2003
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TaxFile%)
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
   DIM a$(25)
   OpenTaxCustFile NUMOFTAXRECS, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "Stuart.mi" FOR INPUT AS #11

1 LINE INPUT #11, a$
  IF ERR = 62 THEN
    CLOSE
    PRINT
    PRINT "Done"
    END
  END IF

  PRINT a$
  didcnt = didcnt + 1

  k1 = INSTR(a$, "|")
   Oln$ = LEFT$(a$, k1 - 1)
  k2 = INSTR(k1 + 1, a$, "|")
   OFN$ = MID$(a$, k1 + 1, k2 - (k1 + 1))
  k3 = INSTR(k2 + 1, a$, "|")
   Omn$ = MID$(a$, k2 + 1, k3 - (k2 + 1))
  k4 = INSTR(k3 + 1, a$, "|")
   COLN$ = MID$(a$, k3 + 1, k4 - (k3 + 1))
  k5 = INSTR(k4 + 1, a$, "|")
   COFN$ = MID$(a$, k4 + 1, k5 - (k4 + 1))
  k6 = INSTR(k5 + 1, a$, "|")
   COMN$ = MID$(a$, k5 + 1, k6 - (k5 + 1))
  k7 = INSTR(k6 + 1, a$, "|")
   Add1$ = MID$(a$, k6 + 1, k7 - (k6 + 1))
  k8 = INSTR(k7 + 1, a$, "|")
   Add2$ = MID$(a$, k7 + 1, k8 - (k7 + 1))
  k9 = INSTR(k8 + 1, a$, "|")
   City$ = MID$(a$, k8 + 1, k9 - (k8 + 1))
  k10 = INSTR(k9 + 1, a$, "|")
   State$ = MID$(a$, k9 + 1, k10 - (k9 + 1))
  k11 = INSTR(k10 + 1, a$, "|")
   Zip$ = MID$(a$, k10 + 1, k11 - (k10 + 1))
  k12 = INSTR(k11 + 1, a$, "|")
   Desc$ = MID$(a$, k11 + 1, k12 - (k11 + 1))
  k13 = INSTR(k12 + 1, a$, "|")
   Dist$ = MID$(a$, k12 + 1, k13 - (k12 + 1))
  k14 = INSTR(k13 + 1, a$, "|")
   Class$ = MID$(a$, k13 + 1, k14 - (k13 + 1))
  k15 = INSTR(k14 + 1, a$, "|")
   Amount$ = MID$(a$, k14 + 1, k15 - (k14 + 1))
   Amount# = VAL(Amount$)

   FirstName$ = RTRIM$(OFN$) + " " + RTRIM$(Omn$)
   IF LEN(FirstName$) > 15 THEN FirstName$ = LEFT$(FirstName$, 15)
   LastName$ = RTRIM$(Oln$)
   IF LEN(LastName$) > 25 THEN LastName$ = LEFT$(LastName$, 25)


   CustName$ = LTRIM$(FirstName$ + " " + LastName$)


   T = T + 1

   FOR Cnt& = 1 TO NUMOFTAXRECS
     GET TaxFile, Cnt&, TaxCustRec(1)
     TaxCustName$ = RTRIM$(TaxCustRec(1).FNAME) + " " + RTRIM$(TaxCustRec(1).LName)
     TaxCustName$ = LTRIM$(TaxCustName$)

     NF = 1
    IF TaxCustName$ = CustName$ THEN
      GOSUB UpdatePersonal
      GOTO 1
    END IF
   NEXT Cnt&

  ' Add Customer
    Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
    TaxCustRec(1).FNAME = RTRIM$(OFN$) + " " + RTRIM$(Omn$)
    TaxCustRec(1).LName = Oln$
    TaxCustRec(1).SName = Oln$
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = Add1$
    TaxCustRec(1).Addr2 = Add2$
    TaxCustRec(1).City = City$
    TaxCustRec(1).State = State$
    TaxCustRec(1).Zip = Zip$
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = Acct!
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("10-10-2004")
    PUT TaxFile, Record!, TaxCustRec(1)
    Cnt& = Record!
    GOSUB UpdatePersonal
    GOTO 1



UpdatePersonal:
    
  Class = VAL(Class$)
  IF Class = 4 THEN
    MH# = MH# + Amount#
  END IF
  IF Class = 3 THEN
    MT# = MT# + Amount#
  END IF



   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF Class = 3 THEN
    PersRec(1).MTVALUE = Amount#
    GOTO upd1
   END IF
   IF Class = 4 THEN
    PersRec(1).MHVALUE = Amount#
    GOTO upd1
   END IF

   PersRec(1).PERSVAL = Amount#
  
upd1:
   p# = p# + Amount#
   PersRec(1).PROPPIN = LTRIM$(STR$(1))
   PersRec(1).PROPDATE = Date2Num%("09-25-2004")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = Desc$
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = "N" + STRING$(25, " ") + "2004"
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2003
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN


END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

