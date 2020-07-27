DEFINT A-Z
'Convert Data From County (pinebluff)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumber%)

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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "pbluff.txt" FOR INPUT AS #10
   rECORD = 0
    CLS
10 LINE INPUT #10, a$
    rECORD = rECORD + 1
    CtyAcct$ = LEFT$(a$, 9)
    NM$ = MID$(a$, 19, 30)
    Add1$ = MID$(a$, 49, 30)
    Add2$ = MID$(a$, 79, 30)
    City$ = MID$(a$, 109, 26)
    St$ = MID$(a$, 133, 2)
    Zip$ = MID$(a$, 136, 9)
    LatePen$ = MID$(a$, 173, 1)
    SC$ = MID$(a$, 178, 6)
    RV$ = MID$(a$, 185, 10)
    PV$ = MID$(a$, 407, 10)
    Bus$ = MID$(a$, 274, 1)
    ss$ = MID$(a$, 418, 9)
    Desc$ = MID$(a$, 658, 55)
    Pin$ = MID$(a$, 1171, 15)
    PRINT RV$, PV$, rECORD
  
    P# = P# + VAL(PV$)
    R# = R# + VAL(RV$)


    GOSUB UpdateMaster
    IF VAL(RV$) > 0 THEN GOSUB UpdateReal
    IF VAL(PV$) > 0 THEN GOSUB UpdatePers
    GOTO 10

UpdateMaster:

   'Split Name Here
    IF Bus$ = "Y" THEN
     LNM$ = NM$
     FM$ = ""
    ELSE
    kk = INSTR(NM$, ",")
    IF kk > 0 THEN
     LNM$ = LEFT$(NM$, kk - 1)
     FM$ = MID$(NM$, kk + 1, LEN(NM$) - kk)
     ELSE
     LNM$ = NM$
     FM$ = ""

     END IF

    END IF
    
   
    TaxCustRec(1).Acct = rECORD
    TaxCustRec(1).OpenDate = Date2Num(DATE$)
    TaxCustRec(1).FName = LTRIM$(FM$)
    TaxCustRec(1).LName = LTRIM$(LNM$)
    TaxCustRec(1).SName = LTRIM$(LNM$)
    TaxCustRec(1).Addr1 = Add1$
    TaxCustRec(1).Addr2 = Add2$
    TaxCustRec(1).City = City$
    TaxCustRec(1).STATE = St$
    TaxCustRec(1).Zip = LEFT$(Zip$, 5)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ss$
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).CountyAcct = VAL(CtyAcct$)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).Pin = rECORD
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, rECORD, TaxCustRec(1)
    RETURN

  
  

UpdateReal:
    PropertyRec(1).REALPIN = Pin$
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(RV$)
    PropertyRec(1).EXMPSENI = VAL(SC$)
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = LEFT$(Desc$, 20)
    PropertyRec(1).PROPNOT2 = MID$(Desc$, 21, 20)
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = rECORD
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = VAL(PV$)
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = LEFT$(Desc$, 20)
   PersRec(1).DESC2 = MID$(Desc$, 21, 20)
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = rECORD
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

