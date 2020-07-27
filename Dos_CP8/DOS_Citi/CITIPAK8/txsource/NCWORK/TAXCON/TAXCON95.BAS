DEFINT A-Z
'7 Devils NC 2002
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

   OPEN "sdtax.txt" FOR INPUT AS #10


  'Init Prop File Values

    FOR Cnt& = 1 TO NumOfPropRecs
     GET PropTaxFile, Cnt&, PropertyRec(1)
      PropertyRec(1).PROPVALU = 0
      PropertyRec(1).EXMPSENI = 0
     PUT PropTaxFile, Cnt&, PropertyRec(1)
    NEXT Cnt&
   
    Record = 0
    CLS
10 LINE INPUT #10, a$
    CtyAcct$ = LEFT$(a$, 9)
    nm$ = MID$(a$, 140, 25)
    Add1$ = MID$(a$, 166, 25)
    Add2$ = MID$(a$, 190, 25)
    City$ = MID$(a$, 215, 25)

    Zip$ = MID$(a$, 240, 5)
    RV$ = MID$(a$, 1, 10)
    Desc$ = MID$(a$, 112, 28)
    
    TPin$ = QPTrim$(MID$(a$, 84, 13))
    PRINT nm$, RV$, Record
  
    
    R# = R# + VAL(RV$)



    FOR Cnt& = 1 TO NumOfPropRecs
     GET PropTaxFile, Cnt&, PropertyRec(1)
     SDPin$ = RTRIM$(PropertyRec(1).REALPIN)
         'Remove -
         SDPin1$ = ""
         FOR ll = 1 TO LEN(SDPin$)
         IF MID$(SDPin$, ll, 1) = "-" THEN
         ELSE
         SDPin1$ = SDPin1$ + MID$(SDPin$, ll, 1)
         END IF
         NEXT ll
         SDPin1$ = QPTrim$(SDPin1$)
         IF SDPin1$ = TPin$ AND PropertyRec(1).CustPin > 0 THEN

         'Found the Property Now Update the Record with Name and Address
           PropertyRec(1).PROPDATE = Date2Num%(DATE$)
           PropertyRec(1).PROPVALU = VAL(RV$)
           PUT PropTaxFile, Cnt&, PropertyRec(1)
           Record! = PropertyRec(1).CustPin
           GET TaxFile, Record!, TaxCustRec(1)
              kk = INSTR(nm$, ",")
              IF kk > 0 THEN
                LNM$ = LEFT$(nm$, kk - 1)
                FM$ = MID$(nm$, kk + 1, LEN(nm$) - kk)
              ELSE
                LNM$ = nm$
                FM$ = ""
              END IF
            TaxCustRec(1).FName = LTRIM$(FM$)
            TaxCustRec(1).LName = LTRIM$(LNM$)
            TaxCustRec(1).SName = LTRIM$(LNM$)
            TaxCustRec(1).Addr1 = Add1$
            TaxCustRec(1).Addr2 = Add2$
            TaxCustRec(1).City = City$
            TaxCustRec(1).Zip = LEFT$(Zip$, 5)
            TaxCustRec(1).CountyAcctString = QPTrim$(MID$(a$, 84, 13))
            PUT TaxFile, Record!, TaxCustRec(1)
           GOTO 10
         
         END IF
         
         
     PUT PropTaxFile, Cnt&, PropertyRec(1)
    NEXT Cnt&

UpdateMaster:
'Split Name Here

    kk = INSTR(nm$, ",")
    IF kk > 0 THEN
     LNM$ = LEFT$(nm$, kk - 1)
     FM$ = MID$(nm$, kk + 1, LEN(nm$) - kk)
     ELSE
     LNM$ = nm$
     FM$ = ""
    END IF

  
    
    Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).OpenDate = Date2Num(DATE$)
    TaxCustRec(1).FName = LTRIM$(FM$)
    TaxCustRec(1).LName = LTRIM$(LNM$)
    TaxCustRec(1).SName = LTRIM$(LNM$)
    TaxCustRec(1).Addr1 = Add1$
    TaxCustRec(1).Addr2 = Add2$
    TaxCustRec(1).City = City$
    TaxCustRec(1).STATE = ""
    TaxCustRec(1).Zip = LEFT$(Zip$, 5)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ss$
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).CountyAcctString = QPTrim$(MID$(a$, 84, 13))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).Pin = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, Record!, TaxCustRec(1)
    GOSUB UpdateReal
    GOTO 10
    
  
  

UpdateReal:
    PropertyRec(1).REALPIN = QPTrim$(MID$(a$, 84, 13))
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
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = LEFT$(Desc$, 20)
    PropertyRec(1).PROPNOT2 = MID$(Desc$, 21, 20)
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
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

