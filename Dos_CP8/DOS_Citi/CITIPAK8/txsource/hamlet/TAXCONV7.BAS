
DEFINT A-Z
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
DECLARE FUNCTION Date2Num% (TheDate$)
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

  'Hamlet Data From Richmond Cty
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED InRec(1) AS HamletRecType

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM A$(120)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   TotalReal# = 0
   TotalPers# = 0

  OPEN "DATA.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 514
   FOR Cnt& = 1 TO LOF(11) / 514
   GET 11, Cnt&, InRec(1)
   RealValue# = VAL(InRec(1).RealValue)
   TotalReal# = TotalReal# + RealValue#
   

   PersValue# = VAL(InRec(1).HouseHoldValue) + VAL(InRec(1).InventoryValue) + VAL(InRec(1).MachValue) + VAL(InRec(1).VehValue)
   TotalPers# = TotalPers# + PersValue#

   SCExp# = VAL(InRec(1).SCEx)
   TotalExp# = TotalExp# + SCExp#

   Record = Record + 1
   'Split Name Here
   kk = INSTR(InRec(1).Name1, " ")
   IF kk > 0 THEN
    LN$ = LEFT$(InRec(1).Name1, kk - 1)
    SN$ = LEFT$(InRec(1).Name1, kk - 1)
     IF kk + 1 > LEN(RTRIM$(InRec(1).Name1)) THEN
      ff1$ = LN$
     ELSE
     ff1$ = MID$(InRec(1).Name1, kk + 1, LEN(RTRIM$(InRec(1).Name1)) - (kk + 1))
     END IF
    ELSE
    LN$ = RTRIM$(InRec(1).Name1)
    SN$ = RTRIM$(InRec(1).Name1)
    ff1$ = ""

   END IF

   TaxCustRec(1).OpenDate = Date2Num("08-01-99")
   TaxCustRec(1).Acct = Record
   TaxCustRec(1).FNAME = ff1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   TaxCustRec(1).Addr1 = InRec(1).Name2
   TaxCustRec(1).Addr2 = InRec(1).Add1
   TaxCustRec(1).City = InRec(1).City
   TaxCustRec(1).STATE = ""
   TaxCustRec(1).Zip = InRec(1).ZipCode
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = InRec(1).SSN
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).CountyAcct = VAL(InRec(1).AcctNumb)
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   IF RealValue# > 0 THEN
    GOSUB UpdateReal
   END IF
   IF PersValue# > 0 THEN
   GOSUB UpdatePers
   END IF

   NEXT Cnt&
   
   
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = InRec(1).PIN
    PropertyRec(1).PROPDATE = Date2Num("08-01-1999")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = InRec(1).DeedBook
    PropertyRec(1).BLOCK = InRec(1).DeedPage
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = InRec(1).AorL
    PropertyRec(1).PROPSIZE = VAL(InRec(1).Acres)
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = InRec(1).Late
    PropertyRec(1).MortCode = InRec(1).MortCode
    PropertyRec(1).PROPVALU = RealValue#
    PropertyRec(1).EXMPSENI = SCExp#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = InRec(1).LandDesc
    PropertyRec(1).PROPNOT2 = InRec(1).Notes
    PropertyRec(1).PROPNOT3 = "Mort Code: " + InRec(1).MortCode
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
   PersRec(1).PERSVAL = PersValue#
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
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

