DEFINT A-Z
'Conversion for Sunset Beach Data
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBCustType
  DIM SHARED SS(1) AS SunsetBeachType

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
   OPEN "CITY11" FOR INPUT AS #11
   Record = NumOfTaxRecs
1 LINE INPUT #11, a$(1)
   PRINT a$(1)
   a = a + 1
   PRINT a; STRING$(74, "-")
   a$ = MID$(a$(1), 1, 8)
   a$(2) = MID$(a$(1), 9, 35)
   a$(3) = MID$(a$(1), 44, 25)
   a$(4) = MID$(a$(1), 69, 25)
   a$(5) = MID$(a$(1), 94, 18)
   a$(6) = MID$(a$(1), 112, 2)
   a$(7) = MID$(a$(1), 114, 5)
   a$(8) = MID$(a$(1), 119, 15)
   a$(9) = MID$(a$(1), 134, 50)
   a$(10) = MID$(a$(1), 184, 4)
   a$(11) = MID$(a$(1), 188, 4)
   a$(12) = MID$(a$(1), 192, 9)
   a$(13) = MID$(a$(1), 201, 9)
   a$(14) = MID$(a$(1), 210, 9)
   a$(15) = MID$(a$(1), 219, 9)
   a$(16) = MID$(a$(1), 228, 9)
   a$(17) = MID$(a$(1), 237, 1)
   a$(18) = MID$(a$(1), 238, 7)
   a$(19) = MID$(a$(1), 245, 9)
   a$(20) = MID$(a$(1), 254, 9)
   a$(21) = MID$(a$(1), 263, 9)
   a$(22) = MID$(a$(1), 272, 9)
   a$(23) = MID$(a$(1), 281, 9)
   a$(24) = MID$(a$(1), 290, 9)
   a$(25) = MID$(a$(1), 299, 9)
   a$(26) = MID$(a$(1), 308, 1)
   a$(27) = MID$(a$(1), 309, 6)
   a$(28) = MID$(a$(1), 315, 2)
   a$(29) = MID$(a$(1), 317, 2)
   a$(30) = MID$(a$(1), 319, 1)
   a$(31) = MID$(a$(1), 320, 5)
   a$(32) = MID$(a$(1), 325, 8)

   Record = Record + 1
   'Split Name Here
   IF Record >= 641 THEN
     kk = INSTR(a$(2), " ")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(2), kk - 1)
     SN$ = LEFT$(a$(2), kk - 1)
     FF1$ = MID$(a$(2), kk + 1, LEN(a$(2)) - (kk + 1))
     ELSE
     LN$ = RTRIM$(a$(2))
     SN$ = RTRIM$(a$(2))
     FF1$ = ""
    END IF
    ELSE
     LN$ = RTRIM$(a$(2))
     SN$ = RTRIM$(a$(2))
     FF1$ = ""

   END IF


   TaxCustRec(1).Acct = Record
   TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
   TaxCustRec(1).FNAME = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   TaxCustRec(1).Addr1 = a$(3)
   TaxCustRec(1).Addr2 = a$(4)
   TaxCustRec(1).City = a$(5)
   TaxCustRec(1).State = a$(6)
   TaxCustRec(1).Zip = LEFT$(a$(7), 5)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   GOSUB UpdateReal
   GOSUB UpdatePers


'NEXT Cnt&
   GOTO 1
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = a$(8)
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = a$(10)
    PropertyRec(1).BLOCK = a$(11)
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = a$(17)
    PropertyRec(1).PROPSIZE = VAL(a$(18)) / 100
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(a$(16))
    PropertyRec(1).EXMPSENI = VAL(a$(27))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = LEFT$(a$(9), 30)
    PropertyRec(1).PROPNOT2 = MID$(a$(9), 31, 20)
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1998
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN
UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = VAL(a$(25))
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
   PersRec(1).LastYrPrinted = 1998
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

