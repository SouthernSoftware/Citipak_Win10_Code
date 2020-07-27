DEFINT A-Z
'Conversion for Sunset Beach Data from county codes
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

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  CLS
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

' First Initialize All Values from Property Files
' Pers Prop First
   FOR Snt& = 1 TO NumOfPersRecs
   GET PersTaxFile, Snt&, PersRec(1)
   PersRec(1).PERSVAL = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
  PUT PersTaxFile, Snt&, PersRec(1)
  NEXT Snt&

 ' Now Real Estate
  FOR Snt& = 1 TO NumOfPropRecs
  GET PropTaxFile, Snt&, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).CustPin = 0
  PUT PropTaxFile, Snt&, PropertyRec(1)
  NEXT Snt&

  'Now Reset Each Acct to 1 property only
   FOR Snt& = 1 TO NumOfTaxRecs
   GET TaxFile, Snt&, TaxCustRec(1)
   IF TaxCustRec(1).FirstPropRec > 0 THEN
    GET PropTaxFile, TaxCustRec(1).FirstPropRec, PropertyRec(1)
    PropertyRec(1).CustPin = Snt&
    PUT PropTaxFile, TaxCustRec(1).FirstPropRec, PropertyRec(1)
   END IF
  NEXT Snt&

  OPEN "SUNSET.TXT" FOR INPUT AS #11
  
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



 FPin$ = RTRIM$(a$(8))


   'Split Name Here
   IF a > 647 THEN
     kk = INSTR(a$(2), " ")

    IF kk > 0 THEN
     IF kk > 25 THEN STOP
     LN$ = LEFT$(a$(2), kk - 1)
     SN$ = LEFT$(a$(2), kk - 1)
     ff1$ = MID$(a$(2), kk + 1, LEN(a$(2)) - (kk + 1))
     ELSE
     LN$ = RTRIM$(a$(2))
     SN$ = RTRIM$(a$(2))
     ff1$ = ""
    END IF
   ELSE
     LN$ = RTRIM$(a$(2))
     SN$ = RTRIM$(a$(2))
     ff1$ = ""
   END IF

  IF LEFT$(LN$, 6) = "ZUCKER" THEN STOP

'if no pin then try and match the name
'Determine if Pers or Real Property Record by Checking Value in a25 and a16
'Pers Property Record

IF VAL(a$(30)) > 0 THEN GOTO 1
CustomerRecord! = 0
'Real Prop Record

IF VAL(a$(16)) > 0 THEN
 GOSUB FindPin
 IF Found$ = "YES" THEN
  GET PropTaxFile, Record!, PropertyRec(1)
  PropertyRec(1).PROPVALU = PropertyRec(1).PROPVALU + VAL(a$(16))
  PropertyRec(1).EXMPSENI = VAL(a$(27))
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).LastYrPrinted = 1999
  PropertyRec(1).CustPin = CRecord!
  TOTREAL# = TOTREAL# + VAL(a$(16))
  PUT PropTaxFile, Record!, PropertyRec(1)
  GET TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
  CustomerRecord! = CRecord!
  'IF TaxCustRec(1).Deleted <> 0 THEN STOP
  ELSE
  'Add Customer and Record
  GOSUB UpdateClient
  GOSUB UpdateReal
  IF VAL(a$(25)) > 0 THEN
   GOSUB UpdatePers
   GOTO 1
  END IF

 END IF
END IF

IF VAL(a$(25)) > 0 THEN
 IF CustomerRecord! = 0 THEN
   GOSUB NoPin
   
  ELSE
   GET TaxFile, CustomerRecord!, TaxCustRec(1)
   Found$ = "YES"
  END IF

  IF Found$ = "YES" THEN
   GET PersTaxFile, TaxCustRec(1).FirstPersRec, PersRec(1)
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = PersRec(1).PERSVAL + VAL(a$(25))
   PersRec(1).LastYrPrinted = 1999
   TOTPERS# = TOTPERS# + VAL(a$(25))
   PUT PersTaxFile, TaxCustRec(1).FirstPersRec, PersRec(1)
  
  ELSE
  'Add customer and Property
   GOSUB UpdateClient
   GOSUB UpdatePers
  END IF
END IF


  
 GOTO 1





UpdateClient:
   Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1

   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
   TaxCustRec(1).FName = ff1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SNAME = LN$
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
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record!, TaxCustRec(1)
   RETURN

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
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    TOTREAL# = TOTREAL# + VAL(a$(16))
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
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1999
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   TOTPERS# = TOTPERS# + VAL(a$(25))
   RETURN

NoPin:
 IF LEN(ff1$) > 15 THEN ff1$ = LEFT$(ff1$, 15)
 SNAME$ = LTRIM$(QPTrim$(ff1$) + " " + QPTrim$(LN$))
 LKUP = INSTR(SNAME$, "BRAUN")
 'IF LKUP > 0 THEN STOP
 FOR Snt& = 1 TO NumOfTaxRecs
 GET TaxFile, Snt&, TaxCustRec(1)
 CName$ = LTRIM$(QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName))
 IF RTRIM$(CName$) = RTRIM$(SNAME$) THEN
  Record! = Snt&
  Found$ = "YES"
  RETURN
 END IF
 NEXT Snt&
 Found$ = "NO"
 RETURN


FindPin:
 FPin$ = RTRIM$(a$(8))
 FOR Snt& = 1 TO NumOfPropRecs
  GET PropTaxFile, Snt&, PropertyRec(1)
  IF RTRIM$(PropertyRec(1).REALPIN) = FPin$ AND PropertyRec(1).CustPin > 0 THEN
   Found$ = "YES"
   Record! = Snt&
   CRecord! = PropertyRec(1).CustPin
   RETURN
  END IF
 NEXT Snt&
  Found$ = "NO"
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

