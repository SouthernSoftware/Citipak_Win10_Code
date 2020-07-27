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



'if no pin then try and match the name
'Determine if Pers or Real Property Record by Checking Value in a25 and a16
'Pers Property Record
IF VAL(a$(30)) > 0 THEN GOTO 1

 

IF VAL(a$(25)) > 0 THEN
  GOSUB NoPin
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
  'IF TaxCustRec(1).Deleted <> 0 THEN STOP
  ELSE
  'Add Customer and Record
  GOSUB UpdateClient
  GOSUB UpdateReal
 END IF

END IF
  
 GOTO 1





UpdateClient:
   LPRINT "Client: "; a$
   RETURN

UpdateReal:
    LPRINT "REAL"
    LPRINT STRING$(79, "-")
    RETURN

UpdatePers:
   LPRINT "PERS RECORD"
   LPRINT STRING$(79, "-")
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

