DEFINT A-Z
'Real Estate Conversion from Text File Information  WARSAW VA 2000
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
  DIM SHARED REALRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS
   GOSUB InitReal
   OPEN "REWARSAW.TXT" FOR INPUT AS #11
1  FOR c = 1 TO 17
    INPUT #11, a$(c)
   NEXT c
   PRINT a$(1), a$(2)


   spin$ = RTRIM$(a$(1))
   Real# = VAL(a$(10))
   Exempt# = VAL(a$(11))
   Bldg# = VAL(a$(12))
   tReal# = tReal# + Real#
   tPers# = tPers# + Bldg#
   LOCATE 5, 1
   PRINT spin$
   PRINT "REAL: "; USING "##########,#.##"; tReal#
   PRINT "BLDG: "; USING "##########,#.##"; tPers#
   PRINT " TOT: "; USING "##########,#.##"; tReal# + tPers#


   

   'Find Exiting Acct
   FOR Snt& = 1 TO NumOfPropRecs
   GET PropTaxFile, Snt&, PropertyRec(1)
    Tpin$ = RTRIM$(PropertyRec(1).RealPin)
    
   IF Tpin$ = spin$ AND LEN(Tpin$) > 0 THEN
   
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = Bldg#
    PropertyRec(1).EXMPOTHR = Exempt#
    PropertyRec(1).PropSize = VAL(a$(9))
    PropertyRec(1).PropNot1 = a$(15)
    PropertyRec(1).PropNot2 = "DB REF# " + a$(14)
    PropertyRec(1).PropNot3 = ""
    PropertyRec(1).PropDate = Date2Num%("09-01-2000")
    PUT PropTaxFile, Snt&, PropertyRec(1)
    GET TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
    class = VAL(a$(8))


   

    
     TaxCustRec(1).Addr1 = a$(3)
     TaxCustRec(1).Addr2 = a$(4)
     TaxCustRec(1).City = a$(5)
     TaxCustRec(1).State = a$(6)
     TaxCustRec(1).Zip = LEFT$((a$(7)), 5)
     TaxCustRec(1).CountyAcct = VAL(a$(17))
    PUT TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
    GOTO 1
   END IF
   NEXT Snt&




STOP
   'Add New Acct Here
   'Decode Name Here
   anam$ = RTRIM$(a$(2))
   kk = INSTR(anam$, " ")
   IF class = 4 THEN
    LN$ = anam$
    FM$ = ""
    ELSE
    LN$ = LEFT$(anam$, kk - 1)
    FM$ = LTRIM$(MID$(anam$, kk + 1, LEN(anam$) - kk))
   END IF

   SN$ = LN$
   LN$ = UCASE$(LN$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)

   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   TaxCustRec(1).FNAME = FM$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = SN$

     TaxCustRec(1).Addr1 = a$(3)
     TaxCustRec(1).Addr2 = a$(4)
     TaxCustRec(1).City = a$(5)
     TaxCustRec(1).State = a$(6)
     TaxCustRec(1).Zip = a$(7)
     TaxCustRec(1).CountyAcct = VAL(a$(17))
    
   TaxCustRec(1).Acct = Record!

   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-2000")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = a$(1)
    PropertyRec(1).PropDate = Date2Num%("02-01-2000")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).Block = ""
    PropertyRec(1).LotNumb = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PropSize = VAL(a$(9))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = Bldg#
    PropertyRec(1).EXMPOTHR = Exempt#
    PropertyRec(1).PropNot1 = a$(15)
    PropertyRec(1).PropNot2 = "DB Ref " + a$(14)
    PropertyRec(1).PropNot3 = ""

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN

InitReal:
   FOR Cnt& = 1 TO NumOfPropRecs
   GET PropTaxFile, Cnt&, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
   PUT PropTaxFile, Cnt&, PropertyRec(1)
   NEXT Cnt&
   RETURN


END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

