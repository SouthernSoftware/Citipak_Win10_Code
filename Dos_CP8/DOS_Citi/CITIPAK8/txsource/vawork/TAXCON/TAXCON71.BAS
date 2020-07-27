DEFINT A-Z
'Real Estate Conversion from Text File Information  INDEPENDENCE VA 2001
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
DIM A$(41)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS

   GOSUB InitReal
   OPEN "INREAL.TXT" FOR INPUT AS #11
1  FOR C = 1 TO 41
    INPUT #11, A$(C)
   NEXT C

   spin$ = RTRIM$(A$(14))

   Real# = VAL(A$(19))
   Exempt# = 0
   Bldg# = VAL(A$(20))
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
    Tpin$ = UCASE$(RTRIM$(PropertyRec(1).RealPin))
    
   IF Tpin$ = spin$ AND LEN(Tpin$) > 0 THEN
   
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = Bldg#
    PropertyRec(1).EXMPOTHR = Exempt#
    PropertyRec(1).PropSize = 0
    PropertyRec(1).PropNot1 = A$(13)
    PropertyRec(1).PropNot2 = ""
    PropertyRec(1).PropNot3 = ""
    PropertyRec(1).PropDate = Date2Num%("11-04-2002")
    PUT PropTaxFile, Snt&, PropertyRec(1)
    GET TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)

   'Add New Acct Here
   'Decode Name Here
   Anam$ = RTRIM$(A$(4))
   kk = INSTR(Anam$, ",")
   IF kk = 0 THEN
    kk = INSTR(Anam$, " ")
   END IF

   LN$ = LEFT$(Anam$, kk - 1)
   FM$ = LTRIM$(MID$(Anam$, kk + 1, LEN(Anam$) - kk))

   SN$ = LN$
   LN$ = UCASE$(LN$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)

   

   TaxCustRec(1).FNAME = FM$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = SN$
   
    
     TaxCustRec(1).Addr1 = A$(6)
     TaxCustRec(1).Addr2 = A$(7)
     TaxCustRec(1).City = A$(8)
     TaxCustRec(1).State = A$(9)
     TaxCustRec(1).Zip = LEFT$((A$(10)), 5)
     TaxCustRec(1).CountyAcct = VAL(A$(1))
     
    PUT TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
    GOTO 1
   END IF
   NEXT Snt&


   'Add New Acct Here
   'Decode Name Here
   Anam$ = RTRIM$(A$(4))
   kk = INSTR(Anam$, ",")
   IF kk = 0 THEN
    kk = INSTR(Anam$, " ")
   END IF
   
   LN$ = LEFT$(Anam$, kk - 1)
   FM$ = LTRIM$(MID$(Anam$, kk + 1, LEN(Anam$) - kk))

   SN$ = LN$
   LN$ = UCASE$(LN$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)

   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   TaxCustRec(1).FNAME = FM$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = SN$

     TaxCustRec(1).Addr1 = A$(6)
     TaxCustRec(1).Addr2 = A$(7)
     TaxCustRec(1).City = A$(8)
     TaxCustRec(1).State = A$(9)
     TaxCustRec(1).Zip = A$(10)
     TaxCustRec(1).CountyAcct = VAL(A$(1))
    
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
   TaxCustRec(1).OPENDATE = Date2Num%("11-05-2002")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = A$(1)
    PropertyRec(1).PropDate = Date2Num%("11-05-2002")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).Block = ""
    PropertyRec(1).LotNumb = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PropSize = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = Bldg#
    PropertyRec(1).EXMPOTHR = Exempt#
    PropertyRec(1).PropNot1 = A$(13)
    PropertyRec(1).PropNot2 = ""
    PropertyRec(1).PropNot3 = ""

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2001
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

