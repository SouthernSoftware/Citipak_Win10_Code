DEFINT A-Z
'Real Estate Conversion from Text File Information  BROOKNEAL VA 2003
'Must Delete Current TaxProp.Dat File BEFORE Running
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

   OPEN "BNREAL.TXT" FOR INPUT AS #11
1 LINE INPUT #11, a$

   PRINT VAL(LEFT$(a$, 10))
   spin$ = MID$(a$, 192, 17)
   Use# = VAL(MID$(a$, 241, 9)): t1# = t1# + Use#
   IF Use# > 0 THEN
   Real# = Use#: t2# = t2# + Real#
   ELSE
   Real# = VAL(MID$(a$, 249, 9)): t2# = t2# + Real#
   END IF
   bldg# = VAL(MID$(a$, 259, 9)): t3# = t3# + bldg#
   PRINT spin$
   PRINT STRING$(79, "-")


   'Find Exiting Acct
   FOR Snt& = 1 TO NumOfTaxRecs
    GET TaxFile, Snt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(LEFT$(a$, 10)) THEN
      GOSUB ChkName
      GOSUB Updatereal
      GOTO 1
    END IF
    
   
   NEXT Snt&
   
   'Add New Acct Here
   'Decode Name Here
   anam$ = MID$(a$, 11, 35)
   kk = INSTR(anam$, " ")
   IF kk = 0 THEN
    ln$ = anam$
    FM$ = ""
    ELSE
    ln$ = LEFT$(anam$, kk - 1)
    FM$ = LTRIM$(MID$(anam$, kk + 1, LEN(anam$) - kk))
   END IF

   SN$ = ln$
   ln$ = UCASE$(ln$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)

   City$ = RTRIM$(UCASE$(MID$(a$, 106, 30)))
   State$ = RIGHT$(City$, 2)
   
   IF LEN(City$) > 0 THEN
    City$ = LEFT$(City$, LEN(City$) - 2)
   END IF
   Addr1$ = QPTrim$(UCASE$(MID$(a$, 46, 30)))
   Addr2$ = QPTrim$(UCASE$(MID$(a$, 76, 30)))
   IF LEN(Addr1$) = 0 THEN Addr1$ = Addr2$: Addr2$ = ""


   

   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   TaxCustRec(1).FNAME = FM$
   TaxCustRec(1).LName = ln$
   TaxCustRec(1).SName = SN$
    
   TaxCustRec(1).Acct = Record!
   
   TaxCustRec(1).Addr1 = Addr1$
   TaxCustRec(1).Addr2 = Addr2$
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = MID$(a$, 137, 5)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).CountyAcct = VAL(LEFT$(a$, 10))
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
   TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
   PUT TaxFile, Record!, TaxCustRec(1)
   Snt& = Record!
   PropNumb = 1
   GOSUB Updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


Updatereal:

    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = spin$
    PropertyRec(1).PROPDATE = Date2Num%("10-14-2003")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = LEFT$(spin$, 6)
    PropertyRec(1).Block = MID$(spin$, 7, 6)
    PropertyRec(1).LotNumb = MID$(spin$, 13, 6)
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = bldg#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = MID$(a$, 142, 25)
    PropertyRec(1).PROPNOT2 = MID$(a$, 167, 25)
    PropertyRec(1).PROPNOT3 = MID$(a$, 192, 25)
    PropertyRec(1).CustPin = Snt&
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2002
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN

ChkName:
   anam$ = MID$(a$, 11, 35)
   kk = INSTR(anam$, " ")
   IF kk = 0 THEN
    ln$ = anam$
    FM$ = ""
    ELSE
    ln$ = LEFT$(anam$, kk - 1)
    FM$ = LTRIM$(MID$(anam$, kk + 1, LEN(anam$) - kk))
   END IF

   SN$ = ln$
   ln$ = QPTrim$(UCASE$(ln$))
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)
   CLNAME$ = UCASE$(TaxCustRec(1).FNAME) + " " + UCASE$(TaxCustRec(1).LName)
   CLNAME$ = QPTrim$(CLNAME$)

   Chk1 = INSTR(CLNAME$, ln$)
   IF Chk1 > 0 THEN RETURN



   City$ = RTRIM$(UCASE$(MID$(a$, 106, 30)))
   State$ = RIGHT$(City$, 2)

   IF LEN(City$) > 0 THEN
    City$ = LEFT$(City$, LEN(City$) - 2)
   END IF
   Addr1$ = QPTrim$(UCASE$(MID$(a$, 46, 30)))
   Addr2$ = QPTrim$(UCASE$(MID$(a$, 76, 30)))
   IF LEN(Addr1$) = 0 THEN Addr1$ = Addr2$: Addr2$ = ""




   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   TaxCustRec(1).FNAME = FM$
   TaxCustRec(1).LName = ln$
   TaxCustRec(1).SName = SN$

   TaxCustRec(1).Acct = Record!

   TaxCustRec(1).Addr1 = Addr1$
   TaxCustRec(1).Addr2 = Addr2$
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = MID$(a$, 137, 5)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).CountyAcct = VAL(LEFT$(a$, 10))
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
   TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
   PUT TaxFile, Record!, TaxCustRec(1)
   Snt& = Record!
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

