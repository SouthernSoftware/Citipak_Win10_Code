
DEFINT A-Z
'Conversion for Carolina Shores MDB Data Base
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

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  OPEN "CADCUS~1.TXT" FOR INPUT AS #11
   Record = NumOfTaxRecs
1 FOR x = 1 TO 104
10 INPUT #11, a$(x)
   NEXT x
   

   Record = Record + 1
   'Split Name Here
   IF Record >= 35 THEN
     kk = INSTR(a$(3), " ")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(3), kk - 1)
     SN$ = LEFT$(a$(3), kk - 1)
     FF1$ = MID$(a$(3), kk + 1, LEN(a$(3)) - (kk + 1))
     ELSE
     LN$ = RTRIM$(a$(3))
     SN$ = RTRIM$(a$(3))
     FF1$ = ""
    END IF
    ELSE
     LN$ = RTRIM$(a$(2))
     SN$ = RTRIM$(a$(2))
     FF1$ = ""

   END IF

   'Strip $Sign From A$(14),A$(21)
   'kk = INSTR(A$(12), "$")
   'A$(14) = MID$(A$(12), kk + 1, LEN(A$(12)) - kk)
  'kk = INSTR(A$(21), "$")
  ' A$(21) = MID$(A$(21), kk + 1, LEN(A$(21)) - kk)

   IF LEN(QPTrim$(a$(4))) = 0 THEN a$(4) = a$(5): a$(5) = ""
   TaxCustRec(1).Acct = Record
   TaxCustRec(1).OpenDate = Date2Num%(DATE$)
   TaxCustRec(1).FName = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   TaxCustRec(1).Addr1 = a$(4)
   TaxCustRec(1).Addr2 = a$(5)
   TaxCustRec(1).City = a$(6)
   TaxCustRec(1).State = a$(7)
   TaxCustRec(1).Zip = LEFT$(a$(8), 5)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = a$(2)
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
'   GOSUB UpdatePers


   GOTO 1
   
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = a$(1)
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = a$(10)
    PropertyRec(1).BLOCK = a$(11)
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = a$(17)
    PropertyRec(1).PROPSIZE = VAL(a$(18))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(a$(16))
    PropertyRec(1).EXMPSENI = VAL(a$(13))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(9)
    PropertyRec(1).PROPNOT2 = a$(10)
    PropertyRec(1).PROPNOT3 = a$(11)
    'PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1998
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN
UpdatePers:
   PersRec(1).PROPPIN = a$(20)
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = VAL(a$(21))
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = VAL(a$(22))
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(23)
   PersRec(1).DESC2 = a$(24)
   PersRec(1).DESC3 = a$(25)
   PersRec(1).DESC4 = a$(26)
   PersRec(1).DESC5 = a$(27)
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

