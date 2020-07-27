
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
  DIM SHARED TBCustRec(1) AS TBCustType

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM A$(120)

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
   OPEN "CARSHORE.TXT" FOR INPUT AS #11
   Record = NumOfTaxRecs
1 FOR x = 1 TO 28
10 INPUT #11, A$(x)
   PRINT x, A$(x)
   NEXT x
   PRINT Record
  'SLEEP 1
   
   Record = Record + 1
   'Split Name Here
   IF Record >= 35 THEN
     kk = INSTR(A$(2), " ")
    IF kk > 0 THEN
     LN$ = LEFT$(A$(2), kk - 1)
     SN$ = LEFT$(A$(2), kk - 1)
     FF1$ = MID$(A$(2), kk + 1, LEN(A$(2)) - (kk + 1))
     ELSE
     LN$ = RTRIM$(A$(2))
     SN$ = RTRIM$(A$(2))
     FF1$ = ""
    END IF
    ELSE
     LN$ = RTRIM$(A$(2))
     SN$ = RTRIM$(A$(2))
     FF1$ = ""

   END IF

   'Strip $Sign From A$(14),A$(21)
   kk = INSTR(A$(14), "$")
   A$(14) = MID$(A$(14), kk + 1, LEN(A$(14)) - kk)
   kk = INSTR(A$(21), "$")
   A$(21) = MID$(A$(21), kk + 1, LEN(A$(21)) - kk)


   TaxCustRec(1).Acct = Record
   TaxCustRec(1).OpenDate = Date2Num%(DATE$)
   TaxCustRec(1).FName = FF1$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   TaxCustRec(1).Addr1 = A$(4)
   TaxCustRec(1).Addr2 = A$(5)
   TaxCustRec(1).City = A$(6)
   TaxCustRec(1).State = A$(7)
   TaxCustRec(1).Zip = LEFT$(A$(8), 5)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).CSSN = A$(9)
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


   GOTO 1
   
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropertyRec(1).REALPIN = A$(10)
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = A$(17)
    PropertyRec(1).BLOCK = A$(18)
    PropertyRec(1).LOTNUMB = A$(19)
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(A$(14))
    PropertyRec(1).EXMPSENI = VAL(A$(15))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = A$(11)
    PropertyRec(1).PROPNOT2 = A$(12)
    PropertyRec(1).PROPNOT3 = A$(13)
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
   PersRec(1).PROPPIN = A$(20)
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PERSVAL = VAL(A$(21))
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = VAL(A$(22))
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = A$(23)
   PersRec(1).DESC2 = A$(24)
   PersRec(1).DESC3 = A$(25)
   PersRec(1).DESC4 = A$(26)
   PersRec(1).DESC5 = A$(27)
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

