DEFINT A-Z
'Real Estate Conversion from Text File Information WINDSOR VA 2003
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumofTaxRecs%, TaxFile%)
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
  '$INCLUDE: 'TAXCONV.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  DIM SHARED RAMAST(1) AS REALTYPE
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(60)
  OpenTaxCustFile NumofTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  Record! = NumofTaxRecs

   CLS
   OPEN "windreal.txt" FOR INPUT AS #11
1   FOR ll = 1 TO 60
     INPUT #11, a$(ll)
    NEXT ll
    PRINT a$(6)

    IF VAL(a$(33)) = 7 THEN
     GOTO 1
    END IF

    IF VAL(a$(20)) > 0 THEN
     a# = a# + VAL(a$(20))
     ELSE
     a# = a# + VAL(a$(18))
    END IF
    a1# = a1# + VAL(a$(18))
    a2# = a2# + VAL(a$(20))
    s# = s# + VAL(a$(19))
    e# = e# + VAL(a$(13))
    e1# = e1# + VAL(a$(14))
    e2# = e2# + VAL(a$(35))

  
    FOR Cnt& = 1 TO NumofTaxRecs
    GET TaxFile, Cnt&, TaxCustRec(1)
    

    IF RTRIM$(TaxCustRec(1).CountyAcctString) = RTRIM$(a$(3)) THEN
    nm$ = a$(6)
    KK = INSTR(nm$, " ")
    kk1 = INSTR(nm$, "INC ")
    kk2 = INSTR(nm$, "LLC ")

    IF KK > 0 THEN
      TaxCustRec(1).FName = MID$(nm$, KK + 1, LEN(nm$) - KK)
      TaxCustRec(1).LName = LEFT$(nm$, KK - 1)
      TaxCustRec(1).SName = LEFT$(nm$, KK - 1)
     ELSE
      TaxCustRec(1).FName = ""
      TaxCustRec(1).LName = RTRIM$(a$(6))
      TaxCustRec(1).SName = RTRIM$(a$(6))
    END IF
    IF kk1 > 0 OR kk2 > 0 THEN
      TaxCustRec(1).FName = ""
      TaxCustRec(1).LName = RTRIM$(a$(6))
      TaxCustRec(1).SName = RTRIM$(a$(6))
    END IF

   
   TaxCustRec(1).Addr1 = a$(8)
   TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).CITY = a$(9)
    TaxCustRec(1).STATE = a$(10)
    TaxCustRec(1).ZIP = a$(11)
    TaxCustRec(1).CountyAcctString = a$(3)
    TaxCustRec(1).OPENDATE = Date2Num%("11-15-2003")
    PUT TaxFile, Cnt&, TaxCustRec(1)
    Record! = Cnt&
    GOSUB updatereal
    GOTO 1
    END IF
    NEXT Cnt&
   








 
    Record! = Record! + 1
    nm$ = a$(6)
    KK = INSTR(nm$, " ")
    kk1 = INSTR(nm$, "INC ")
    kk2 = INSTR(nm$, "LLC ")

    IF KK > 0 THEN
      TaxCustRec(1).FName = MID$(nm$, KK + 1, LEN(nm$) - KK)
      TaxCustRec(1).LName = LEFT$(nm$, KK - 1)
      TaxCustRec(1).SName = LEFT$(nm$, KK - 1)
     ELSE
      TaxCustRec(1).FName = ""
      TaxCustRec(1).LName = RTRIM$(a$(6))
      TaxCustRec(1).SName = RTRIM$(a$(6))
    END IF
    IF kk1 > 0 OR kk2 > 0 THEN
      TaxCustRec(1).FName = ""
      TaxCustRec(1).LName = RTRIM$(a$(6))
      TaxCustRec(1).SName = RTRIM$(a$(6))
    END IF
    
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = a$(8)
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).CITY = a$(9)
   TaxCustRec(1).STATE = a$(10)
   TaxCustRec(1).ZIP = a$(11)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = a$(3)
   TaxCustRec(1).CountyAcct = VAL(a$(1))
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("11-15-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = a$(4)
    PROPERTYRec(1).PROPDATE = Date2Num%("11-15-2003")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    IF VAL(a$(31)) > 0 THEN
    PROPERTYRec(1).LOTACRE = "L"
    PROPERTYRec(1).PROPSIZE = VAL(a$(31))
    ELSE
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(a$(32))
    END IF
    
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    IF VAL(a$(20)) > 0 THEN
     PROPERTYRec(1).PROPVALU = VAL(a$(20))
     ELSE
     PROPERTYRec(1).PROPVALU = VAL(a$(18))
    END IF
    PROPERTYRec(1).EXMPSENI = VAL(a$(19))
    PROPERTYRec(1).EXMPOTHR = 0
    PROPERTYRec(1).PROPNOT1 = "CLASS " + a$(33)
    PROPERTYRec(1).PROPNOT2 = a$(37)
    PROPERTYRec(1).PROPNOT3 = a$(38)
    
    PROPERTYRec(1).CustPin = Record!
    PROPERTYRec(1).NextRec = 0
    PROPERTYRec(1).LastYrPrinted = 2002
    PROPERTYRec(1).Deleted = 0
    PROPERTYRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PROPERTYRec(1)
    RETURN



END SUB

SUB OpenTaxCustFile (NumofTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumofTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PROPERTYRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PROPERTYRec(1))
END SUB

