  DEFINT A-Z
  'Convert Master PP Records from TXT File (Independence 2002)
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
DECLARE FUNCTION Round# (b#)
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
  
  
  'New Market Va Tax PP Conversion of Master (Customers)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED taxcustrec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  ON ERROR RESUME NEXT
  CLS
  DIM a$(55)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  OPEN "indname.txt" FOR INPUT AS #11
  CLS
   RecCnt# = 0
  'RecCnt# = RecCnt# + 1
  b = 0
1
  FOR cc = 1 TO 27
    INPUT #11, a$(cc)
    IF ERR = 62 THEN
      CLOSE
      PRINT "Done"
      END
    END IF

  NEXT cc
  RecCnt# = RecCnt# + 1
  PRINT "Name "; a$(3)
  PRINT "Acct Number "; a$(1)
  PRINT "SS# "; a$(2)
  PRINT STRING$(79, "-")
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(taxcustrec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(taxcustrec(1))
  FOR Srch! = 1 TO NumOfTaxRecs
    GET TaxFile, Srch!, taxcustrec(1)
    IF taxcustrec(1).CountyAcct = VAL(a$(1)) THEN
      GOSUB UpdateRecord
      CLOSE TaxFile
      GOTO 1
    END IF
  NEXT Srch!
  'Must Add Record : Could Not Find
  
  'Decode Name Here
  
  LN$ = UCASE$(a$(3))
  FM$ = UCASE$(a$(4) + " " + RTRIM$(a$(5)))
  SN$ = LN$
  
  Record! = NumOfTaxRecs + 1

  taxcustrec(1).FNAME = LTRIM$(FM$)
  taxcustrec(1).LName = LTRIM$(LN$)
  taxcustrec(1).SName = LTRIM$(SN$)
  taxcustrec(1).Acct = Record!
  taxcustrec(1).Addr1 = a$(8)
  taxcustrec(1).Addr2 = a$(9)
  taxcustrec(1).City = a$(10)
  taxcustrec(1).STATE = a$(11)
  taxcustrec(1).Zip = a$(12)
  taxcustrec(1).HPHONE = a$(13)
  taxcustrec(1).WPHONE = ""
  taxcustrec(1).CSSN = a$(2)
  taxcustrec(1).SSSN = a$(14)
  taxcustrec(1).Active = "Y"
  taxcustrec(1).Interest = "Y"
  taxcustrec(1).TaxExempt = "N"
  taxcustrec(1).Penalty = "Y"
  taxcustrec(1).CountyAcct = VAL(a$(1))
  taxcustrec(1).LastTrans = 0
  taxcustrec(1).FirstPropRec = 0
  taxcustrec(1).FirstPersRec = 0
  taxcustrec(1).PIN = Record!
  taxcustrec(1).Deleted = 0
  taxcustrec(1).FileVer = 8
  taxcustrec(1).OPENDATE = Date2Num%("11-05-2004")
  PUT TaxFile, Record!, taxcustrec(1)
  CLOSE TaxFile
NextOne:
  GOTO 1
  CLOSE
  EXIT SUB
  
  
UpdateRecord:
  taxcustrec(1).CountyAcct = VAL(a$(1))
  LN$ = UCASE$(a$(3))
  FM$ = UCASE$(a$(4) + " " + RTRIM$(a$(5)))
  SN$ = LN$
  
  taxcustrec(1).FNAME = LTRIM$(FM$)
  taxcustrec(1).LName = LTRIM$(LN$)
  taxcustrec(1).SName = LTRIM$(SN$)
  taxcustrec(1).Addr1 = a$(8)
  taxcustrec(1).Addr2 = a$(9)
  taxcustrec(1).City = a$(10)
  taxcustrec(1).STATE = a$(11)
  taxcustrec(1).Zip = a$(12)
  taxcustrec(1).HPHONE = a$(13)
  taxcustrec(1).WPHONE = ""
  taxcustrec(1).CSSN = a$(2)
  taxcustrec(1).SSSN = a$(14)
  
  PUT TaxFile, Srch!, taxcustrec(1)
  RETURN
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(taxcustrec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(taxcustrec(1))
  
END SUB

