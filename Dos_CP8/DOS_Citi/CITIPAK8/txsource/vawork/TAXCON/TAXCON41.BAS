DEFINT A-Z
'Convert PP Records from TXT File (New Market 2003)
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(55)

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "NMPERS.TXT" FOR INPUT AS #11
   CLS
   b = 0
1 LINE INPUT #11, a$
 
   IF MID$(a$, 26, 1) = "2" THEN GOTO 1
  a$(7) = MID$(a$, 58, 35)
  a$(3) = MID$(a$, 34, 5)
  a$(18) = LTRIM$(RTRIM$(MID$(a$, 195, 9)))
  a$(19) = LTRIM$(RTRIM$(MID$(a$, 205, 9)))
  a$(8) = MID$(a$, 93, 30)
  a$(9) = MID$(a$, 123, 30)
  a$(10) = MID$(a$, 153, 30)
  a$(11) = MID$(a$, 184, 5)

   PRINT "Name "; a$(7)
   PRINT "Acct Number "; a$(3)
   PRINT "SS# "; a$(18)
   PRINT STRING$(79, "-")
                       
   TaxFile = FREEFILE
   OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
   NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
   FOR Srch! = 1 TO NumOfTaxRecs
   GET TaxFile, Srch!, TaxCustRec(1)
   IF LEN(a$(18)) = 8 THEN a$(18) = "0" + a$(18)
   IF LEN(a$(18)) = 7 THEN a$(18) = "00" + a$(18)
   IF RTRIM$(TaxCustRec(1).CSSN) = RTRIM$(a$(18)) OR RTRIM$(TaxCustRec(1).SSSN) = RTRIM$(a$(18)) OR TaxCustRec(1).CountyAcct = VAL(a$(3)) THEN
   
     GOSUB UpdateRecord
     CLOSE TaxFile
     GOTO 1
   

   END IF
   NEXT Srch!

   


   'Must Add Record : Could Not Find

  'Decode Name Here
    nm$ = UCASE$(a$(7))
    IF INSTR(nm$, "INC") > 0 THEN
     LN$ = UCASE$(a$(7))
     FM$ = ""
     SN$ = LN$
    ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(7), kk - 1)
     FM$ = RIGHT$(a$(7), (LEN(a$(7)) - kk))
     SN$ = LN$
      ELSE
     LN$ = UCASE$(a$(7))
     FM$ = ""
     SN$ = LN$
    END IF

    END IF

    
    Record! = NumOfTaxRecs + 1


    
    TaxCustRec(1).FNAME = FM$
    TaxCustRec(1).LName = LN$
    TaxCustRec(1).SName = SN$
    TaxCustRec(1).Acct = Record!
    IF LEN(RTRIM$(a$(8))) = 0 THEN
     TaxCustRec(1).Addr1 = a$(9)
     TaxCustRec(1).Addr2 = ""
     ELSE
     TaxCustRec(1).Addr1 = a$(8)
     TaxCustRec(1).Addr2 = a$(9)
    END IF
    TaxCustRec(1).City = a$(10)
    TaxCustRec(1).STATE = ""
    TaxCustRec(1).Zip = a$(11)
    TaxCustRec(1).HPHONE = a$(25)
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = a$(18)
    TaxCustRec(1).SSSN = a$(19)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(a$(3))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("05-01-2003")
    PUT TaxFile, Record!, TaxCustRec(1)
    CLOSE TaxFile

    GOSUB UPDATEPERS
NextOne:
   GOTO 1
   CLOSE
   EXIT SUB






UpdateRecord:

     TaxCustRec(1).OPENDATE = Date2Num%("05-01-2003")
     TaxCustRec(1).CountyAcct = VAL(a$(3))
     PUT TaxFile, Srch!, TaxCustRec(1)
     CLOSE TaxFile
     Record! = Srch!
     GOSUB UPDATEPERS
     RETURN
   RETURN


UPDATEPERS:
'   kk = INSTR(a$, "STATUTORY")
'   IF kk > 0 THEN RETURN


   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0


   IF MID$(a$, 457, 2) = "MH" THEN
    PersRec(1).MHVALUE = VAL(MID$(a$, 437, 9))
    mh# = mh# + VAL(MID$(a$, 437, 9))
    GOTO upd1
   END IF
   IF MID$(a$, 457, 2) = "MC" THEN
    PersRec(1).MCVALUE = VAL(MID$(a$, 437, 9))
    RETURN
   END IF
   IF MID$(a$, 457, 2) = "MT" THEN
    PersRec(1).MTVALUE = VAL(MID$(a$, 437, 9))
    mt# = mt# + VAL(MID$(a$, 437, 9))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(MID$(a$, 437, 9))
   CAR# = CAR# + VAL(MID$(a$, 437, 9))
upd1:
   a# = a# + VAL(MID$(a$, 437, 9))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("05-01-2003")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = MID$(a$, 417, 17)
   PersRec(1).DESC2 = MID$(a$, 377, 10) + "/" + MID$(a$, 387, 10) + "/" + MID$(a$, 402, 10)
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = MID$(a$, 503, 4)
   IF MID$(a$, 507, 1) = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2003"
   ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2003"
   END IF
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN
  
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))

END SUB

