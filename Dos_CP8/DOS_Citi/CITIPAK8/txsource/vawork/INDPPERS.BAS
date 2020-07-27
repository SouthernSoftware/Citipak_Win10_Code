DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  Independence Va  2002
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TaxFile%)
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
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TaxConv.Bi'

  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED IOwner(1) AS IOwner

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  ON ERROR RESUME NEXT

  CLS

  CALL KillFile("taxpers.dat")
  DIM a$(88)


  OpenTaxCustFile NUMOFTAXRECS, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile



n1:
  OPEN "pers2005.txt" FOR INPUT AS #13
  DidHeader = False
1 FOR cc = 1 TO 85
    IF ERR = 62 THEN
      CLOSE
      PRINT
      PRINT "Processed:"; DidCnt
      PRINT "    Found:"; Updated
      PRINT " New Cust:"; NotFound
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Error:"; ERR
      END
    END IF
    INPUT #13, a$(cc)
  NEXT cc
  IF NOT DidHeader THEN
    DidHeader = True
    GOTO 1
  END IF
  DidCnt = DidCnt + 1
  PRINT a$(1), a$(2), a$(6)
  PRINT STRING$(79, "-")

'  GOTO 1

ReRun:

   'Nf = True
   NUMOFTAXRECS = LOF(TaxFile) / LEN(TaxCustRec(1))
   FOR Cnt& = 1 TO NUMOFTAXRECS
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(a$(1)) THEN
      Record! = Cnt&
      Updated = Updated + 1
      GOSUB UpdateAcct
      GOSUB updatepersonal
      GOTO 1
     ' Nf = False
     ' EXIT FOR
    END IF
   
   NEXT Cnt&

   'IF Nf THEN
     NotFound = NotFound + 1
     Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
     GOSUB UpdateAcct
     GOSUB updatepersonal
   'END IF
   GOTO 1

    
UpdateAcct:

  TaxCustRec(1).FNAME = a$(7)
  TaxCustRec(1).LName = a$(6)
  TaxCustRec(1).SName = (QPTrim$(a$(7)) + QPTrim$(a$(6)))

  TaxCustRec(1).Addr1 = a$(11)
  TaxCustRec(1).Addr2 = a$(12)
  TaxCustRec(1).acct = Record!
  TaxCustRec(1).City = QPTrim$(a$(13))
  TaxCustRec(1).STATE = QPTrim$(a$(14))
  TaxCustRec(1).zip = QPTrim$(a$(15))
  TaxCustRec(1).CSSN = QPTrim$(a$(5))
  TaxCustRec(1).SSSN = ""
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "Y"
  TaxCustRec(1).CountyAcct = VAL(a$(1))
  'TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).PIN = Record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  TaxCustRec(1).OPENDATE = Date2Num%(DATE$)

'****************************
  PUT TaxFile, Record!, TaxCustRec(1)
  RETURN

updatepersonal:

   IF VAL(a$(49)) = 12 THEN RETURN

   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF VAL(a$(49)) = 11 THEN           'mobile home
    PersRec(1).MHVALUE = VAL(a$(50))
    MH# = MH# + VAL(a$(50))
    GOTO upd1
   END IF
   IF VAL(a$(49)) = 13 THEN           'machine tools
    PersRec(1).MTVALUE = VAL(a$(50))
    MT# = MT# + VAL(a$(50))
    GOTO upd1
   END IF
   PV# = PV# + VAL(a$(50))
   PersRec(1).PERSVAL = VAL(a$(50))

upd1:
   a# = a# + VAL(a$(50))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("11-14-2005")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(58)
   PersRec(1).DESC2 = RTRIM$(a$(55)) + "/" + RTRIM$(a$(56)) + "/" + RTRIM$(a$(57))
   PersRec(1).DESC3 = a$(57)
   PersRec(1).DESC4 = "" 'LEFT$(a$(10), 5)

   IF QPTrim$(a$(65)) = "P" THEN
     IF VAL(a$(49)) = 1 THEN
       PersRec(1).Desc5 = "Y" + "/" + STRING$(24, " ") + "2005"
     ELSE
       PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2005"
     END IF
   ELSE
     PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2005"
   END IF

   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

END SUB

SUB keepme

END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

