DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  Independence Va  2002
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TAXFILE%)
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
CLS
  ON ERROR RESUME NEXT
  DIM a$(88)
  OpenTaxCustFile NUMOFTAXRECS, TAXFILE
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

'convert ppowner first

     OPEN "I", 10, "ppowner.txt"
c1:  IF EOF(10) THEN GOTO n1
     FOR cc = 1 TO 17
     INPUT #10, a$(cc)
     NEXT cc

     CLOSE Pfile
     Pfile = 11
     OPEN "PPOWNER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #Pfile LEN = 8
     IOwner(1).Owner = VAL(a$(1))
     IOwner(1).NameID = VAL(a$(2))
     PUT Pfile, LOF(Pfile) / 8 + 1, IOwner(1)
     GOTO c1




n1: OPEN "indeppp.TXT" FOR INPUT AS #13
1 FOR cc = 1 TO 40
  INPUT #13, a$(cc)
  IF ERR = 62 THEN
  CLOSE
  PRINT "Done"
  END
  END IF

  NEXT cc
  PRINT a$(2), a$(4), a$(7)
  PRINT STRING$(79, "-")

ReRun:
   FOR Cnt& = 1 TO NUMOFTAXRECS
    GET TAXFILE, Cnt&, TaxCustRec(1)
    nf = 1
   IF TaxCustRec(1).CountyAcct = VAL(a$(2)) THEN
    GOSUB updatepersonal
    nf = 0
    lt = 0
    EXIT FOR
   END IF
   NEXT Cnt&
n2:
    IF nf = 1 AND lt = 0 THEN
     'Find File in PPOwner.dat
     CLOSE Pfile: OPEN "PPOWNER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #Pfile LEN = 8
     FOR PCnt& = 1 TO LOF(Pfile) / 8
      GET Pfile, PCnt&, IOwner(1)
        IF IOwner(1).Owner = VAL(a$(2)) THEN
           a$(2) = STR$(IOwner(1).NameID)
           CLOSE Pfile
           GOTO ReRun
        END IF
      NEXT PCnt&
      CLOSE Pfile
      
     
    END IF
    
Nextone:
   lt = 0
   GOTO 1
   CLOSE
   EXIT SUB


updatepersonal:

   IF VAL(a$(4)) = 12 THEN RETURN

   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF VAL(a$(4)) = 11 THEN
    PersRec(1).MHVALUE = VAL(a$(5))
    MH# = MH# + VAL(a$(5))
    GOTO upd1
   END IF
   IF VAL(a$(4)) = 13 THEN
    PersRec(1).MTVALUE = VAL(a$(5))
    MT# = MT# + VAL(a$(5))
    GOTO upd1
   END IF
    PV# = PV# + VAL(a$(5))
    PersRec(1).PERSVAL = VAL(a$(5))
upd1:
   a# = a# + VAL(a$(5))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("11-01-2004")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(14)
   PersRec(1).DESC2 = RTRIM$(a$(11)) + "/" + RTRIM$(a$(12)) + "/" + RTRIM$(a$(13))
   PersRec(1).DESC3 = a$(15)
   PersRec(1).DESC4 = LEFT$(a$(10), 5)
   IF VAL(a$(7)) > 0 THEN
    PersRec(1).Desc5 = "Y" + "/" + STRING$(24, " ") + "2004"
   ELSE
    PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2004"
   END IF
   PersRec(1).CustPin = Cnt&
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2003
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN



END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

