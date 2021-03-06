DEFINT A-Z
'MAIN OLD TO NEW CONVERSION (NORTH CAROLINA NO BALANCES)
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
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
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
  '$INCLUDE: 'TaxConv.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCust(1) AS TBCustType
'  DIM SHARED TBCust1(1) AS TBCust1Type
  
  


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing

  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  OpenTaxCustFile NumOfTaxRecs, TaxFile

   OPEN "TBCust.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 512
    

   
1 FOR X = 1 TO LOF(11) / 512
   GET 11, X, TBCust(1)
  
   RECORD! = X
   
     IF VAL(TBCust(1).acct) > 0 THEN

  

     TotalReal# = TBCust(1).pland + TBCust(1).pbldg
     TotalReal# = INT((TotalReal# * 100) + .5) / 100
     TotalPers# = TBCust(1).LS + TBCust(1).MH + TBCust(1).FM
     TotalPers# = INT((TotalPers# * 100) + .5) / 100



 ' Split Name Here
   TaxCustRec(1).acct = RECORD!
   TaxCustRec(1).OPENDATE = Date2Num(DATE$)
   TaxCustRec(1).FName = TBCust(1).FIRSTNAME
   TaxCustRec(1).LNAME = TBCust(1).LNAME
   TaxCustRec(1).SName = TBCust(1).LNAME
   TaxCustRec(1).Addr1 = TBCust(1).Addr1
   TaxCustRec(1).Addr2 = TBCust(1).Addr2
   TaxCustRec(1).City = TBCust(1).City
   TaxCustRec(1).State = TBCust(1).State
   TaxCustRec(1).Zip = LEFT$(TBCust(1).Zip, 5)
   TaxCustRec(1).HPHONE = TBCust(1).PHONE
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = LEFT$(TBCust(1).SSN, 3) + MID$(TBCust(1).SSN, 5, 2) + MID$(TBCust(1).SSN, 8, 4)
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = RECORD!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, RECORD!, TaxCustRec(1)

    IF TotalReal# > 0 THEN
     GOSUB UpdateReal
    END IF

    IF TotalPers# > 0 THEN
     GOSUB UpdatePers
    END IF

  


    ELSE
     TaxCustRec(1).Deleted = -1
     PUT TaxFile, RECORD!, TaxCustRec(1)
    END IF      'Check VALID ACCT

   NEXT X
   
   CLOSE
   EXIT SUB


UpdateReal:
    PropertyRec(1).REALPIN = RTRIM$(TBCust(1).PMAP) + "*"
    PropertyRec(1).PROPDATE = Date2Num(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = TBCust(1).PMAP
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = TBCust(1).psize
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LateList = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCust(1).pland + TBCust(1).pbldg
    PropertyRec(1).EXMPSENI = TBCust(1).ME
    PropertyRec(1).EXMPOTHR = TBCust(1).homeexp + TBCust(1).useexp
    PropertyRec(1).PROPNOT1 = TBCust(1).pdesc
    PropertyRec(1).PROPNOT2 = TBCust(1).PMAP
    PropertyRec(1).PROPNOT3 = TBCust(1).Pdesc2
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = RECORD!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num("08-01-2001")
   PersRec(1).PersVal = TBCust(1).LS + TBCust(1).MH + TBCust(1).FM
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LateList = "N"
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = ""
   PersRec(1).Desc5 = ""
   PersRec(1).CustPin = RECORD!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
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

