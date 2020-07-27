DEFINT A-Z
'MAIN OLD TO NEW CONVERSION (NORTH CAROLINA W/1999 BALANCES ONLY)
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCust(1) AS TBCustType
  DIM SHARED TBCust1(1) AS TBCust1Type
  DIM SHARED TBCust2(1) AS TBCust2Type

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing

  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile


   OPEN "TBCust.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 512
   TaxTranRecLen = LEN(TaxTran(1))

   Record! = 0
1 FOR x = 1 TO LOF(11) / 512
   GET 11, x, TBCust(1)
 
 

   
     IF VAL(TBCust(1).acct) > 0 THEN

     GOSUB FigureBalance

     TotalReal# = TBCust(1).pland + TBCust(1).pbldg
     TotalReal# = INT((TotalReal# * 100) + .5) / 100
     TotalPers# = TBCust(1).ls + TBCust(1).mh + TBCust(1).fm
     TotalPers# = INT((TotalPers# * 100) + .5) / 100

   IF CurBal# > 0 OR TotalReal# > 0 OR TotalPers# > 0 THEN

   

    OpenTaxCustFile NumOfTaxRecs, TaxFile
     
    Record! = NumOfTaxRecs + 1

   'Split Name Here
   'Lilesville put entire name into the last name field
   'This routine takes the leftmost name and puts it in last name
   'The rest goes to first name
     LasName$ = TBCust(1).LName
     kk = INSTR(LasName$, " ")
     IF kk > 0 THEN
       LsName$ = LEFT$(LasName$, kk - 1)
       FsName$ = MID$(LasName$, kk + 1, LEN(LasName$) - kk)
       ELSE
       LsName$ = LasName$
     END IF



   TaxCustRec(1).acct = Record!
   TaxCustRec(1).FNAME = FsName$
   TaxCustRec(1).LName = LsName$
   TaxCustRec(1).SName = LsName$
   TaxCustRec(1).Addr1 = TBCust(1).Addr1
   TaxCustRec(1).Addr2 = TBCust(1).Addr2
   TaxCustRec(1).City = TBCust(1).City
   TaxCustRec(1).STATE = TBCust(1).STATE
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
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, Record!, TaxCustRec(1)

    IF TotalReal# > 0 THEN
     GOSUB UpdateReal
    END IF

    IF TotalPers# > 0 THEN
     GOSUB UpdatePers
    END IF

    BillType$ = "C"

    IF CurBal# > 0 THEN
     GOSUB UpdateCurTrans
    END IF
    
    END IF      ' Check for Valid Acct
    END IF      'Check Value or Balance
    CLOSE TaxFile
   NEXT x
   
   CLOSE
   EXIT SUB

UpdateCurTrans:
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
   IF Round#((TBCust(1).CTax)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CTax
     TaxTran(1).Revenue.Principle1 = TBCust(1).CTax
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust(1).CNotice)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = 0
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF

     'Check Interest
'    IF Round#((TBCust(1).CInterest)) > 0 THEN
'     TaxTran(1).TransDate = Date2Num%("09-01-1999")
'     TaxTran(1).TaxYear = 1999
'     TaxTran(1).TranType = 4
'     TaxTran(1).BillType = BillType$
'     TaxTran(1).Amount = TBCust(1).CInterest
'     TaxTran(1).Revenue.Principle1 = 0
'     TaxTran(1).Revenue.Interest = TBCust(1).CInterest
'     TaxTran(1).Revenue.Penalty = 0
'     TaxTran(1).Revenue.Collection = 0
'     TaxTran(1).Description = "Interest"
'     TaxTran(1).Posted2GL = "Y"
'     TaxTran(1).CustomerRec = Record!
'     TaxTran(1).LastTrans = LastTrans&
'     TaxTran(1).BelongTo = BillTran&
'     TaxTran(1).DMVSubmitted = ""
'     TaxTran(1).DMVBatch = 0
'     TaxTran(1).Padding = ""
'     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
'     PUT TaxTran, NextTran&, TaxTran(1)
'     GET TaxFile, Record!, TaxCustRec(1)
'      TaxCustRec(1).LastTrans = NextTran&
'      LastTrans& = NextTran&
'     PUT TaxFile, Record!, TaxCustRec(1)
'    END IF
'     'Check Penalty
'    IF Round#((TBCust(1).CLate)) > 0 THEN
'     TaxTran(1).TransDate = Date2Num%("09-01-1999")
'     TaxTran(1).TaxYear = 1999
'     TaxTran(1).TranType = 5
'     TaxTran(1).BillType = BillType$
'     TaxTran(1).Amount = TBCust(1).CLate
'     TaxTran(1).Revenue.Principle1 = 0
'     TaxTran(1).Revenue.Interest = 0
'     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
'     TaxTran(1).Revenue.Collection = 0
'     TaxTran(1).Description = "Penalty"
'     TaxTran(1).Posted2GL = "Y"
'     TaxTran(1).CustomerRec = Record!
'     TaxTran(1).LastTrans = LastTrans&
'     TaxTran(1).BelongTo = BillTran&
'     TaxTran(1).DMVSubmitted = ""
'     TaxTran(1).DMVBatch = 0
'     TaxTran(1).Padding = ""
'     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
'     PUT TaxTran, NextTran&, TaxTran(1)
'     GET TaxFile, Record!, TaxCustRec(1)
'      TaxCustRec(1).LastTrans = NextTran&
'      LastTrans& = NextTran&
'     PUT TaxFile, Record!, TaxCustRec(1)
'    END IF
    CLOSE TaxTran
    RETURN



  

UpdateReal:
    PropertyRec(1).REALPIN = TBCust(1).acct
    PropertyRec(1).PROPDATE = -32767
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = TBCust(1).MAP
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = TBCust(1).psize
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCust(1).pland + TBCust(1).pbldg
    PropertyRec(1).EXMPSENI = TBCust(1).ME
    PropertyRec(1).EXMPOTHR = TBCust(1).homeexp + TBCust(1).Dogs
    PropertyRec(1).PROPNOT1 = TBCust(1).pdesc
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1999
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = -32767
   PersRec(1).PERSVAL = TBCust(1).ls + TBCust(1).mh + TBCust(1).fm
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = ""
   PersRec(1).DESC5 = ""
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1999
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN

FigureBalance:
     CurBal# = 0: PastDue# = 0
     CurBal# = TBCust(1).CLate + TBCust(1).CTax + TBCust(1).CInterest + TBCust(1).CAdv
     CurBal# = Round#(CurBal#)
     IF CurBal# > 99999999 THEN CurBal# = 0
     PastDue# = 0
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

