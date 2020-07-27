DEFINT A-Z
'TOWN OF VICTORIA CONVERSION FROM LUNENBURG COUNTY REAL ESTATE AFTER BILLS RUN
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
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'lctax.bi'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType

  DIM SHARED TBVeh(1) AS PPVehType
  DIM SHARED TBCustRec(1) AS TBPPCust

  DIM SHARED TaxTran(1) AS TaxTransactionType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS

   TaxTranRecLen = LEN(TaxTran(1))
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   'Init Prop File to Zero Values
   FOR Cnt& = 1 TO LOF(PersTaxFile) / LEN(PersRec(1))
   GET PersTaxFile, Cnt&, PersRec(1)
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
   PUT PersTaxFile, Cnt&, PersRec(1)
   NEXT Cnt&


   OPEN "F:\TBPTCUST.DAT" FOR RANDOM AS #11 LEN = LEN(TBCustRec(1))
   VFile = FREEFILE
   OPEN "F:\TBTVEH.DAT" FOR RANDOM AS #VFile LEN = LEN(TBVeh(1))
   TaxFile = FREEFILE
   OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
   NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

   FOR x = 1 TO LOF(11) / LEN(TBCustRec(1))
   GET 11, x, TBCustRec(1)
    Updated = 0

   IF VAL(TBCustRec(1).District) = 10 AND LEFT$(LTRIM$(TBCustRec(1).TBCurYrBillNumber), 2) = "99" THEN

    FOR Snt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
       GET TaxFile, Snt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = x THEN
      Record! = Snt&
      GOSUB UpdateCustomer
      Updated = 1
      EXIT FOR
    END IF

    NEXT Snt&

    IF Updated = 0 THEN
     'Must Create New Customer First
    Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  
  kk = INSTR(TBCustRec(1).CustName, " ")
  lasn$ = LEFT$(TBCustRec(1).CustName, kk - 1)
  FirstN$ = MID$(TBCustRec(1).CustName, kk + 1, (LEN(TBCustRec(1).CustName) - kk))

  TaxCustRec(1).acct = Record!
  TaxCustRec(1).OPENDATE = Date2Num("12-01-1999")
  TaxCustRec(1).FName = FirstN$
  TaxCustRec(1).LName = lasn$
  TaxCustRec(1).SName = lasn$
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = LEFT$(TBCustRec(1).SocSec, 3) + MID$(TBCustRec(1).SocSec, 5, 2) + MID$(TBCustRec(1).SocSec, 8, 4)
  TaxCustRec(1).SSSN = LEFT$(TBCustRec(1).SocSec2, 3) + MID$(TBCustRec(1).SocSec2, 4, 2) + MID$(TBCustRec(1).SocSec2, 7, 4)
  TaxCustRec(1).Addr1 = TBCustRec(1).Address1
  TaxCustRec(1).Addr2 = TBCustRec(1).Address2
  TaxCustRec(1).City = TBCustRec(1).City
  TaxCustRec(1).State = TBCustRec(1).State
  TaxCustRec(1).Zip = TBCustRec(1).Zip
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "Y"
  TaxCustRec(1).Employer = ""
  TaxCustRec(1).Bankrupt = "N"
  TaxCustRec(1).TownShip = ""
  TaxCustRec(1).PAD1 = ""
  TaxCustRec(1).CountyAcctString = ""
  TaxCustRec(1).CountyAcct = x
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).pin = Record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  PUT TaxFile, Record!, TaxCustRec(1)
  
  GOSUB UpdateCustomer
    
    END IF

   END IF 'end search for victoria accounts
   NEXT x

   CLOSE
   STOP

   

       
UpdateCustomer:
   GOSUB UpdatePers
   'Set Trans & Bills to Zero
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     TaxAmount# = (TBCustRec(1).TBCurYrPersPropTaxAmount - TBCustRec(1).TBPPTRADiscnt) + TBCustRec(1).TBCurYrMachToolsTaxAmount + TBCustRec(1).TBCurYrMerchCapTaxAmount + TBCustRec(1).TBCurYrFarmEquipTaxAmount + TBCustRec(1). _
TBCurYrMobileHomeTaxAmount
     IF TaxAmount# < 0 THEN
      TaxAmount# = 0
     END IF
     IF TBCustRec(1).TBPPTRADiscnt > TBCustRec(1).TBCurYrPersPropTaxAmount THEN
      TBCustRec(1).TBPPTRADiscnt = TBCustRec(1).TBCurYrPersPropTaxAmount
     END IF


     TaxTran(1).TransDate = Date2Num%("12-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = TaxAmount#
     TaxTran(1).Revenue.Principle1 = TBCustRec(1).TBCurYrPersPropTaxAmount - TBCustRec(1).TBPPTRADiscnt
     TaxTran(1).Revenue.Principle2 = TBCustRec(1).TBCurYrMachToolsTaxAmount
     TaxTran(1).Revenue.Principle3 = TBCustRec(1).TBCurYrMerchCapTaxAmount
     TaxTran(1).Revenue.Principle4 = TBCustRec(1).TBCurYrFarmEquipTaxAmount
     TaxTran(1).Revenue.Principle5 = TBCustRec(1).TBCurYrMobileHomeTaxAmount
     TaxTran(1).Revenue.Future1 = TBCustRec(1).TBPPTRADiscnt
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Description = "Tax Bill" + TBCustRec(1).TBCurYrBillNumber
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = TaxCustRec(1).LastTrans
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
     TaxCustRec(1).LastTrans = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
     CLOSE TaxTran
     t = t + 1
     PRINT t
     RETURN

UpdatePers:
  IF TBCustRec(1).MerchCap > 0 THEN
    PersRec(1).PROPPIN = ""
    PersRec(1).PROPDATE = Date2Num("12-01-1999")
    PersRec(1).PersVal = 0
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = TBCustRec(1).MerchCap
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = "Merch Capital"
    PersRec(1).DESC2 = ""
    PersRec(1).DESC3 = ""
    PersRec(1).Desc4 = ""
    PersRec(1).Desc5 = ""
    PersRec(1).CustPin = Record!
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 1999
    PersRec(1).Deleted = 0
    PersRec(1).VehTaxYear = 1999
    PersRec(1).Blank = ""
    NextPersRec& = LOF(PersTaxFile) / LEN(PersRec(1))
    NextPersRec& = NextPersRec& + 1
    PUT PersTaxFile, NextPersRec&, PersRec(1)
    END IF
  IF TBCustRec(1).MachTools > 0 THEN
    PersRec(1).PROPPIN = ""
    PersRec(1).PROPDATE = Date2Num("12-01-1999")
    PersRec(1).PersVal = 0
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = TBCustRec(1).MachTools
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = "Merch Capital"
    PersRec(1).DESC2 = ""
    PersRec(1).DESC3 = ""
    PersRec(1).Desc4 = ""
    PersRec(1).Desc5 = ""
    PersRec(1).CustPin = Record!
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 1999
    PersRec(1).Deleted = 0
    PersRec(1).VehTaxYear = 1999
    PersRec(1).Blank = ""
    NextPersRec& = LOF(PersTaxFile) / LEN(PersRec(1))
    NextPersRec& = NextPersRec& + 1
    PUT PersTaxFile, NextPersRec&, PersRec(1)
    END IF


   IF TBCustRec(1).MobileHome > 0 THEN
    PersRec(1).PROPPIN = ""
    PersRec(1).PROPDATE = Date2Num("12-01-1999")
    PersRec(1).PersVal = 0
    PersRec(1).MHVALUE = TBCustRec(1).MobileHome
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = "Mobile Home"
    PersRec(1).DESC2 = ""
    PersRec(1).DESC3 = ""
    PersRec(1).Desc4 = ""
    PersRec(1).Desc5 = ""
    PersRec(1).CustPin = Record!
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 1999
    PersRec(1).Deleted = 0
    PersRec(1).VehTaxYear = 1999
    PersRec(1).Blank = ""
    NextPersRec& = LOF(PersTaxFile) / LEN(PersRec(1))
    NextPersRec& = NextPersRec& + 1
    PUT PersTaxFile, NextPersRec&, PersRec(1)
   END IF


  'Now Check Heavy Mach - Boats - Bus Furn
   PersValu# = TBCustRec(1).HeavyEquip + TBCustRec(1).Boats + TBCustRec(1).BusFurn

  IF PersValu# > 0 THEN
    PersRec(1).PROPPIN = ""
    PersRec(1).PROPDATE = Date2Num("12-01-1999")
    PersRec(1).PersVal = PersValu#
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = "Other Pers Property"
    PersRec(1).DESC2 = "Boats/Heavy Equip"
    PersRec(1).DESC3 = "Bus Furn"
    PersRec(1).Desc4 = ""
    PersRec(1).Desc5 = ""
    PersRec(1).CustPin = Record!
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 1999
    PersRec(1).Deleted = 0
    PersRec(1).VehTaxYear = 1999
    PersRec(1).Blank = ""
    NextPersRec& = LOF(PersTaxFile) / LEN(PersRec(1))
    NextPersRec& = NextPersRec& + 1
    PUT PersTaxFile, NextPersRec&, PersRec(1)
    END IF


    'Now Go Thru Each Vehicle and Update All 1999
  VehRecord& = TBCustRec(1).FirstVeh
  WHILE VehRecord& <> 0
  GET VFile, VehRecord&, TBVeh(1)
  IF TBVeh(1).VehTyp <> "D" THEN
   IF TBVeh(1).VehTaxYr = 1999 THEN
    PersRec(1).PROPPIN = ""
    PersRec(1).PROPDATE = Date2Num("12-01-1999")
    PersRec(1).PersVal = TBVeh(1).VehValue
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = TBVeh(1).VehVin
    PersRec(1).DESC2 = TBVeh(1).VehMake + TBVeh(1).VehModel
    PersRec(1).DESC3 = LTRIM$(STR$(TBVeh(1).VehWght))
    PersRec(1).Desc4 = TBVeh(1).VehYear
    PersRec(1).Desc5 = TBVeh(1).VehQ + STRING$(24, " ") + "1999"
    PersRec(1).CustPin = Record!
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 1999
    PersRec(1).Deleted = 0
    PersRec(1).VehTaxYear = 1999
    PersRec(1).Blank = ""
    NextPersRec& = LOF(PersTaxFile) / LEN(PersRec(1))
    NextPersRec& = NextPersRec& + 1
    PUT PersTaxFile, NextPersRec&, PersRec(1)
   END IF
  END IF
   VehRecord& = TBVeh(1).VehNext
 WEND
      
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

