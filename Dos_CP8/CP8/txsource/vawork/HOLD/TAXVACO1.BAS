DEFINT A-Z
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
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBPPCust
  DIM SHARED TBPPCust(1) AS TBPPCust
  DIM SHARED TBPPCustBal(1) AS TBPPBAL
  DIM SHARED TBCustRec1(1) AS TBPPCust
  DIM SHARED TaxTran(1) AS TaxTransactionType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
   TaxTranRecLen = LEN(TaxTran(1))
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

    OPEN "TBPPCUST.DAT" FOR RANDOM AS #11 LEN = LEN(TBCustRec(1))
    OPEN "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = LEN(TBPPCustBal(1))
    TaxFile = FREEFILE
    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

   FOR x = 1 TO LOF(11) / LEN(TBCustRec(1))
   GET 11, x, TBCustRec(1)
   GET 11, x, TBPPCust(1)
   GET 12, x, TBPPCustBal(1)
   Record! = x

    PRINT RTRIM$(TBCustRec(1).CustName)
    kk = INSTR(RTRIM$(TBCustRec(1).CustName), " ")
   IF kk > 0 THEN
     LNM$ = LEFT$(TBCustRec(1).CustName, kk - 1)
     FME$ = MID$(TBCustRec(1).CustName, kk, LEN(TBCustRec(1).CustName) - kk)
    ELSE
     LNM$ = RTRIM$(TBCustRec(1).CustName)
     FME$ = ""
   END IF
   TaxCustRec(1).FName = LTRIM$(FME$)
   TaxCustRec(1).LName = LTRIM$(LNM$)
   TaxCustRec(1).SName = LTRIM$(LNM$)
   TaxCustRec(1).acct = Record!
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Addr1 = TBCustRec(1).Address1
   TaxCustRec(1).Addr2 = TBCustRec(1).Address2
   TaxCustRec(1).City = TBCustRec(1).City
   TaxCustRec(1).State = TBCustRec(1).State
   TaxCustRec(1).Zip = TBCustRec(1).Zip
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = TBCustRec(1).SocSec
   TaxCustRec(1).SSSN = TBCustRec(1).SocSec2
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).pin = x
   TaxCustRec(1).Deleted = 0
   IF LEN(QPTrim$(TBCustRec(1).CustName)) = 0 THEN
    TaxCustRec(1).Deleted = -1
   END IF
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatepersonal
   

   'Set Trans & Bills to Zero
    LastTrans& = 0
    BillTran& = 0

    IF (TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1).CurYrPenAmount) > 0 THEN
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     TaxTran(1).TransDate = Date2Num%("10-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1).CurYrPenAmount
     TaxTran(1).Revenue.Principle1 = TBPPCust(1).CurYrPersPropTax
     TaxTran(1).Revenue.Principle2 = TBPPCust(1).CurYrMachToolsTax
     TaxTran(1).Revenue.Principle3 = TBPPCust(1).CurYrMerchCapTax
     TaxTran(1).Revenue.Principle4 = TBPPCust(1).CurYrFarmEquipTax
     TaxTran(1).Revenue.Principle5 = TBPPCust(1).CurYrMobileHomeTax
     TaxTran(1).Revenue.Interest = TBPPCust(1).CurYrIntAmount
     TaxTran(1).Revenue.Penalty = TBPPCust(1).CurYrPenAmount
     TaxTran(1).Description = "Tax Bill"
     TaxTran(1).Posted2Gl = "Y"
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
    IF TBPPCust(1).CurYrIntAmount > 0 THEN
     TaxTran(1).TransDate = Date2Num%("07-01-1999")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = TBPPCust(1).CurYrIntAmount
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBPPCust(1).CurYrIntAmount
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
     'Check Penalty
    IF TBPPCust(1).CurYrPenAmount > 0 THEN
     TaxTran(1).TransDate = Date2Num%("12-08-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = TBPPCust(1).CurYrPenAmount
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBPPCust(1).CurYrPenAmount
     TaxTran(1).Description = "Tax Bill Penalty"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF

     CLOSE TaxTran

  'Update Tax File if Past Bal >0

     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
    
    FOR ll = 1 TO 10
      ThisYear = (1998 - ll)
    IF (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).PenAmt(ll) + TBPPCustBal(1).IntAmt(ll)) THEN
     Yr$ = LTRIM$(STR$(ThisYear))
     TaxTran(1).TransDate = Date2Num%(("10-01-" + Yr$))
     TaxTran(1).TaxYear = ThisYear
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).PenAmt(ll) + TBPPCustBal(1).IntAmt(ll _
))
     TaxTran(1).Revenue.Principle1 = TBPPCustBal(1).PerPropTaxAmt(ll)
     TaxTran(1).Revenue.Principle2 = TBPPCustBal(1).FarmEquipAmt(ll)
     TaxTran(1).Revenue.Principle3 = TBPPCustBal(1).MachToolsAmt(ll)
     TaxTran(1).Revenue.Principle4 = TBPPCustBal(1).MerchCapAmt(ll)
     TaxTran(1).Revenue.Principle5 = TBPPCustBal(1).MobileHomeAmt(ll)
     TaxTran(1).Revenue.Interest = TBPPCustBal(1).IntAmt(ll)
     TaxTran(1).Revenue.Penalty = TBPPCustBal(1).PenAmt(ll)
     TaxTran(1).Description = "Tax Bill"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
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
    IF TBPPCustBal(1).IntAmt(ll) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("07-01-1999")
     TaxTran(1).TaxYear = ThisYear
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = TBPPCustBal(1).IntAmt(ll)
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBPPCustBal(1).IntAmt(ll)
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
    IF TBPPCustBal(1).PenAmt(ll) > 0 THEN
     TaxTran(1).TransDate = Date2Num%(("12-05-" + Yr$))
     TaxTran(1).TaxYear = ThisYear
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = "P"
     TaxTran(1).Amount = TBPPCustBal(1).PenAmt(ll)
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBPPCustBal(1).PenAmt(ll)
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
    NEXT ll
    CLOSE TaxTran
   

Nextone:
   NEXT x
   CLOSE
   EXIT SUB


updatepersonal:
   PersRec(1).PERSVAL = TBCustRec(1).PersProp
   PersRec(1).MCVALUE = TBCustRec(1).MerchCap
   PersRec(1).CVALUE = TBCustRec(1).FarmEquip
   PersRec(1).MTVALUE = TBCustRec(1).MachTools
   PersRec(1).MHVALUE = TBCustRec(1).MobileHome
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("03-01-1999")
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
   PersRec(1).LastYrPrinted = 1998
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

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

