DEFINT A-Z

'IMPORTANT MESSAGE PLEASE READ
'Convert Master PP Records from TXT File for Victoria
'MUST SET VEHICLE YEAR EACH TIME TO PULL PROPER VEHICLES FROM FILE



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


  'Stephens City Va Tax PP Conversion of Master (Customers)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'lctax.bi'

   DIM SHARED LCTaxCust(1) AS TBPPCust
   DIM SHARED LCVeh(1) AS PPVehType
   DIM SHARED LCKeNTaxCust(1) AS TBKENPPCust
   DIM SHARED LCKenVeh(1) AS PPKenVehType
   
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS

  OPEN "KENPP.TXT" FOR OUTPUT AS #10 LEN = LEN(LCKeNTaxCust(1))
  OPEN "KENVE.TXT" FOR OUTPUT AS #11 LEN = LEN(LCKenVeh(1))

  LCTaxFile = FREEFILE
  OPEN "F:TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #LCTaxFile LEN = LEN(LCTaxCust(1))
  NumOfLCTaxRecs = LOF(LCTaxFile) / LEN(LCTaxCust(1))
  VFile = FREEFILE
  OPEN "F:TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #VFile LEN = LEN(LCVeh(1))


  FOR MCnt& = 1 TO LOF(LCTaxFile) / LEN(LCTaxCust(1))
  GET LCTaxFile, MCnt&, LCTaxCust(1)
  Record! = MCnt&

  IF VAL(LCTaxCust(1).District) <= 9 AND VAL(LCTaxCust(1).TBCurYrBillNumber) > 0 THEN
        LCKeNTaxCust(1).AcctNumber = LTRIM$(STR$(Record!))
        LCKeNTaxCust(1).CustName = LCTaxCust(1).CustName
        LCKeNTaxCust(1).Address1 = LCTaxCust(1).Address1
        LCKeNTaxCust(1).Address2 = LCTaxCust(1).Address2
        LCKeNTaxCust(1).City = LCTaxCust(1).City
        LCKeNTaxCust(1).State = LCTaxCust(1).State
        LCKeNTaxCust(1).Zip = LCTaxCust(1).Zip
        LCKeNTaxCust(1).SocSec = LCTaxCust(1).SocSec
        LCKeNTaxCust(1).SocSec2 = LCTaxCust(1).SocSec2
        LCKeNTaxCust(1).PersProp = LTRIM$(STR$(LCTaxCust(1).PersProp))
        LCKeNTaxCust(1).MachTools = LTRIM$(STR$(LCTaxCust(1).MachTools))
        LCKeNTaxCust(1).MobileHome = LTRIM$(STR$(LCTaxCust(1).MobileHome))
        LCKeNTaxCust(1).MerchCap = LTRIM$(STR$(LCTaxCust(1).MerchCap))
        LCKeNTaxCust(1).FarmEquip = LTRIM$(STR$(LCTaxCust(1).FarmEquip))
        LCKeNTaxCust(1).Business = LCTaxCust(1).Business
        LCKeNTaxCust(1).District = LCTaxCust(1).District
        LCKeNTaxCust(1).LateFile = LCTaxCust(1).LateFile
        LCKeNTaxCust(1).LName2 = LCTaxCust(1).LName2
        LCKeNTaxCust(1).FName2 = LCTaxCust(1).FName2
        LCKeNTaxCust(1).HeavyEquip = LTRIM$(STR$(LCTaxCust(1).HeavyEquip))
        LCKeNTaxCust(1).Boats = LTRIM$(STR$(LCTaxCust(1).Boats))
        LCKeNTaxCust(1).BusFurn = LTRIM$(STR$(LCTaxCust(1).BusFurn))
        LCKeNTaxCust(1).TBCurYrBillNumber = LCTaxCust(1).TBCurYrBillNumber
        LCKeNTaxCust(1).TBCurYrPersPropTaxAmount = LTRIM$(STR$(LCTaxCust(1).TBCurYrPersPropTaxAmount))
        LCKeNTaxCust(1).TBCurYrFarmEquipTaxAmount = LTRIM$(STR$(LCTaxCust(1).TBCurYrFarmEquipTaxAmount))
        LCKeNTaxCust(1).TBCurYrMachToolsTaxAmount = LTRIM$(STR$(LCTaxCust(1).TBCurYrMachToolsTaxAmount))
        LCKeNTaxCust(1).TBCurYrMerchCapTaxAmount = LTRIM$(STR$(LCTaxCust(1).TBCurYrMerchCapTaxAmount))
        LCKeNTaxCust(1).TBCurYrMobileHomeTaxAmount = LTRIM$(STR$(LCTaxCust(1).TBCurYrMobileHomeTaxAmount))
        LCKeNTaxCust(1).TBLateFeeAmount = LTRIM$(STR$(LCTaxCust(1).TBLateFeeAmount))
        LCKeNTaxCust(1).TBPPTRAValue = LTRIM$(STR$(LCTaxCust(1).TBPPTRAValue))
        LCKeNTaxCust(1).TBPPTRADiscnt = LTRIM$(STR$(LCTaxCust(1).TBPPTRADiscnt))

        PRINT #10, LCKeNTaxCust(1).AcctNumber;
        PRINT #10, LCKeNTaxCust(1).CustName;
        PRINT #10, LCKeNTaxCust(1).Address1;
        PRINT #10, LCKeNTaxCust(1).Address2;
        PRINT #10, LCKeNTaxCust(1).City;
        PRINT #10, LCKeNTaxCust(1).State;
        PRINT #10, LCKeNTaxCust(1).Zip;
        PRINT #10, LCKeNTaxCust(1).SocSec;
        PRINT #10, LCKeNTaxCust(1).SocSec2;
        PRINT #10, LCKeNTaxCust(1).PersProp;
        PRINT #10, LCKeNTaxCust(1).MachTools;
        PRINT #10, LCKeNTaxCust(1).MobileHome;
        PRINT #10, LCKeNTaxCust(1).MerchCap;
        PRINT #10, LCKeNTaxCust(1).FarmEquip;
        PRINT #10, LCKeNTaxCust(1).Business;
        PRINT #10, LCKeNTaxCust(1).District;
        PRINT #10, LCKeNTaxCust(1).LateFile;
        PRINT #10, LCKeNTaxCust(1).LName2;
        PRINT #10, LCKeNTaxCust(1).FName2;
        PRINT #10, LCKeNTaxCust(1).HeavyEquip;
        PRINT #10, LCKeNTaxCust(1).Boats;
        PRINT #10, LCKeNTaxCust(1).BusFurn;
        PRINT #10, LCKeNTaxCust(1).TBCurYrBillNumber;
        PRINT #10, LCKeNTaxCust(1).TBCurYrPersPropTaxAmount;
        PRINT #10, LCKeNTaxCust(1).TBCurYrFarmEquipTaxAmount;
        PRINT #10, LCKeNTaxCust(1).TBCurYrMachToolsTaxAmount;
        PRINT #10, LCKeNTaxCust(1).TBCurYrMerchCapTaxAmount;
        PRINT #10, LCKeNTaxCust(1).TBCurYrMobileHomeTaxAmount;
        PRINT #10, LCKeNTaxCust(1).TBLateFeeAmount;
        PRINT #10, LCKeNTaxCust(1).TBPPTRAValue;
        PRINT #10, LCKeNTaxCust(1).TBPPTRADiscnt

        VRec! = LCTaxCust(1).FirstVeh
        WHILE VRec! <> 0
        GET VFile, VRec!, LCVeh(1)
        IF LCVeh(1).VehTyp <> "D" THEN
        IF LCVeh(1).VehTaxYr = 2004 THEN
        LCKenVeh(1).VehTyp = LCVeh(1).VehTyp
        LCKenVeh(1).VehVin = LCVeh(1).VehVin
        LCKenVeh(1).VehMake = LCVeh(1).VehMake
        LCKenVeh(1).VehYear = LCVeh(1).VehYear
        LCKenVeh(1).VehModel = LCVeh(1).VehModel
        LCKenVeh(1).VehWght = LTRIM$(STR$(LCVeh(1).VehWght))
        LCKenVeh(1).VehValue = LTRIM$(STR$(LCVeh(1).VehValue))
        LCKenVeh(1).VehQ = LCVeh(1).VehQ
        LCKenVeh(1).VehLic = LCVeh(1).VehLic
        LCKenVeh(1).VehAcct = LTRIM$(STR$(LCVeh(1).VehAcct))

         PRINT #11, LCKenVeh(1).VehTyp;
         PRINT #11, LCKenVeh(1).VehVin;
         PRINT #11, LCKenVeh(1).VehMake;
         PRINT #11, LCKenVeh(1).VehYear;
         PRINT #11, LCKenVeh(1).VehModel;
         PRINT #11, LCKenVeh(1).VehWght;
         PRINT #11, LCKenVeh(1).VehValue;
         PRINT #11, LCKenVeh(1).VehQ;
         PRINT #11, LCKenVeh(1).VehLic;
         PRINT #11, LCKenVeh(1).VehAcct
         END IF
        END IF
        VRec! = LCVeh(1).VehNext
        WEND


  END IF

  NEXT MCnt&
  CLOSE
  STOP
  END

END SUB

