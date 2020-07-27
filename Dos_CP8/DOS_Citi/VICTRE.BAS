DEFINT A-Z
'Real Estate Conversion from Text File Information  KENBRIDGE VA
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
  '$INCLUDE: 'LCTax.BI'
  
  DIM SHARED LCTaxCust(1) AS TBRECust
  DIM SHARED KNTaxCust(1) AS TBREKENCust
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM a$(36)

  OPEN "VICTREAL.TXT" FOR OUTPUT AS #10 LEN = LEN(KNTaxCust(1))

  LCTaxFile = FREEFILE
  OPEN "f:TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #LCTaxFile LEN = LEN(LCTaxCust(1))
  
 'Main Loop
  FOR MCnt& = 1 TO LOF(LCTaxFile) / LEN(LCTaxCust(1))
  GET LCTaxFile, MCnt&, LCTaxCust(1)

  IF VAL(LCTaxCust(1).District) <> 9 AND LCTaxCust(1).MPCurYrTaxAmount > 0 THEN
   KNTaxCust(1).AccountNumber = LTRIM$(STR$(MCnt&))
   KNTaxCust(1).MapNumber = QPTrim$(LCTaxCust(1).MapNumber)
   KNTaxCust(1).District = LCTaxCust(1).District
   KNTaxCust(1).CustName = LCTaxCust(1).CustName
   KNTaxCust(1).Address1 = QPTrim$(LCTaxCust(1).Address1)
   KNTaxCust(1).Address2 = QPTrim$(LCTaxCust(1).Address2)
   KNTaxCust(1).Address3 = QPTrim$(LCTaxCust(1).Address3)
   KNTaxCust(1).Zipcode = QPTrim$(LCTaxCust(1).Zipcode)
   KNTaxCust(1).PropDesc1 = QPTrim$(LCTaxCust(1).PropDesc1)
   KNTaxCust(1).PropDesc2 = QPTrim$(LCTaxCust(1).PropDesc2)
   KNTaxCust(1).PropDesc3 = QPTrim$(LCTaxCust(1).PropDesc3)
   KNTaxCust(1).PropDesc4 = QPTrim$(LCTaxCust(1).PropDesc4)
   KNTaxCust(1).Zone = QPTrim$(LCTaxCust(1).Zone)
   KNTaxCust(1).ClassCode = QPTrim$(LCTaxCust(1).ClassCode)
   KNTaxCust(1).Tract = QPTrim$(LCTaxCust(1).Tract)
   KNTaxCust(1).SqFtSize = QPTrim$(LCTaxCust(1).SqFtSize)
   KNTaxCust(1).Acres = QPTrim$(LCTaxCust(1).Acres)
   KNTaxCust(1).Comment1 = QPTrim$(LCTaxCust(1).Comment1)
   KNTaxCust(1).Comment2 = QPTrim$(LCTaxCust(1).Comment2)
   KNTaxCust(1).Comment3 = QPTrim$(LCTaxCust(1).Comment3)
   KNTaxCust(1).LandValue = QPTrim$(STR$(LCTaxCust(1).LandValue))
   KNTaxCust(1).BldgValue = QPTrim$(STR$(LCTaxCust(1).BldgValue))
   KNTaxCust(1).LandUseValue = QPTrim$(STR$(LCTaxCust(1).LandUseValue))
   KNTaxCust(1).MPCurYrBillNumber = LCTaxCust(1).MPCurYrBillNumber
   KNTaxCust(1).MPCurYrTaxAmount = LTRIM$(STR$(LCTaxCust(1).MPCurYrTaxAmount))
   TOT# = TOT# + LCTaxCust(1).MPCurYrTaxAmount
   PRINT #10, KNTaxCust(1).AccountNumber;
   PRINT #10, KNTaxCust(1).MapNumber;
   PRINT #10, KNTaxCust(1).District;
   PRINT #10, KNTaxCust(1).CustName;
   PRINT #10, KNTaxCust(1).Address1;
   PRINT #10, KNTaxCust(1).Address2;
   PRINT #10, KNTaxCust(1).Address3;
   PRINT #10, KNTaxCust(1).Zipcode;
   PRINT #10, KNTaxCust(1).PropDesc1;
   PRINT #10, KNTaxCust(1).PropDesc2;
   PRINT #10, KNTaxCust(1).PropDesc3;
   PRINT #10, KNTaxCust(1).PropDesc4;
   PRINT #10, KNTaxCust(1).Zone;
   PRINT #10, KNTaxCust(1).ClassCode;
   PRINT #10, KNTaxCust(1).Tract;
   PRINT #10, KNTaxCust(1).SqFtSize;
   PRINT #10, KNTaxCust(1).Acres;
   PRINT #10, KNTaxCust(1).Comment1;
   PRINT #10, KNTaxCust(1).Comment2;
   PRINT #10, KNTaxCust(1).Comment3;
   PRINT #10, KNTaxCust(1).LandValue;
   PRINT #10, KNTaxCust(1).BldgValue;
   PRINT #10, KNTaxCust(1).LandUseValue;
   PRINT #10, KNTaxCust(1).MPCurYrBillNumber;
   PRINT #10, KNTaxCust(1).MPCurYrTaxAmount
   
  
  
  END IF



  NEXT MCnt&
  CLOSE

  STOP
  END



END SUB

