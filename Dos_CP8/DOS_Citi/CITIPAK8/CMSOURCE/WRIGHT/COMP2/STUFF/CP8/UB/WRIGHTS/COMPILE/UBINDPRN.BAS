DEFINT A-Z
DECLARE FUNCTION ChkBillRecType% (FromFlag%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION nFindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION nGetAdjFactor# ()
DECLARE FUNCTION nGetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE FUNCTION nGetNumOfRevs% ()
DECLARE FUNCTION nGetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE FUNCTION ChkAndAskIfOk% (RecNo&)
DECLARE SUB SmallPause ()
DECLARE SUB PrintAlign1Mask ()
DECLARE SUB Post1BillTrans (VoidLastFlag%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB Print1Bill (FromFlag%)
DECLARE SUB Calc1Bill (RecNo&, FromFlag%)
DECLARE SUB VoidReprintMenu ()
DECLARE SUB IndBillPrintMenu ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB BlockClear ()
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetCustMeterType% (UBLocatRec() AS ANY, ThisMeter%)
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WazzWind (BYVAL a, BYVAL b, BYVAL c, BYVAL d, BYVAL E, BYVAL f, BYVAL g)
DECLARE SUB RestScrn (Array())
DECLARE SUB SaveScrn (Array())
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB UBLog (Text$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION Date2Num% (DateString$)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubpinfo.bi'
  '$INCLUDE: 'ubdraft.BI'

  STACK 5000
  
  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)
  MChoice$(1) = " Void & Reprint Last Bill "
  MChoice$(2) = " Print an Individual Bill "
  MChoice$(3) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 3, Col, MaxLen + 3, "Individual Bill Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      VoidReprintMenu
    CASE 2
      IndBillPrintMenu
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("ubmenu" + Ext$) THEN
    RUN "ubmenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB Calc1Bill (RecNo&, FromFlag%)
  
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
  
  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpLen
  TownName$ = UBSetUp(1).UTILNAME
  
  REDIM RevDesc(1 TO MaxRevsCnt) AS STRING * 12
  FOR Cnt = 1 TO MaxRevsCnt     'find last active revenue
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(Cnt).RevName)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = Cnt - 1       'set actual number of revenues
      EXIT FOR
    ELSE        'build revenue description lines
      LSET RevDesc(Cnt) = UCASE$(TempRev$)
      IF INSTR(RevDesc(Cnt), "ELECTRIC") THEN
        DoFuelAdjFlag = True
      END IF
    END IF
  NEXT
  
  IF DoFuelAdjFlag THEN
    FuelAdjAmt# = nGetAdjFactor#
  ELSE
    FuelAdjAmt# = 0
  END IF
  
  BlockClear
  ShowProcessingScrn "Calculating Utility Charges."
  QPrintRC "100", 13, 40, -1

  'find the electric revenue position
  FOR ThisRevCnt = 1 TO 15
    IF INSTR(UBSetUp(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
    END IF
  NEXT
  
  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRates = GetNumRateRecs%
  
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE
  
  IF Exist(UBIBillFile) THEN
    KILL UBIBillFile
  END IF
  
  BillFile = FREEFILE
  OPEN UBIBillFile FOR RANDOM SHARED AS BillFile LEN = UBBillRecLen
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  
  BillCnt = 0
  NumOfRevs = nGetNumOfRevs%
  
  'clears bill rec ready for next customer
  REDIM UBBillRec(1) AS UBTransRecType
  
  GET CustFile, RecNo&, UBCustRec(1)
  
  MeterConsp& = 0
  TMeterConsp& = 0
  
  FOR FRCnt = 1 TO 4
    WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
    IF UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 AND WhatService > 0 THEN
      '11/19/96 Fixed Rev. amt. to add to current rev amt
      'UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + UBCustRec(1).FlatRates(FRCnt).FRAmt)
      '01-09-97 Fixed Multiplier bug in flat rates
      Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
      IF Multi < 1 THEN Multi = 1
      FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
      UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
      UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FlatAmt#)
      IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
        UBBillRec(1).TaxAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
      END IF
    END IF
  NEXT
  
  '12-6-96  Monthly Billed amounts
  FOR MRCnt = 1 TO 2
    WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
    IF UBCustRec(1).Monthly(MRCnt).PayAmt > 0 AND WhatService > 0 THEN
      IF UBCustRec(1).Monthly(MRCnt).AMTOWED < UBCustRec(1).Monthly(MRCnt).PayAmt THEN
        HowMuch# = UBCustRec(1).Monthly(MRCnt).AMTOWED
      ELSE
        HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
      END IF
      UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + HowMuch#)
      UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + HowMuch#)
      IF UBSetUp(1).Revenues(WhatService).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
        TaxAmt# = Round#(HowMuch# * UBSetUp(1).Revenues(WhatService).TAXRATE)
        UBBillRec(1).TaxAmt(WhatService) = Round#(UBBillRec(1).TaxAmt(WhatService) + TaxAmt#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
      END IF
    END IF
  NEXT
  
  FOR RCnt = 1 TO NumOfRevs     'look at each rev line
    MeterConsp& = 0
    TMeterConsp& = 0
    WhatTbl = nFindRateTbl(UBCustRec(1).Serv(RCnt).RateCode, NumOfRates, UBRateTbls())
    
    IF WhatTbl THEN
      'if this is a non-metered service
      IF UBSetUp(1).Revenues(RCnt).UseMtr = "N" THEN
        '02-05-97 added fix add to current rev amt
        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + UBRateTbls(WhatTbl).MINAMT)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBRateTbls(WhatTbl).MINAMT)
        GOTO GotAmt
      END IF
      
      MeterType$ = UBCustRec(1).Serv(RCnt).RMTRTYPE
      
      FOR MCnt = 1 TO 7
        IF MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType THEN
          MeterLocNum = MCnt
          UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
          UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
          UBBillRec(1).MtrTypes(MCnt) = nGetCustMeterType(UBCustRec(), MCnt)
          'Found correct meter
          '052797 Added to stop overflow error.
          IF (UBCustRec(1).LocMeters(MCnt).CurRead < 0) OR (UBCustRec(1).LocMeters(MCnt).PrevRead < 0) THEN
            MeterConsp& = 0
          ELSE
            MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
          END IF
          
          IF MeterConsp& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
            MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
          END IF
          'Meter multi??
          IF UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 THEN
            MeterConsp& = MeterConsp& * UBCustRec(1).LocMeters(MCnt).MTRMulti
          END IF
          'PRINT UBCustRec(1).LocMeters(MCnt).NumUser
          'convert units here if necessary
          'To Convert From Cubic Feet to Gallons Mult by 7.481 Here
          IF (UBBillRec(1).MtrTypes(MCnt) = 1 OR UBBillRec(1).MtrTypes(MCnt) = 2 OR UBBillRec(1).MtrTypes(MCnt) = 3) AND UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
            MeterConsp& = MeterConsp& * 7.481
          END IF
          TMeterConsp& = TMeterConsp& + MeterConsp&
        END IF
      NEXT
      
      AddRevAmt# = 0
      
      IF UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 THEN
        'adjust min consumption for calc below
        NumUser& = UBCustRec(1).LocMeters(MeterLocNum).NumUser - 1
        AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
        MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
        TMeterConsp& = TMeterConsp& - MinimumConsp&
        IF (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 THEN
          '062697 Adjust the min consump test to actual NumUsers * MINUNITS
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
          GOTO GotAmt
        END IF
      ELSE
        NumUser& = 1
        IF TMeterConsp& <= UBRateTbls(WhatTbl).MINUNITS THEN
          'if we bill the minium
          RevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
          GOTO GotAmt
        END IF
      END IF
      
      '01-20-97 Add Fix for minium units test for equal to also.
      '04-23-1997 'Fixed to ADD TO TOTAL
      'AddRevAmt# = 0
      RevAmt# = nGetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
      RevAmt# = RevAmt# + AddRevAmt#
      UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
      IF RCnt = ElecRev THEN
        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + Round#(UBBillRec(1).RevAmt(RCnt) * FuelAdjAmt#))
      END IF
      UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
GotAmt:
      IF UBSetUp(1).Revenues(RCnt).TAXRATE > 0 AND UBCustRec(1).TAXEXPT <> "Y" THEN
        UBBillRec(1).TaxAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(RCnt))
      END IF
    END IF
  NEXT          'loop through all revenue sources
  
  IF UBCustRec(1).STATUS = "I" AND UBBillRec(1).TransAmt > 0 THEN
    UBBillRec(1).TransAmt = 0
    FOR RCnt = 1 TO NumOfRevs
      UBBillRec(1).RevAmt(RCnt) = 0
    NEXT
    UBBillRec(1).ActiveFlag = False
    BadCnt = BadCnt + 1
    UBBillRec(1).CustAcctNo = LCnt
  ELSE
    UBBillRec(1).ActiveFlag = False
  END IF
  
  IF UBBillRec(1).TransAmt > 0 THEN
    BillCnt = BillCnt + 1
    UBBillRec(1).ActiveFlag = True
    UBBillRec(1).CustAcctNo = LCnt
  END IF
  
MSkipem:
  UBBillRec(1).CustAcctNo = RecNo&
  UBBillRec(1).OperatorNumber = FromFlag
  PUT BillFile, 1, UBBillRec(1)
  
  CLOSE
  
  ERASE UBBillRec, UBCustRec, UBSetUp, UBRateTbls

  SmallPause
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

END SUB

FUNCTION ChkAndAskIfOk (RecNo&)

  REDIM TempScrn(0)

  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpLen
  'TownName$ = UBSetUp(1).UTILNAME
  
  TRevCnt = nGetNumOfRevs%
  REDIM RevText(1 TO 15)   AS STRING * 12
  FOR Cnt = 1 TO TRevCnt     'find last active revenue
    RevText(Cnt) = QPTrim$(UBSetUp(1).Revenues(Cnt).RevName)
  NEXT

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  GET #UBTran, UBCustRec(1).LastTrans, UBTransRec(1)
  CLOSE UBTran

  IF UBTransRec(1).TransType = TranUtilityBill THEN
    GOSUB ShowInfo
    Ok = MsgBox%("UBIBIL", "VOIDBTNS")
    SELECT CASE Ok
    CASE 1
      UBLog "ABORTED: VOID & PRINT:" + STR$(RecNo&)
      ChkAndAskIfOk = False
    CASE 2
      GOSUB DOVoidLastTrans
      ChkAndAskIfOk = True
    END SELECT
  ELSE
    SaveScrn TempScrn()
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "This Accounts last transaction is NOT", 10, 22, -1
    QPrintRC "a BILL transaction. Can't Void & Print.", 12, 21, -1
    WaitForAction
    RestScrn TempScrn()
    ChkAndAskIfOk = 0
  END IF
EXIT FUNCTION

DOVoidLastTrans:

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen
  GET #UBTran, UBCustRec(1).LastTrans, UBTransRec(1)
'  UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'  UBCustRec(1).CurrBalance = UBTransRec(1).TransAmt
'  UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  CLOSE UBTran

RETURN

ShowInfo:
  'DisplayUBScrn

  LibFile2Scrn "UBIBIL", "VOIDPRNT", MonoCode%, Attribute%, ErrorCode%

  QPrintRC UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB, 4, 27, -1
  QPrintRC UBCustRec(1).STATUS, 4, 48, -1
  QPrintRC STR$(RecNo&), 4, 63, -1
  QPrintRC UBCustRec(1).CustName, 5, 27, -1
  QPrintRC UBCustRec(1).ServAddr, 6, 27, -1
  QPrintRC Num2Date$(UBTransRec(1).BillDate), 9, 28, -1
  QPrintRC QPTrim$(FUsing$(STR$(UBTransRec(1).TransAmt), "$$#####.##")), 9, 60, -1

  FOR zz = 1 TO 5
    'QPrintRC QPTrim$(RevText(Cnt), , zz + 10, 22, -1
    QPrintRC RevText$(zz), zz + 10, 10, -1
    QPrintRC RevText$(zz + 5), zz + 10, 32, -1
    QPrintRC RevText$(zz + 10), zz + 10, 55, -1
    QPrintRC FUsing$(STR$(UBTransRec(1).RevAmt(zz)), "$$#####.##"), zz + 10, 18, -1
    QPrintRC FUsing$(STR$(UBTransRec(1).RevAmt(zz + 5)), "$$#####.##"), zz + 10, 40, -1
    QPrintRC FUsing$(STR$(UBTransRec(1).RevAmt(zz + 10)), "$$#####.##"), zz + 10, 63, -1

  NEXT
RETURN
END FUNCTION

FUNCTION ChkBillRecType% (FromFlag%)

  ChkBillRecType% = True      'assume the user isn't an idiot.

  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))

  UBBill = FREEFILE
  OPEN UBIBillFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  GET UBBill, 1, UBBillRec(1)
  CLOSE UBBill

  IF UBBillRec(1).OperatorNumber <> FromFlag THEN
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    SELECT CASE FromFlag
    CASE 1
      QPrintRC "A Bill is in Process in Void & Reprint", 10, 22, -1
      QPrintRC "Use: Select Customer for Bill Print.", 12, 23, -1
    CASE 2
      QPrintRC "A Bill is in Process in Individual Print", 10, 21, -1
      QPrintRC "Use: Select Customer for Reprint.", 12, 24, -1
    CASE ELSE

    END SELECT
    ChkBillRecType% = False   'nope they are!
    WaitForAction
  END IF

END FUNCTION

SUB IndBillPrintMenu

  FromFlag = 1

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  MChoice$(1) = " Select Customer for Bill Print "
  MChoice$(2) = " Print Individual Utility Bill"
  MChoice$(3) = " Post Individual Utility Bill"
  MChoice$(4) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 3, Col, MaxLen + 3, "Individual Print Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      BlockClear
      LookUp RecNo&, "Individual Print", 2, True, True
      IF RecNo& > 0 THEN
        Calc1Bill RecNo&, FromFlag
      END IF
    CASE 2
      'PrintUtilBills
      Print1Bill 1
    CASE 3
      Post1BillTrans False
    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  EXIT SUB
  
END SUB

FUNCTION nFindRateTbl% (RateCode$, NumOfRates, UBRateTbls() AS UBRateTblRecType)
  nFindRateTbl = 0
  RateCode$ = QPTrim$(RateCode$)
  IF LEN(RateCode$) THEN        'if this rev has a rate code
    FOR TCnt = 1 TO NumOfRates  'find the right one
      IF RateCode$ = QPTrim$(UBRateTbls(TCnt).RateCode) THEN
        nFindRateTbl = TCnt
        EXIT FOR
      END IF
    NEXT
  END IF

END FUNCTION

FUNCTION nGetAdjFactor#

  LibName$ = "UB"
  ScrnName$ = "FUELADJ"

  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0

  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True

  '--Set screen number to one and display screen

  BlockClear

  DisplayUBScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      AdjFactor# = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      AdjFactor# = -10000
      ExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF

  LOOP UNTIL ExitFlag

  nGetAdjFactor# = AdjFactor#
  ERASE Frm, Form$, Fld

END FUNCTION

FUNCTION nGetCustMeterType% (UBCustRec() AS NewUBCustRecType, ThisMeter)

  LMtrType$ = QPTrim$(UBCustRec(1).LocMeters(ThisMeter).MTRType)
  LMtrTypeLen = LEN(LMtrType$)
  IF LMtrTypeLen > 0 THEN
    SELECT CASE LMtrType$
    CASE "W"
      LThisMeter = MtrWaterOnly
    CASE "S"
      LThisMeter = MtrSewerOnly
    CASE "C"
      LThisMeter = MtrCombined
    CASE "E"
      LThisMeter = MtrElectric
    CASE "D"
      LThisMeter = MtrDemand
    CASE "G"
      LThisMeter = MtrGas
    CASE "T"
      LThisMeter = MtrTouchRead
    CASE ELSE
      LThisMeter = True
    END SELECT
    nGetCustMeterType = LThisMeter
  ELSE
    nGetCustMeterType = 0
  END IF


END FUNCTION

FUNCTION nGetNumOfRevs%
  NumOfRevs = 15
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  FOR RevCnt = 1 TO 15
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  nGetNumOfRevs = NumOfRevs
  ERASE UBSetUpRec

END FUNCTION

FUNCTION nGetRevCharge# (RateTbl AS UBRateTblRecType, TMeterConsp&, MeterMulti&)

  MinBillAmt# = RateTbl.MINAMT

  LastTblCnt = 10
  FOR BCnt = 1 TO 10
    IF RateTbl.TblBreaks(BCnt).UNITAMT <= 0 THEN
      LastTblCnt = BCnt - 1
      EXIT FOR
    END IF
  NEXT

  MeterConsump& = TMeterConsp&

  TAmt# = 0

  IF LastTblCnt >= 2 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(1).UNITS AND MeterConsump& < RateTbl.TblBreaks(2).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    END IF
  ELSE          'no other rate breaks
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    GOTO GotTAmt
  END IF

  'Break 2
  IF LastTblCnt >= 3 THEN
    IF MeterConsump& > RateTbl.TblBreaks(2).UNITS AND MeterConsump& < RateTbl.TblBreaks(3).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    GOTO GotTAmt
  END IF

  'Break 3
  IF LastTblCnt >= 4 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(3).UNITS AND MeterConsump& < RateTbl.TblBreaks(4).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    GOTO GotTAmt
  END IF

  'Break 4
  IF LastTblCnt >= 5 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(4).UNITS AND MeterConsump& < RateTbl.TblBreaks(5).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    GOTO GotTAmt
  END IF

  'break 5
  IF LastTblCnt >= 6 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(5).UNITS AND MeterConsump& < RateTbl.TblBreaks(6).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    GOTO GotTAmt
  END IF

  'break 6
  IF LastTblCnt >= 7 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(6).UNITS AND MeterConsump& < RateTbl.TblBreaks(7).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GOTO GotTAmt
  END IF

  'break 7
  IF LastTblCnt >= 8 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(7).UNITS AND MeterConsump& < RateTbl.TblBreaks(8).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GOTO GotTAmt
  END IF

  'break 8
  IF LastTblCnt >= 9 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(8).UNITS AND MeterConsump& < RateTbl.TblBreaks(9).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GOTO GotTAmt
  END IF

  'break 9
  IF LastTblCnt >= 10 THEN
    IF MeterConsump& >= RateTbl.TblBreaks(9).UNITS AND MeterConsump& < RateTbl.TblBreaks(10).UNITS THEN
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GOTO GotTAmt
    ELSE
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    END IF
  ELSE
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    GOTO GotTAmt
  END IF

GotTAmt:
  nGetRevCharge# = Round#(MinBillAmt# + TAmt#)


END FUNCTION

SUB Post1BillTrans (VoidFlag%)

  UBLog "IN: 1 Bill Posting."

  IF NOT Exist(UBIBillFile) THEN
    UBLog "ERROR: " + UBIBillFile + " Calculation file NOT FOUND!"
    CursorOff
    BlockClear
    DisplayUBScrn "NON2POST"
    WaitForAction
    GOTO ExitBillPost:
  END IF

  IF NOT Exist("UBIBILL.PRN") THEN
    UBLog "ERROR: UBIBILL.PRN Print File NOT FOUND!"
    CursorOff
    BlockClear
    DisplayUBScrn "NOTPRNTD"
    WaitForAction
    GOTO ExitBillPost:
  END IF

  LibName$ = "UB"
  ScrnName$ = "PSTBILLS"

  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0

  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True

  '--Set screen number to one and display screen

  BlockClear

  DisplayUBScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      OKFlag = True
    CASE EscKey
      ExitFlag = True
    END SELECT

  LOOP UNTIL OKFlag OR ExitFlag

  IF ExitFlag THEN
    BlockClear
    CursorOff
    DisplayUBScrn "POSTCAN"
    WaitForAction
    UBLog "ABORTED:  Bill Posting"
    GOTO ExitBillPost
  END IF

  UBLog "START: 1 Post Transaction."

  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType

  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBBill = FREEFILE
  OPEN UBIBillFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen

  ShowProcessingScrn "Posting Bill Transaction."
  QPrintRC "100", 13, 40, -1

  GET UBBill, 1, UBBillRec(1)

  CustAcct& = UBBillRec(1).CustAcctNo
  GET UBCust, CustAcct&, UBCustRec(1)

'070297 Added void last trans process
  'IF VoidFlag THEN
  '  GOSUB DOVoidLastTrans
  'END IF

  PostedCnt& = PostedCnt& + 1

  '06-25-97 added removal of the nonrecurring flat rates
  FRFlag = False
  FOR FRCnt = 1 TO 4 'Remove non-recurring flat rates
    IF UBCustRec(1).FlatRates(FRCnt).FRFREQ = "N" THEN
      UBCustRec(1).FlatRates(FRCnt).FRDESC = ""
      UBCustRec(1).FlatRates(FRCnt).FRAMT = 0
      UBCustRec(1).FlatRates(FRCnt).FRFREQ = ""
      UBCustRec(1).FlatRates(FRCnt).REVSRC = 0
      UBCustRec(1).FlatRates(FRCnt).NumMin = 0
      FRFlag = True
    END IF
  NEXT
  IF FRFlag THEN
    UBLog "1 BILL POST: Removed Flat Rate. Acct:" + STR$(BillCnt)
  END IF
  UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
  UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  FOR RevCnt = 1 TO MaxRevsCnt
    UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
  NEXT
  UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
  UBBillRec(1).TransDesc = "Utility Billing"
  UBBillRec(1).TransDate = UBBillRec(1).BillDate
  FOR MtrCnt = 1 TO 7
    IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
      ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
      IF ReadAmt& < 0 THEN  'Meter rolled over or, been misread
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
        ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
      END IF
      UBCustRec(1).LocMeters(MtrCnt).AvgUse = Round#(UBCustRec(1).LocMeters(MtrCnt).AvgUse + ReadAmt&)
      UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
      UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
    END IF
  NEXT
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBBillRec(1).PrevTrans = PrevLastTrans&
  NumOfTranRecs& = (LOF(UBTran) \ UBBillRecLen) + 1
  'point to next trans to write
  UBCustRec(1).LastTrans = NumOfTranRecs&

  PUT UBCust, CustAcct&, UBCustRec(1)
  PUT UBTran, NumOfTranRecs&, UBBillRec(1)

  CLOSE
  UBLog "FINISH: 1 Post Transaction."
  UBLog "POSTED:" + STR$(PostedCnt&) + " New BILL Transactions."

  KILL UBIBillFile
  KILL "UBIBILL.PRN"

  UBLog "KILLED: " + UBIBillFile + " & UBIBILL.PRN"

  SmallPause
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitBillPost:
  UBLog "OUT: 1 Bill Posting." + CrLf$
'EXIT SUB

END SUB

SUB Print1Bill (FromFlag%)

  LPIFlag = False

  IF NOT Exist(UBIBillFile) THEN
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "NO Customer has been selected.", 10, 26, -1
    QPrintRC "Use: Select a Customer for Bill Print.", 12, 22, -1
    WaitForAction
    GOTO ExitPrintBill
  ELSE
    IF NOT ChkBillRecType%(FromFlag%) THEN
      GOTO ExitPrintBill
    END IF
  END IF

  REDIM ScrnArray(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file

  'Section to check for customer modifications

  TownName$ = UBSetUpRec(1).UTILNAME

  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF

  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF

  LibName$ = "UB"

  ScrnName$ = "PRNIBILL"

  MActionRow = 19

  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0

  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  FrstBFld = FldNum%("FRSTBILL", Fld())
  LastBFld = FldNum%("LASTBILL", Fld())

  BillDFld = FldNum%("BILLDATE", Fld())
  PastDFld = FldNum%("PASTDATE", Fld())

  PRDateFld = FldNum%("PRDATE", Fld())
  CRDateFld = FldNum%("CRDATE", Fld())

  DraftDFld = FldNum%("DRFTDATE", Fld())
  MesgFld = FldNum%("MSGLINE1", Fld())

  Desc1Fld = FldNum%("MSGLINE2", Fld())
  Desc2Fld = FldNum%("MSGLINE3", Fld())
  Desc3Fld = FldNum%("MSGLINE4", Fld())

  IF IndianFlag THEN
    Fld(Desc1Fld).Protected = False
    Fld(Desc2Fld).Protected = False
    Fld(Desc3Fld).Protected = False
  END IF

  Today$ = DATE$

  Action = 1
  Frm(1).StayOnField = True

  FirstTime = True
  '--Set screen number to one and display screen

  BlockClear

  DisplayUBScrn ScrnName$

  IF IndianFlag THEN
    FOR Cnt = 1 TO 3
      QPrintRC "Description " + QPTrim$(STR$(Cnt)) + ":", Cnt + 15, 21, -1
    NEXT
  END IF

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(BillDFld, 0) = Today$
      LSET Form$(PastDFld, 0) = Today$
      LSET Form$(DraftDFld, 0) = Today$
      Action = 2
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE F5KEY
      PrintAlign1Mask
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 18 TO 30           '--Cancel button
          PressButton 27, MActionRow, 18, 30
        CASE 31 TO 46           '--F5 button
          PressButton F5KEY, MActionRow, 31, 46
        CASE 47 TO 64           '--F10 button
          PressButton F10Key, MActionRow, 47, 64
        END SELECT
      END SELECT                'row
    END IF

    '--Check screen page
  LOOP UNTIL ExitFlag

  IF AbortFlag GOTO ExitPrintBill

  PastDate$ = Form$(PastDFld, 0)

  'do bill printing here
  '**************************************************************************

  REDIM BillInfoRec(1) AS PrintBillInfoType

  BillInfoRec(1).MsgLine2 = Form$(Desc1Fld, 0)
  BillInfoRec(1).MsgLine3 = Form$(Desc2Fld, 0)
  BillInfoRec(1).MsgLine4 = Form$(Desc3Fld, 0)

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBDraftPayRec(1) AS UBDraftPayRecType
  UBDraftPayLen = LEN(UBDraftPayRec(1))

  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))

  UBBill = FREEFILE
  OPEN UBIBillFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBIBILL.PRN" FOR OUTPUT AS UBRpt

  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF

  ShowProcessingScrn "Creating Utility Bills."

  '-----------------------------------------
  PrintedCnt = 0

  GET UBBill, 1, UBBillRec(1)
  CustAcctNo& = UBBillRec(1).CustAcctNo

  GET UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)

  Num2Print = UBCustRec(1).BILLCOPY
  IF Num2Print < 1 THEN Num2Print = 1

  IF UBBillRec(1).ActiveFlag THEN
    IF UBBillRec(1).TransAmt <> 0 OR UBCustRec(1).CurrBalance <> 0 OR UBCustRec(1).PrevBalance THEN
      PrintedCnt = PrintedCnt + 1
      UBBillRec(1).BillNumber = PrintedCnt
      UBBillRec(1).TransDate = BillDate
      UBBillRec(1).TransDesc = "UTILITY BILL"

      FOR MtrCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MtrCnt).CurDate > 0 THEN
          UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
          UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
          EXIT FOR
        END IF
      NEXT

      IF UBBillRec(1).ReadDate <= 0 THEN
        UBBillRec(1).ReadDate = BillDate - 30
      END IF
      IF UBBillRec(1).PrevDate <= 0 THEN
        UBBillRec(1).PrevDate = UBBillRec(1).ReadDate - 30
      END IF

      IF UseEDateFlag THEN
        UBBillRec(1).PrevDate = PRDate
        UBBillRec(1).ReadDate = CRDate
      END IF

      '*****************
      DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate

      UBBillRec(1).BillDate = BillDate
      UBBillRec(1).PastDueDate = PastDate
      UBBillRec(1).DraftDate = DraftDate
      UBBillRec(1).BillMsg = Message$

      'these are for reprinting bills
      'UBBillRec(1)CustLocation = CustAcctNo&
      UBBillRec(1).CustAcctNo = CustAcctNo&

      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)

      TotalTax# = 0
      FOR TaxCnt = 1 TO MaxRevsCnt - 1
        TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
      NEXT

      PUT UBBill, 1, UBBillRec(1)

      DidADraftFlag = False
      IF UseDraftFlag AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag THEN
        UBDraftPayRec(1).CustAcctNum = CustAcctNo&
        UBDraftPayRec(1).DraftAmt = UBBillRec(1).TransAmt
        PUT UBDraft, , UBDraftPayRec(1)
        DidADraftFlag = True
      END IF

      'Custom Mod Here For Lilesville, NC
      IF Lilesville > 0 THEN
        IF UBCustRec(1).Serv(1).RateCode = "WIN " OR UBCustRec(1).Serv(1).RateCode = "WOUT" THEN
          TenPercentAmount# = (UBBillRec(1).TransAmt - UBBillRec(1).RevAmt(1)) + (UBBillRec(1).RevAmt(1) * 1.1111)
        ELSE
          TenPercentAmount# = UBBillRec(1).TransAmt
        END IF
      END IF
      'End Lilesville Custom Mod

      FOR BillCopies = 1 TO Num2Print
        ''$INCLUDE: 'UBSTDBIL.BI'        'Std  Lilesville
        '$INCLUDE: 'UBLANDIS.BI'  'landis
        ''$INCLUDE: 'UBLONGVW.BI'  'Longview
        ''$INCLUDE: 'UBHOLLY.BI'   '
        ''$INCLUDE: 'UBTROUTM.BI'   'Troutman Format
        ''$INCLUDE: 'UBMCMICK.BI'  'Mccormick
        ''$INCLUDE: 'UBOLDSDS.BI'  'Cleveland NC
        ''$INCLUDE: 'UBOLDSTD.BI'  'Old citipak standard bill
        ''$INCLUDE: 'UBCONCRD.BI'  'Concord
        ''$INCLUDE: 'UBELLEN.BI'   'Ellenboro NC
        ''$INCLUDE: 'UBGROVER.BI'  'Grover NC
        ''$INCLUDE: 'UBGILES.BI'   'Giles Cty Va
        ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
        ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
        ''$INCLUDE: 'UBGORDON.BI'        'Gordonsville VA
        ''$INCLUDE: 'UBTUSK.BI'          'Tuskaseigee NC
        ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
        ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
      NEXT
    END IF
  END IF

SkipEm:
  
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF

  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec

  IF NOT AbortFlag THEN
    PrintRptFile "Individual Bill Printing", "UBIBILL.PRN", 1, RetCode, 0
  END IF

'**************************************************************************
'end bill printing
  GOTO ExitPrintBill:

CheckReqFields:

  FBillNO& = VAL(Form$(FrstBFld, 0))
  LBillNO& = VAL(Form$(LastBFld, 0))
  BillDate = Date2Num(Form$(BillDFld, 0))

  BillDate$ = Num2Date$(BillDate)

  PastDate = Date2Num(Form$(PastDFld, 0))

  DraftDate = Date2Num(Form$(DraftDFld, 0))
  DueDate$ = Form$(PastDFld, 0)

  Message$ = Form$(MesgFld, 0)

  PRDate = Date2Num(Form$(PRDateFld, 0))
  CRDate = Date2Num(Form$(CRDateFld, 0))

  IF (CRDate > 0) AND (PRDate > 0) THEN
    UseEDateFlag = True
  ELSE
    UseEDateFlag = False
  END IF

  IF BillDate = -32768 THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "BADBDATE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = BillDFld
  ELSEIF PastDate < BillDate THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "BADBDATE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = PastDFld
  ELSEIF (UseDraftFlag AND DraftDate = -32768) OR (UseDraftFlag AND DraftDate < BillDate) THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "BADDDATE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 3
  ELSE
    ReqFldsOK = True
  END IF

  IF UseDraftFlag THEN
    DraftDate$ = Form$(DraftDFld, 0)
    DFFileName$ = "DF" + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 4, 2) + RIGHT$(DraftDate$, 2) + ".DAT"
  END IF

  RETURN

ExitPrintBill:

END SUB

SUB PrintAlign1Mask
  CursorOff
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  PrintRptFile "Utility Bill Alignment Mask ", "UBBLMASK.DAT", 1, RetCode, 4
  RestScrn ScrnArray()
  ERASE ScrnArray
END SUB

SUB SmallPause
  st! = TIMER
  DO: LOOP UNTIL TIMER > st! + .9
END SUB

SUB VoidReprintMenu

  FromFlag = 2

  '--Dim the choice array to the number of menu items

  REDIM MChoice$(1 TO 4)
  MChoice$(1) = " Select Customer for Reprint "
  MChoice$(2) = " Reprint New Utility Bill"
  MChoice$(3) = " Post New Utility Bill"
  MChoice$(4) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 3, Col, MaxLen + 3, "Void & Reprint Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      BlockClear
      LookUp RecNo&, "Void & Individual Print", 2, True, True
      IF RecNo& > 0 THEN
        IF ChkAndAskIfOk(RecNo&) THEN
          Calc1Bill RecNo&, FromFlag
        END IF
      END IF
    CASE 2
      Print1Bill 2
    CASE 3

    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
END SUB

