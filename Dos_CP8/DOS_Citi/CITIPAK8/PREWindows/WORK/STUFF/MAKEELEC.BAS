DEFINT A-Z
DECLARE SUB ExportConsumptionInformation ()
DECLARE SUB ExportCustomerInformation ()
DECLARE SUB CustCutOffListing ()
DECLARE SUB CustomerFlatRates ()

DECLARE SUB WestCaroConsumpList ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)

DECLARE SUB SortServAddress (IndexText$)
DECLARE SUB CustEditMeterEST (CustRec&)
DECLARE SUB CustByRate ()
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB CustWorkOrders (CustRec&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PaymentSumReport ()
DECLARE SUB MailingLabel ()
DECLARE SUB AgedBalanceListing ()
DECLARE SUB CustomerCutOffListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB DetailedTransJournal ()
DECLARE SUB MasterBalanceListing ()
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetCursor (x, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE FUNCTION Chk4DupeBookSeqNum (BOOK$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakePostalIndex (IdxTypeText$)
DECLARE FUNCTION QPStrL$ (LongNumber&)
DECLARE FUNCTION WEnvTest% ()
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'
   
      ExportConsumptionInformation

SUB ExportConsumptionInformation
  q$ = CHR$(34)
  c$ = ","

  MaxTRS = 12

  LastGood = Date2Num(DATE$)

  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True

  Dash80$ = STRING$(80, "-")

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBSetUpRec(1)  AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen%
  FOR Rev = 1 TO 15
    IF INSTR(UBSetUpRec(1).Revenues(Rev).RevName, "ELECTRIC") > 0 THEN
      ElecRev = Rev
      EXIT FOR
    END IF
  NEXT

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs
  
  UBRpt = FREEFILE
  OPEN "CONSUMP.ASC" FOR OUTPUT AS UBRpt
  PRINT #UBRpt, q$; "ACCT"; q$; c$; q$; "CUSTNAME"; q$; c$; q$; "SERVADDR";
  FOR zz = MaxTRS TO 1 STEP -1
    zz$ = QPTrim$(STR$(zz))
    PRINT #UBRpt, q$; c$; q$; "TRDATE"; zz$; q$; c$; q$; "CREAD"; zz$; q$; c$; q$; "PREAD"; zz$; q$; c$; q$; "CON"; zz$; ' q$; c$; q$; "TRANAMT"; zz$;
  NEXT
  PRINT #UBRpt, q$;
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfCust = LOF(UBCust) \ UBCustRecLen
  
  BlockClear
  ShowProcessingScrn "Customer Export."
  FOR CCnt = 1 TO NumOfCust
    RecNo& = CCnt    'IdxBuff(CCnt).RecNum
    GET #UBCust, RecNo&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      DidCnt = 0
      ThisTrans& = UBCustRec(1).LastTrans
      DO WHILE ThisTrans& > 0
        GET #UBTran, ThisTrans&, UBTranRec(1)
        IF UBTranRec(1).TransType = TranUtilityBill AND UBTranRec(1).TransDate <= LastGood THEN
          GOSUB PrintConsDetail
          IF DidCnt = MaxTRS THEN
            EXIT DO
          END IF
        END IF
        ThisTrans& = UBTranRec(1).PrevTrans
      LOOP
      IF DidCnt > 0 THEN
        IF DidCnt < MaxTRS THEN
          FOR zz = DidCnt + 1 TO MaxTRS
             PRINT #UBRpt, c$; q$; ""; q$; c$; q$; ""; q$; c$; q$; ""; q$; c$; q$; ""; q$; c$; q$; ""; q$;
          NEXT
        END IF
      END IF
    END IF
    ShowPctComp CCnt, NumOfRecs

    IF ExitFlag THEN
      EXIT FOR
    END IF
  NEXT

  CLOSE

  'COLOR 0, 7
  FOR ll = 8 TO 15
    LOCATE ll, 22
    PRINT STRING$(38, 32);
  NEXT ll
  LOCATE 8, 23
  LOCATE 11, 28
  PRINT "File Created is CONSUMP.ASC";
  LOCATE 13, 29
  PRINT "Press any key to Continue."

'  LPRINT "Map of Consumption Export file: UBCONSMP.TXT"
'  LPRINT
'  LPRINT "Field Listing:"
'  LPRINT
'  LPRINT "Account Number"
'  LPRINT "Customer Name"
'  LPRINT "Past 12 months will loop showing:"
'  LPRINT "Billing Date"
'  LPRINT "Current Read"
'  LPRINT "Previous Read"
'  LPRINT "Consumption"
'  LPRINT "Bill Amount"
'  LPRINT CHR$(12);
  WaitForAction


ExitConsumpHist:

EXIT SUB

PrintConsDetail:
  DidAMeter = False
  FOR MtrCnt = 1 TO 7
    SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
    CASE 1, 2, 3, 7
      DidAMeter = True
      IF DidCnt = 0 THEN
        GOSUB PrintCustInfo
      END IF
      GOSUB PrintThisMeter
      EXIT FOR
    END SELECT
  NEXT
  'IF NOT DidAMeter THEN
  '  MeterType$ = "        "
  '  MtrCnt = 1
  '  GOSUB PrintThisMeter
  'END IF

RETURN

PrintThisMeter:
  DidCnt = DidCnt + 1
  PRINT #UBRpt, c$; q$; Num2Date(UBTranRec(1).TransDate); q$; c$; q$;
  PRINT #UBRpt, QPTrim$(STR$(UBTranRec(1).CurRead(MtrCnt))); q$; c$; q$;
  PRINT #UBRpt, QPTrim$(STR$(UBTranRec(1).PrevRead(MtrCnt))); q$; c$; q$;
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  IF MeterConsp& < 0 THEN
    MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  END IF
  'IF DidCnt = 12 THEN
    PRINT #UBRpt, QPTrim$(STR$(MeterConsp&)); q$;
  'ELSE
  '  PRINT #UBRpt, QPTrim$(STR$(MeterConsp&)); q$; c$; q$;
  'END IF

  'PRINT #UBRpt, QPTrim$(FUsing$(STR$(UBTranRec(1).RevAmt(1)), "######.##")); q$;

  TotalConsump& = TotalConsump& + MeterConsp&

RETURN

PrintCustInfo:
  'IF CCCnt > 0 THEN
    PRINT #UBRpt,
  'END IF
  DO
    QPos = INSTR(UBCustRec(1).CustName, q$)
    IF QPos > 0 THEN
      MID$(UBCustRec(1).CustName, QPos, 1) = " "
    END IF
  LOOP WHILE QPos > 0

  DO
    CPos = INSTR(UBCustRec(1).CustName, c$)
    IF CPos > 0 THEN
      MID$(UBCustRec(1).CustName, CPos, 1) = " "
    END IF
  LOOP WHILE CPos > 0

  'PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
  RateCode$ = QPTrim$(UBCustRec(1).ServAddr)   'Serv(1).RateCode)

  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$; c$; q$; RateCode$; q$;

  'PRINT #UBRpt, UBCustRec(1).CustName
  'CCCnt = CCCnt + 1
  'IF CCCnt > 99 THEN
  '  ExitFlag = True
  'END IF
RETURN
END SUB

