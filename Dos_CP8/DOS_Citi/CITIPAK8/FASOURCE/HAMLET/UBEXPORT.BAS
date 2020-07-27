DEFINT A-Z
DECLARE SUB BuckSportReadingExport ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (LongNumber&)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB AgedBalanceListing ()
DECLARE SUB BlockClear ()
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB ExportConsumptionInformation ()
DECLARE SUB ExportCustomerInformation ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB GetCursor (x, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MailingLabel ()
DECLARE SUB MakePostalIndex (IdxTypeText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MasterBalanceListing ()
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PaymentSumReport ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (array())
DECLARE SUB SaveScrn (array())
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB SortServAddress (IndexText$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB WestCaroConsumpList ()

  
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

TYPE SReadType
   Book        AS STRING * 2
   Seq         AS STRING * 6
   CustName    AS STRING * 35
   CurrRead    AS STRING * 10
   ReadDate    AS STRING * 10
   CrLf        AS STRING * 2
END TYPE


  REDIM MChoice$(1 TO 4)

  MChoice$(1) = " Export Customer Information "
  MChoice$(2) = " Export Customer Consumption "
  MChoice$(3) = " Export 'S' Reading Info "
  MChoice$(4) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
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
  Col = ((80 - MaxLen) \ 2)
  'Help$ = "Utility Billing Main Menu"

  UsingBook = False
  UsingAcct = False
  UsingName = False

  AbortFlag = False

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 3, Col, MaxLen + 3, "Master Customer Export ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF

    SELECT CASE Choice
    CASE 0
      ExitFlag = True
    CASE 1
      ExportCustomerInformation
    CASE 2
      ExportConsumptionInformation
    CASE 3
      BuckSportReadingExport

    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT

  LOOP UNTIL OKFlag OR ExitFlag

  'IF ExitFlag THEN GOTO ExitExport
'

  IF INSTR(COMMAND$, "TEST") = 0 THEN
    RUN "ubmisc"
  END IF

  HideCursor

  ClearScrn

  END

SUB BuckSportReadingExport

  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True
  CrLf$ = CHR$(13) + CHR$(10)

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM SExport(1)   AS SReadType
  SRecLen = LEN(SExport(1))
  
  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  KillFile "UBSREAD.TXT"

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBSREAD.TXT" FOR RANDOM AS UBRpt LEN = SRecLen

  BlockClear
  ShowProcessingScrn "'S' Reading Export."

  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 AND UBCustRec(1).Status = "A" THEN
      UCode$ = UCASE$(QPTrim$(UBCustRec(1).UserCode1))
      IF UCode$ = "S" THEN
        REDIM SExport(1)   AS SReadType
        SExport(1).CrLf$ = CrLf$
        SExport(1).Book = UBCustRec(1).Book
        SExport(1).Seq = UBCustRec(1).SeqNumb
        LSET SExport(1).CustName = QPTrim$(UBCustRec(1).CustName)
        LSET SExport(1).CurrRead = QPTrim$(STR$(UBCustRec(1).LocMeters(1).CurRead))
        SExport(1).ReadDate = Num2Date$(UBCustRec(1).LocMeters(1).CurDate)
        PUT #UBRpt, , SExport(1)
      END IF
    END IF

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF

    ShowPctComp Cnt, NumOfRecs
  NEXT

  CLOSE UBCust, UBRpt

  ERASE IdxBuff, UBCustRec

   FOR ll = 8 TO 15
     LOCATE ll, 22
     PRINT STRING$(38, 32);
   NEXT ll
   LOCATE 8, 23
   LOCATE 11, 28
   PRINT "File Created is UBSREAD.TXT";
   LOCATE 13, 29
   PRINT "Press any key to Continue."
   WaitForAction

END SUB

SUB ExportConsumptionInformation

  REDIM TempScrn(0)

  q$ = CHR$(34)
  c$ = ","

  LibName$ = "UB"
  ScrnName$ = "UBEXPCON"
  NumScrns = 1

  '--define the multi-choice fields
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT

  '--Set screen number to one and display screen
  Scr = 1
  BlockClear
  DisplayUBScrn ScrnName$

  ShowCursor

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB ExpCheckDate
      IF OKFlag THEN
        EXIT DO
      END IF
'here
    CASE EscKey
      ExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF

    '--Check screen page

  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag

  IF ExitFlag THEN
    GOTO ExitConsumpHist
  END IF

  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True

  Dash80$ = STRING$(80, "-")

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs
  
  UBRpt = FREEFILE
  OPEN "UBCONSMP.TXT" FOR OUTPUT AS UBRpt
  PRINT #UBRpt, q$; "ACCT"; q$; c$; q$; "LOCATION"; q$; c$; q$; "CUSTNAME";
  PRINT #UBRpt, q$; c$; q$; "CUSTTYPE"; q$; c$; q$; "ADDR1"; q$; c$; q$; "ADDR2"; q$; c$; q$; "CITY"; q$; c$; q$; "STATE"; q$; c$; q$; "ZIP"; q$; c$; q$; "SERVADDR";
  FOR zz = NumOfPeriods TO 1 STEP -1
    zz$ = QPTrim$(STR$(zz))
    PRINT #UBRpt, q$; c$; q$; "TRDATE"; zz$; q$; c$; q$; "CURRREAD"; zz$; q$; c$; q$; "PREVREAD"; zz$; q$; c$; q$; "CONSUMP"; zz$; q$; c$; q$; "TRANAMT"; zz$;
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
        IF UBTranRec(1).TransType = TranUtilityBill THEN
          IF UBTranRec(1).TransDate >= FromDate AND UBTranRec(1).TransDate <= ThruDate THEN
            IF DidCnt = 0 THEN
              GOSUB PrintCustInfo
            END IF
            GOSUB PrintConsDetail
            DidCnt = DidCnt + 1
            IF DidCnt = NumOfPeriods THEN
              EXIT DO
            END IF
          END IF
        END IF
        ThisTrans& = UBTranRec(1).PrevTrans
      LOOP
      IF DidCnt > 0 THEN
        IF DidCnt < NumOfPeriods THEN
          FOR zz = DidCnt + 1 TO NumOfPeriods
             PRINT #UBRpt, c$; q$; "01-01-1980"; q$; c$; q$; "0"; q$; c$; q$; "0"; q$; c$; q$; "0"; q$; c$; q$; "0.00"; q$;
          NEXT
        END IF
      END IF
    END IF
    ShowPctComp CCnt, NumOfRecs
    'IF CCnt > 149 THEN EXIT FOR
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
  PRINT "File Created is UBCONSMP.TXT";
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

ExpCheckDate:
  
  FromDate = Date2Num(Form$(1, 0))
  ThruDate = Date2Num(Form$(2, 0))
  IF (FromDate < 0) OR (ThruDate < 0) OR (FromDate > ThruDate) THEN
    SaveScrn TempScrn()
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Date Range!", 10, 31, -1
    QPrintRC "Press any key to continue.", 12, 28, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 1
  ELSE
    FYear = VAL(RIGHT$(Form$(1, 0), 4))
    TYear = VAL(RIGHT$(Form$(2, 0), 4))
    FMonth = VAL(LEFT$(Form$(1, 0), 2))
    TMonth = VAL(LEFT$(Form$(2, 0), 2))
    IF FYear = TYear THEN
      IF FMonth = TMonth THEN
        NumOfPeriods = 1
        OKFlag = True
      ELSE
        NumOfPeriods = (TMonth - FMonth) + 1
        OKFlag = True
      END IF
    ELSE
      FMCnt = (12 - FMonth) + 1
      NumOfPeriods = FMCnt + TMonth
      IF TYear - FYear > 1 THEN
        NumOfPeriods = NumOfPeriods + 12
      END IF
      OKFlag = True
    END IF
  END IF
ExpDateRet:
RETURN

PrintConsDetail:
  DidAMeter = False
  FOR MtrCnt = 1 TO 7
    IF UBTranRec(1).MtrTypes(MtrCnt) > 0 THEN
      DidAMeter = True
      GOSUB PrintThisMeter
    END IF
  NEXT
  IF NOT DidAMeter THEN
    MeterType$ = "        "
    MtrCnt = 1
    GOSUB PrintThisMeter
  END IF
RETURN

PrintThisMeter:
  PRINT #UBRpt, c$; q$; Num2Date(UBTranRec(1).TransDate); q$; c$; q$;
  PRINT #UBRpt, QPTrim$(STR$(UBTranRec(1).CurRead(MtrCnt))); q$; c$; q$;
  PRINT #UBRpt, QPTrim$(STR$(UBTranRec(1).PrevRead(MtrCnt))); q$; c$; q$;
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  IF MeterConsp& < 0 THEN
    MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  END IF
  PRINT #UBRpt, QPTrim$(STR$(MeterConsp&)); q$; c$; q$;
  PRINT #UBRpt, QPTrim$(FUsing$(STR$(UBTranRec(1).TransAmt), "######.##")); q$;

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
'  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
'  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; c$; q$; QPTrim$(UBCustRec(1).CustName); q$;
  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb)
  Zip$ = QPTrim$(UBCustRec(1).ZipCode)
  IF LEN(Zip$) > 5 THEN
    Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  END IF

  PRINT #UBRpt, q$; QPTrim$(STR$(RecNo&)); q$; ","; q$; LocationNumber$; q$; ",";
  PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).CustName); q$; ","; q$; QPTrim$(UBCustRec(1).CustType); q$; ",";
  PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).ADDR1); q$; ","; q$; QPTrim$(UBCustRec(1).ADDR2); q$; ",";
  PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).city); q$; ","; q$; QPTrim$(UBCustRec(1).State); q$; ",";
  PRINT #UBRpt, q$; Zip$; q$; ","; q$; QPTrim$(UBCustRec(1).ServAddr); q$;
  'PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).HPhone); q$

  'PRINT #UBRpt, UBCustRec(1).CustName
  CCCnt = CCCnt + 1
  'IF CCCnt > 99 THEN
  '  ExitFlag = True
  'END IF
RETURN
END SUB

SUB ExportCustomerInformation

  q$ = CHR$(34)
  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBEXPORT.TXT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Customer Export."
  
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    AcctNumber = IdxBuff(Cnt).RecNum
    
    LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb)
    
    '*************************************
    '   Main body of Printing goes here
    IF UBCustRec(1).DelFlag <> -1 THEN 'AND UBCustRec(1).Status = "A" THEN
      
      Zip$ = QPTrim$(UBCustRec(1).ZipCode)
      IF LEN(Zip$) > 5 THEN
        Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
      END IF
      'WRITE #UBRpt, STR$(AcctNumber), LocationNumber$, QPTrim$(UBCustRec(1).CustName), QPTrim$(UBCustRec(1).ADDR1), QPTrim$(UBCustRec(1).ADDR2), QPTrim$(UBCustRec(1).city), QPTrim$(UBCustRec(1).State), Zip$, QPTrim$(UBCustRec(1).ServAddr)

      PRINT #UBRpt, q$; QPTrim$(STR$(AcctNumber)); q$; ","; q$; LocationNumber$; q$; ",";
      PRINT #UBRpt, q$; UBCustRec(1).Status; q$; ","; q$; RTRIM$(UBCustRec(1).CustName); q$; ",";
      PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).ADDR1); q$; ","; q$; QPTrim$(UBCustRec(1).ADDR2); q$; ",";
      PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).city); q$; ","; q$; QPTrim$(UBCustRec(1).State); q$; ",";
      PRINT #UBRpt, q$; Zip$; q$; ","; q$; QPTrim$(UBCustRec(1).ServAddr); q$; ",";
      PRINT #UBRpt, q$; QPTrim$(UBCustRec(1).HPhone); q$;

      FOR zz = 1 TO 15
        PRINT #UBRpt, ","; q$; QPTrim$(UBCustRec(1).Serv(zz).RateCode); q$;
      NEXT
      PRINT #UBRpt,
'      FOR zz = 1 TO 4
'        PRINT #UBRpt, QPTrim$(UBCustRec(1).FlatRates(zz).FRDESC)
'      NEXT
'
'      PRINT #UBRpt, QPTrim$(UBCustRec(1).Zone)
'      PRINT #UBRpt, QPTrim$(UBCustRec(1).CUSTTYPE)
'      PRINT #UBRpt, UBCustRec(1).SeqNumb
'      PRINT #UBRpt, UBCustRec(1).LocMeters(1).MtrNum
'      PRINT #UBRpt, UBCustRec(1).LocMeters(1).AvgUse
'
      
      '*************************************
      
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
     
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  CLOSE UBCust, UBRpt

  'PrintRptFile RptTitle$, "UBEXPORT.TXT", LPTPort%, RetCode%, EntryPoint%

  ERASE IdxBuff, UBCustRec

  ' COLOR 0, 7
   FOR ll = 8 TO 15
     LOCATE ll, 22
     PRINT STRING$(38, 32);
   NEXT ll
   LOCATE 8, 23
   LOCATE 11, 28
   PRINT "File Created is UBEXPORT.TXT";
   LOCATE 13, 29
   PRINT "Press any key to Continue."

'   LPRINT "Map of Customer Export file: UBEXPORT.TXT"
'   LPRINT
'   LPRINT "Field Listing:"
'   LPRINT
'   LPRINT "Account Number"
'   LPRINT "Location Number"
'   LPRINT "Customer Status"
'   LPRINT "Name"
'   LPRINT "Address Line 1"
'   LPRINT "Address Line 2"
'   LPRINT "City"
'   LPRINT "State"
'   LPRINT "Zip"
'   LPRINT "Service Address"
'   LPRINT "Phone#"
'   LPRINT "Service Codes 1 to 15"
'   LPRINT "Flat Rate Codes 1 to 4"
'   LPRINT "Zone Code"
'   LPRINT "Customer Type"
'   LPRINT "Meter Sequence Number"
'   LPRINT "Physical Meter Number"
'   LPRINT "Meter Average Consumption"
'   LPRINT CHR$(12);
   WaitForAction
'
ExitMastCustListing:
  'COLOR 7, 0

END SUB

