DEFINT A-Z
DECLARE SUB PostedReprint (RecNo&)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB BStatusBills ()
DECLARE SUB DisplayPrnScrn (ScrnName$)
DECLARE SUB PrintPostalRtReport ()
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB PrintZipReport ()
DECLARE FUNCTION ChkBillFile% ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB UBLog (Text$)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintAlignMask ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RePrintUtilBills ()
DECLARE SUB RestScrn (Array())
DECLARE SUB SaveScrn (Array())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION GetZipEDigit$ (Zip$)
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION BErrorScrn% (WhatError%, Acct&)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
  
  'Zip$ = Zip$ + EDigit$
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'UBOWNER.bi'
  '$INCLUDE: 'ubpinfo.bi'
  '$INCLUDE: 'ubscsprn.bi'

  CONST False = 0, True = NOT False
  
  STACK 8000
  
  CrLf$ = CHR$(13) + CHR$(10)
  q$ = CHR$(34)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Create Billing Transfer File "
  MChoice$(2) = "Print Zipcode Report"
  MChoice$(3) = "Print Postal Route Report"
  MChoice$(4) = "'B' Status Billing Transfer File"
  MChoice$(5) = "Exit to OS"
  
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
  
Top:
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Bill Processing Menu", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      EXIT DO
    END IF
    
    SELECT CASE Choice
    CASE 1
      PrintUtilBills
    CASE 2
      PrintZipReport
    CASE 3
      PrintPostalRtReport
    CASE 4
      BStatusBills
    CASE 5
      ExitFlag = True
      EXIT DO
    END SELECT
  LOOP
  
  IF INSTR(COMMAND$, "TEST") OR ExitFlag THEN
    HideCursor
    ClearScrn
  ELSE
    RUN "ubbillin"
  END IF
  
  END

FUNCTION BErrorScrn (WhatError%, Acct&)

  BErrorScrn = True

  AcCol = 22
  REDIM TempArray(0) AS INTEGER
  SaveScrn TempArray()

  HideCursor
  BlockClear
  DisplayUBScrn "ERRSCRN1"

  QPrintRC "ACCOUNT:" + STR$(Acct&), 10, AcCol, -1
  QPrintRC "HAS NO METER NUMBER!!", 10, 37, -1
  QPrintRC "Correct and Try Again.", 13, 29, -1

  ShowCursor
  Get.Moose.OR.Key Ky$, MooseButton%, MRow%, MCol%

  IF LEN(Ky$) = 2 THEN
    IF RIGHT$(Ky$, 1) = "g" THEN
      BErrorScrn = False
      'LPRINT Acct&
    END IF
  END IF
  RestScrn TempArray()
  ERASE TempArray

END FUNCTION

SUB BStatusBills

  SHARED Choice$()
  UBLog " IN: B-Status Bill printing."
  CrLf$ = CHR$(13) + CHR$(10)
  Fmt10$ = "##########"
  Fmt10a$ = "#######.##"
  Fmt15$ = "############.##"

  'REDIM PrintRec(1) AS BillOutRecType
  'BillOutRecLen = LEN(PrintRec(1))
  
  NoUpDate = True

  LPIFlag = False
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  REDIM ScrnArray(0)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  'Section to check for customer modifications
  REDIM Choice$(0 TO 3, 0)
  
  Choice$(1, 0) = "Balance Due Customers"
  Choice$(2, 0) = "Credit Due Customers"
  Choice$(3, 0) = "ALL Credit & Balance Due"
  
  LibName$ = "UBPRNBIL"
  ScrnName$ = "PRNTBBIL"
  
  Choice$(0, 0) = "2"
  
  MActionRow = 20
  
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
  
  BillDFld = FldNum%("BILLDATE", Fld())
  BalTFld = FldNum%("BALTYPE", Fld())
  Msgf1 = FldNum%("MSGLINE1", Fld())
  Msgf2 = FldNum%("MSGLINE2", Fld())
  Msgf3 = FldNum%("MSGLINE3", Fld())
  Msgf4 = FldNum%("MSGLINE4", Fld())
  
  Action = 1
  Frm(1).StayOnField = True
  
  FirstTime = True
  Today = Date2Num(DATE$)
  PastDay = Today + 10
  Today$ = DATE$
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayPrnScrn ScrnName$
  
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(BillDFld, 0) = DATE$
      Action = 1
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB BCheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
'    CASE F5KEY
'      PrintAlignMask
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
  
  IF AbortFlag THEN
    UBLog "ABORTED: Bill printing."
    GOTO BExitPrintBill
  END IF
  
  IF AbortFlag GOTO BExitPrintBill
  
  'do bill printing here
  '**************************************************************************
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  'REDIM UBBillRec(1) AS UBTransRecType
  'UBBillRecLen = LEN(UBBillRec(1))

  IF MowFlag THEN
    Ext$ = ".MOW"
  ELSE
    Ext$ = ".BTF"
  END IF

  SCSFileName$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2) + "B"
  FOR Cnt = 1 TO 9
    ChkName$ = SCSFileName$ + QPTrim$(STR$(Cnt)) + Ext$
    IF Exist(ChkName$) = False THEN
      SCSFileName$ = ChkName$
      EXIT FOR
    END IF
  NEXT
  
  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen
  
  NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  'UBRpt = FREEFILE
  'OPEN "UBBILLB.PRN" FOR RANDOM SHARED AS UBRpt LEN = BillOutRecLen

  UBRpt = FREEFILE
  OPEN SCSFileName$ FOR RANDOM SHARED AS UBRpt LEN = BillOutRecLen
  
  UBLog "Printing utility bills to disk."
  ShowProcessingScrn "Creating Utility Bills."
  
  '-----------------------------------------
  PrintedCnt = 0
  NotDone = True
  
  FOR Cnt = 1 TO NumOfRecs
    CustAcctNo& = Cnt
    TBal# = 0
    
    GET UBCust, CustAcctNo&, UBCustRec(1)
    
    IF UBCustRec(1).Status = "B" THEN
      TBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      IF TBal# <> 0 THEN
        REDIM UBBillRec(1) AS UBTransRecType
        SELECT CASE BPrntType
        CASE 1  'credit bills
          IF TBal# > 0 THEN
            GOSUB PrintThemOne
          END IF
        CASE 2  'balance bills
          IF TBal# < 0 THEN
            GOSUB PrintThemOne
          END IF
        CASE 3  'all
          GOSUB PrintThemOne
        END SELECT
      END IF
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  GOTO BSkipEm
  
PrintThemOne:

'  IF MowFlag THEN
'    REDIM PrintRec(1) AS BillOutRec1Type
'  ELSE
'    REDIM PrintRec(1) AS BillOutRec2Type
'  END IF
  
  REDIM PrintRec(1) AS BillOutRecType
  BillOutRecLen = LEN(PrintRec(1))

  IF UBCustRec(1).BillTo = "O" THEN
    GET UBOwn, CustAcctNo&, UBOwnerRec(1)
    OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
    UBCustRec(1).CUSTNAME = OName$
    UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
    UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
    UBCustRec(1).CITY = UBOwnerRec(1).CITY
    UBCustRec(1).STATE = UBOwnerRec(1).STATE
    UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
  END IF
  
  FOR RCnt = 1 TO 15
    IF UBCustRec(1).CurrRevAmts(RCnt) <> 0 THEN
      UBBillRec(1).RevAmt(RCnt) = UBCustRec(1).CurrRevAmts(RCnt)
    END IF
  NEXT
  
  Num2Print = UBCustRec(1).BILLCOPY
  IF Num2Print < 1 THEN Num2Print = 1
  
  PrintedCnt = PrintedCnt + 1
  UBBillRec(1).BillNumber = PrintedCnt
  
  'Look for a valid meter read date,  maybe?
  'from one of the meters
  
  FOR MtrCnt = 1 TO 7
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrType)) > 0 THEN
      UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
      UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
      'UBBillRec(1).CurRead(1) = UBCustRec(1).LocMeters(MtrCnt).CurRead
      'UBBillRec(1).PrevRead(1) = UBCustRec(1).LocMeters(MtrCnt).PrevRead
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      UBBillRec(1).MtrTypes(1) = 1
      EXIT FOR
    END IF
  NEXT
  
  UBBillRec(1).CustAcctNo = CustAcctNo&
  UBBillRec(1).BillDate = BillDate
  UBBillRec(1).PastDueDate = UBBillRec(1).BillDate
  
  BillDate$ = Num2Date$(UBBillRec(1).BillDate)
  PastDueDate$ = BillDate$

  CurrBal# = Round#(UBCustRec(1).CurrBalance)
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
  
'------------------------------
  PrintRec(1).AcctNo = FUsing(STR$(CustAcctNo&), "########")
  PrintRec(1).LocationNum = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
  RSET PrintRec(1).CUSTNAME = QPTrim$(UBCustRec(1).CUSTNAME)
  RSET PrintRec(1).ADDR1 = QPTrim$(UBCustRec(1).ADDR1)
  RSET PrintRec(1).ADDR2 = QPTrim$(UBCustRec(1).ADDR2)
  RSET PrintRec(1).SERVADDR = QPTrim$(UBCustRec(1).SERVADDR)
  RSET PrintRec(1).CITY = QPTrim$(UBCustRec(1).CITY)
  RSET PrintRec(1).STATE = QPTrim$(UBCustRec(1).STATE)
  RSET PrintRec(1).ZIPCODE = QPTrim$(UBCustRec(1).ZIPCODE)
  PrintRec(1).BillType = "N"
  PrintRec(1).DepAppAmt = ""
  PrintRec(1).PrevDue = FUsing(STR$(Previous#), Fmt15$)
  PrintRec(1).CurrDue = FUsing(STR$(CurrBal#), Fmt15$)
  PrintRec(1).TotalDue = FUsing(STR$(TotalAmt#), Fmt15$)
  PrintRec(1).CurrDate = DateRead$
  PrintRec(1).PrevDate = PrevDate$

  IF WFoundMtr THEN
    PrintRec(1).CurrRead = FUsing(STR$(WCurrRead&), Fmt10$)
    PrintRec(1).PrevRead = FUsing(STR$(WPrevRead&), Fmt10$)
    PrintRec(1).Consump = FUsing(STR$(WUsageAmt&), Fmt10$)
    PrintRec(1).ServDays = FUsing(STR$(DaysINRead), "####")
  ELSE
    PrintRec(1).CurrRead = ""
    PrintRec(1).PrevRead = ""
    PrintRec(1).Consump = ""
    PrintRec(1).ServDays = ""
  END IF

  IF IFoundMtr = False THEN
    PrintRec(1).ICurrRead = ""
    PrintRec(1).IPrevRead = ""
    PrintRec(1).IConsump = ""
    PrintRec(1).IServDays = ""
  ELSE
    PrintRec(1).ICurrRead = FUsing(STR$(ICurrRead&), Fmt10$)
    PrintRec(1).IPrevRead = FUsing(STR$(IPrevRead&), Fmt10$)
    PrintRec(1).IConsump = FUsing(STR$(IUsageAmt&), Fmt10$)
    PrintRec(1).IServDays = FUsing(STR$(DaysINRead), "####")
  END IF
  
  FOR Serv = 1 TO 15
    PrintRec(1).ServInfo(Serv).ServText = QPTrim$(UBSetUpRec(1).Revenues(Serv).REVNAME)
    PrintRec(1).ServInfo(Serv).ServAmt = FUsing(STR$(UBBillRec(1).RevAmt(Serv)), Fmt10a$)
  NEXT

  PrintRec(1).MtrType = MtrType$
  IF CubicMtr THEN
    PrintRec(1).MtrUnit = "C"
  ELSE
    PrintRec(1).MtrUnit = "G"
  END IF

  PrintRec(1).BillDate = BillDate$
  PrintRec(1).PastDueDate = PastDueDate$
  IF DidADraftFlag THEN
    PrintRec(1).DraftDate = DraftDate$
  ELSE
    PrintRec(1).DraftDate = ""
  END IF
  RSET PrintRec(1).MsgLine1 = Message$
  RSET PrintRec(1).MsgLine2 = Msg2$
  RSET PrintRec(1).MsgLine3 = Msg3$
  RSET PrintRec(1).MsgLine4 = Msg4$

'unrem
'  LSET PrintRec(1).MtrNumb = MtrNumb$
  
  PrintRec(1).CrLf = CrLf$
  PUT #UBRpt, , PrintRec(1)

'------------------------------
'**********************************
  RETURN
  
BSkipEm:
  
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF
  
  IF FFFlag THEN
    PRINT #UBRpt, CHR$(12);
  END IF
  
  CLOSE
  
  IF AbortFlag THEN
    UBLog "ABORTED: Bill printing, AFTER START."
    GOTO BExitPrintBill
  ELSE
    UBLog "Finished printing to disk."
  END IF
  
  'UBLog "Updated: Bill Information File."
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, BillInfoRec

  IF NOT AbortFlag THEN
   ' PrintRptFile "Utility Bill Printing ", "UBBILLS.MOW", 1, RetCode, 1
    BlockClear
    DisplayUBScrn "UPDATEOK"
    QPrintRC "File Name:  " + SCSFileName$, 10, 29, 112
    WaitForAction
  END IF
  
'  IF NOT AbortFlag THEN
'    PrintRptFile "Utility Bill Printing ", SCSFileName$, 1, RetCode, 1
'  END IF
  
  '**************************************************************************
  'end bill printing
  GOTO BExitPrintBill:
  
  '******************
BCheckReqFields:
  
  BillDate = Date2Num(Form$(BillDFld, 0))
  BALType$ = QPTrim$(LEFT$(Form$(BalTFld, 0), 1))
  
  Message$ = Form$(Msgf1, 0)
  Msg2$ = QPTrim$(Form$(Msgf2, 0))
  Msg3$ = QPTrim$(Form$(Msgf3, 0))
  Msg4$ = QPTrim$(Form$(Msgf4, 0))
  
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
  ELSEIF LEN(BALType$) = 0 THEN
    SaveScrn ScrnArray()
    DisplayPrnScrn "BADBTYPE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  IF ReqFldsOK THEN
    SELECT CASE BALType$
    CASE "B"
      BPrntType = 1
    CASE "C"
      BPrntType = 2
    CASE "A"
      BPrntType = 3
    END SELECT
  END IF
  
  RETURN
  
BGetOut:
BExitPrintBill:
  
  UBLog "OUT: B-Status Bill Printing." + CrLf$
  
  
END SUB

SUB DisplayPrnScrn (ScrnName$)
  LibFile2Scrn "UBPRNBIL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB PrintPostalRtReport
  
  UBLog " IN: Postal Route Report."
  
  IF NOT ChkBillFile% THEN
    UBLog "ERROR: NO BILL FILE."
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "No Billing Information File Found!", 11, 24, -1
    WaitForAction
    GOTO ExitPostalReport
  END IF
  
  RptName$ = "UBPOSTRT.RPT"
  
  MaxLine = 50
  
  ShowProcessingScrn "Reading Billing Information."
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  
  BHandle = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS BHandle LEN = UBBillRecLen
  NumBillRecs& = LOF(BHandle) \ UBBillRecLen
  
  REDIM PostalIdx(1 TO NumBillRecs&)  AS UBPostalIndexType
  IdxRecLen = LEN(PostalIdx(1))
  
  FOR Cnt& = 1 TO NumBillRecs&
    GET BHandle, Cnt&, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
      GET CHandle, UBBillRec(1).CustAcctNo, UBCustRec(1)
      PostalCnt = PostalCnt + 1
      RSET PostalIdx(PostalCnt).Route = QPTrim$(UBCustRec(1).POSTRTE)
    END IF
    ShowPctComp Cnt&, NumBillRecs&
  NEXT
  CLOSE
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT PostalIdx(1), PostalCnt, 0, 16, 10, 4
  
  PrnCnt = 1
  REDIM PRESERVE PrnArray(1 TO PrnCnt) AS UBPostalIndexType
  PrnArray(PrnCnt).Route = PostalIdx(1).Route
  PrnArray(PrnCnt).RecNum = 1
  
  FOR ZCnt = 2 TO PostalCnt
    HadIt = False
    FOR PCnt = 1 TO PrnCnt
      IF PrnArray(PCnt).Route = PostalIdx(ZCnt).Route THEN
        PrnArray(PCnt).RecNum = PrnArray(PCnt).RecNum + 1
        HadIt = True
        EXIT FOR
      END IF
    NEXT
    IF NOT HadIt THEN
      PrnCnt = PrnCnt + 1
      REDIM PRESERVE PrnArray(1 TO PrnCnt) AS UBPostalIndexType
      PrnArray(PrnCnt).Route = PostalIdx(ZCnt).Route
      PrnArray(PrnCnt).RecNum = 1
    END IF
  NEXT
  ERASE PostalIdx, UBBillRec, UBCustRec
  
  ShowProcessingScrn "Processing Zipcode Report."
  
  UBRpt = FREEFILE
  OPEN RptName$ FOR OUTPUT AS UBRpt
  GOSUB PostalHeader
  
  FOR Cnt = 1 TO PrnCnt
    PRINT #UBRpt, PrnArray(Cnt).Route, TAB(30); USING "#####"; PrnArray(Cnt).RecNum
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLine THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB PostalHeader
    END IF
    
    ShowPctComp Cnt, PrnCnt
  NEXT
  
  GOSUB PostalFooter
  
  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Billing Postal Route Report.", "UBPOSTRT.RPT", 1, RetCode, EntryPoint
  END IF
  
ExitPostalReport:
  UBLog "OUT: Postal Route Report."
  EXIT SUB
  
PostalHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(24); "Billing Postal Route Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, TAB(4); "Postal Route"; TAB(30); "Count"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  RETURN
  
PostalFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, " Unique Postal Routes:"; PrnCnt; "   Bills Printed:"; PostalCnt
  PRINT #UBRpt, CHR$(12)
  RETURN
  
  
END SUB

SUB PrintUtilBills
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  Fmt10$ = "##########"
  Fmt10a$ = "#######.##"
  Fmt15$ = "############.##"
  
  UBLog " IN: Bill printing."
  
  SHARED Choice$()
  
  LPIFlag = False
  
  IF NOT ChkBillFile% THEN
    UBLog "ERROR: NO BILL FILE."
    CursorOff
    BlockClear
    DisplayUBScrn "NON2PRNT"
    WaitForAction
    GOTO ExitPrintBill
  END IF
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  REDIM ScrnArray(0)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  FOR ThisRevCnt = 1 TO 15
    IF INSTR(UBSetUpRec(1).Revenues(ThisRevCnt).REVNAME, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
    END IF
  NEXT
  
  'Section to check for customer modifications
  TownName$ = UCASE$(UBSetUpRec(1).UTILNAME)

  IF INSTR(TownName$, "JOHNSTON") > 0 THEN
    JohnCOFlag = True
  END IF
  IF INSTR(TownName$, "MOWAS") > 0 THEN
    MowFlag = True
  END IF

  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF
  
  REDIM Choice$(0 TO 6, 0)
  
  Choice$(1, 0) = "Customer Name Order"
  Choice$(2, 0) = "Account Number Order"
  Choice$(3, 0) = "Location Number Order"
  Choice$(4, 0) = "Postal Carrier Route Order"
  Choice$(5, 0) = "ZipCode Order"
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    Choice$(6, 0) = "Sequence Number Order"
  END IF
  
  LibName$ = "UBPRNBIL"
  
  ScrnName$ = "PRNTBILL"
  
  Choice$(0, 0) = "6"
  
  MActionRow = 20
  
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
  BillDFld = FldNum%("BILLDATE", Fld())
  PastDFld = FldNum%("PASTDATE", Fld())
  DraftDFld = FldNum%("DRFTDATE", Fld())
  BillOFld = FldNum%("PRNORDER", Fld())
  
  PRDateFld = FldNum%("PRDATE", Fld())
  CRDateFld = FldNum%("CRDATE", Fld())
  
  Msgf1 = FldNum%("MSGLINE1", Fld())
  Msgf2 = FldNum%("MSGLINE2", Fld())
  Msgf3 = FldNum%("MSGLINE3", Fld())
  Msgf4 = FldNum%("MSGLINE4", Fld())
  
  Action = 1
  Frm(1).StayOnField = True
  
  FirstTime = True
  Today = Date2Num(DATE$)
  PastDay = Today + 10
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayPrnScrn ScrnName$
  
  IF IndianFlag THEN
    FOR Cnt = 1 TO 3
      QPrintRC " Description " + QPTrim$(STR$(Cnt)) + ":", Cnt + 15, 20, -1
    NEXT
  END IF
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(BillDFld, 0) = DATE$
      LSET Form$(PastDFld, 0) = Num2Date$(PastDay)
      Action = 1
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
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           '--Cancel button
          PressButton 27, MActionRow, 29, 40
        CASE 41 TO 53           '--F10 button
          PressButton F10Key, MActionRow, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  'PastDate$ = Form$(2, 0)
  
  IF AbortFlag THEN
    UBLog "ABORTED: Bill printing."
    GOTO ExitPrintBill
  END IF
  
  SELECT CASE BillOrder$
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    OKFlag = True
  CASE "A"
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  CASE "L"
    IndexName$ = BookIndexFile
    UsingBook = True
    OKFlag = True
  CASE "P", "Z"
    IdxTypeText$ = "Zip-Code"
    IF BillOrder$ = "P" THEN
      IdxTypeText$ = "Postal Route"
      MakePostalIndex IdxTypeText$
      IndexName$ = TempIndexName
      OKFlag = True
    ELSEIF (BillOrder$ = "Z" AND PSAFlag) OR (BillOrder$ = "Z" AND MowFlag) THEN
      IF MowFlag THEN
        MakeMowZipCodeIndex IdxTypeText$
      ELSE
        MakeZipCodeIndex IdxTypeText$
      END IF
      IndexName$ = TempIndexName
      OKFlag = True
    ELSE
      MakeMowZipCodeIndex IdxTypeText$
      'MakePostalIndex IdxTypeText$
      IndexName$ = TempIndexName
      OKFlag = True
    END IF
  CASE "S"
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  END SELECT
  
  IF AbortFlag GOTO ExitPrintBill

  Today$ = DATE$

  PastDate$ = Form$(PastDFld, 0)
  
  'do bill printing here
  '**************************************************************************
  
  REDIM PrintRec(1) AS BillOutRecType
  BillOutRecLen = LEN(PrintRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM UBDraftPayRec(1) AS UBDraftPayRecType
  UBDraftPayLen = LEN(UBDraftPayRec(1))
  
  BillInfoRec(1).BillDate = Date2Num(Form$(BillDFld, 0))
  BillInfoRec(1).PastDate = Date2Num(Form$(PastDFld, 0))
  BillInfoRec(1).PRDate = Date2Num(Form$(PRDateFld, 0))
  BillInfoRec(1).CRDate = Date2Num(Form$(CRDateFld, 0))
  BillInfoRec(1).DrftDate = Date2Num(Form$(DraftDFld, 0))
  BillInfoRec(1).PrnOrder = QPTrim$(Form$(BillOFld, 0))
  BillInfoRec(1).MsgLine1 = Form$(Msgf1, 0)
  BillInfoRec(1).MsgLine2 = Form$(Msgf2, 0)
  BillInfoRec(1).MsgLine3 = Form$(Msgf3, 0)
  BillInfoRec(1).MsgLine4 = Form$(Msgf4, 0)
  
  'GOSUB UpdateInfoRec
  
  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  'KillFile "UBBILLS.MOW"
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen

'add mowasa
  IF MowFlag THEN
    Ext$ = ".MOW"
  ELSE
    Ext$ = ".BTF"
  END IF

  SCSFileName$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2) + "N"
  FOR Cnt = 1 TO 9
    ChkName$ = SCSFileName$ + QPTrim$(STR$(Cnt)) + Ext$
    IF Exist(ChkName$) = False THEN
      SCSFileName$ = ChkName$
      EXIT FOR
    END IF
  NEXT

  UBRpt = FREEFILE
  OPEN SCSFileName$ FOR RANDOM SHARED AS UBRpt LEN = BillOutRecLen

'  UBRpt = FREEFILE
'  OPEN "UBBILLS.MOW" FOR RANDOM SHARED AS UBRpt LEN = BillOutRecLen
  
  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF
  
  UBLog "Printing utility bills to disk."
  ShowProcessingScrn "Creating Transfer Bill File."
  
  '-----------------------------------------
  PrintedCnt = 0
  NotDone = True
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      CustAcctNo& = Cnt
    ELSE
      CustAcctNo& = IndexArray(Cnt).RecNum
    END IF
    GET UBCust, CustAcctNo&, UBCustRec(1)
    '121598 Finished adding bill to owner
    IF UBCustRec(1).BillTo = "O" THEN           'if they want to send the bill
      'temp copy the owners info to the customers rec
      GET UBOwn, CustAcctNo&, UBOwnerRec(1)
      OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
      UBCustRec(1).CUSTNAME = OName$
      UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
      UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
      UBCustRec(1).CITY = UBOwnerRec(1).CITY
      UBCustRec(1).STATE = UBOwnerRec(1).STATE
      UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
    END IF
    'Num2Print = UBCustRec(1).BILLCOPY
    'IF Num2Print < 1 THEN
    Num2Print = 1
    IF UBCustRec(1).Status <> "F" AND UBCustRec(1).Status <> "I" THEN
      GET UBBill, CustAcctNo&, UBBillRec(1)
      UBBillRec(1).TransDate = BillDate
      UBBillRec(1).TransDesc = "UTILITY BILL"
      IF UBBillRec(1).ActiveFlag THEN
        IF UBBillRec(1).TransAmt <> 0 OR UBCustRec(1).PrevBalance THEN
          PrintedCnt = PrintedCnt + 1
          UBBillRec(1).BillNumber = PrintedCnt
          '****************
          '02-05-97 added code to get a valid meter read date,  maybe?  ;-]
          '         from one of the meters
          FOR MtrCnt = 1 TO 7
            IF UBBillRec(1).MtrTypes(MtrCnt) > 0 THEN
              UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
              UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
              EXIT FOR
            END IF
          NEXT
          '02-05-97 this is to add a read date to a bill that has no metered
          '         services
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
          
          UBBillRec(1).BillDate = BillDate
          UBBillRec(1).PastDueDate = PastDate
          UBBillRec(1).DraftDate = DraftDate
          UBBillRec(1).BillMsg = Message$
          
          'these are for reprinting bills
          'UBBillRec(1)CustLocation = CustAcctNo&
          UBBillRec(1).CustAcctNo = CustAcctNo&
          BillDate$ = Num2Date$(UBBillRec(1).BillDate)
          
          'if they entered a previous read date
          IF BillInfoRec(1).PRDate > 0 THEN
            PrevDate$ = Num2Date$(BillInfoRec(1).PRDate)
            PastDate$ = Num2Date$(BillInfoRec(1).PRDate)
            DateRead$ = Num2Date$(BillInfoRec(1).CRDate)
            DaysINRead = BillInfoRec(1).CRDate - BillInfoRec(1).PRDate
          ELSE
            PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
            PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
            DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
            DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
          END IF
          
          PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)
          
          PUT UBBill, CustAcctNo&, UBBillRec(1)
          
          TotalTax# = 0
          FOR TaxCnt = 1 TO MaxRevsCnt
            TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
          NEXT
          DidADraftFlag = False
          IF UseDraftFlag AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag THEN
            UBDraftPayRec(1).CustAcctNum = CustAcctNo&
            UBDraftPayRec(1).DraftAmt = UBBillRec(1).TransAmt
            PUT UBDraft, , UBDraftPayRec(1)
            DidADraftFlag = True
          END IF
          
'***************************************************
          WFoundMtr = False
          IFoundMtr = False
          CubicMtr = False

          WCurrRead& = 0
          WPrevRead& = 0
          WUsageAmt& = 0
          ICurrRead& = 0
          IPrevRead& = 0
          IUsageAmt& = 0
          MtrNumb$ = ""
          MtrType$ = ""

          FOR mChk = 1 TO 7
            IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
              MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MTRNUM)
              IF LEN(MtrNumb$) = 0 THEN
                MtrNumb$ = "???"
              END IF
              MtrTyp = UBBillRec(1).MtrTypes(mChk)
              SELECT CASE MtrTyp
              CASE 1
                IFoundMtr = True
                ICurrRead& = UBBillRec(1).CurRead(mChk)
                IPrevRead& = UBBillRec(1).PrevRead(mChk)
                IUsageAmt& = ICurrRead& - IPrevRead&
                IF IUsageAmt& < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(IPrevRead&)) - 1)
                  IUsageAmt& = (MaxMeterAmt& - IPrevRead&) + ICurrRead&
                END IF
                IF UBCustRec(1).LocMeters(mChk).MtrUnit = "C" THEN
                  CubicMtr = True
                  IUsageAmt& = IUsageAmt& * 7.481
                END IF
                IF MtrTyp = 1 THEN
                  MtrType$ = "W"
                ELSEIF MtrTyp = 2 THEN
                  MtrType$ = "S"
                ELSEIF MtrTyp = 3 THEN
                  MtrType$ = "C"
                ELSEIF MtrType = 7 THEN
                  MtrType$ = "T"
                END IF
              CASE 2, 3, 7
                WFoundMtr = True
                WCurrRead& = UBBillRec(1).CurRead(mChk)
                WPrevRead& = UBBillRec(1).PrevRead(mChk)
                WUsageAmt& = WCurrRead& - WPrevRead&
                IF WUsageAmt& < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
                  WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
                END IF
                IF UBCustRec(1).LocMeters(mChk).MtrUnit = "C" THEN
                  CubicMtr = True
                  WUsageAmt& = WUsageAmt& * 7.481
                END IF
                IF MtrTyp = 1 THEN
                  MtrType$ = "W"
                ELSEIF MtrTyp = 2 THEN
                  MtrType$ = "S"
                ELSEIF MtrTyp = 3 THEN
                  MtrType$ = "C"
                ELSEIF MtrType = 7 THEN
                  MtrType$ = "T"
                END IF
              END SELECT
            END IF
          NEXT
          
          IF WFoundMtr = False AND IFoundMtr = False THEN
            'if no metered services then adjust read dates to billdate
            'and billdate - 30
            DateRead$ = Num2Date$(UBBillRec(1).BillDate)
            PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
            DaysINRead = 30
          END IF
          
          Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
          TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
          
          '          IF FinalFlag AND CDeposit# THEN
          '            TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
          '          END IF
          
          IF TotalAmt# > 0 THEN
            TenPct# = 0
            IF DaysINRead < 1 THEN DaysINRead = 1
            AvgCst# = Round#(TotalAmt# / DaysINRead)
          ELSE
            TenPct# = 0
            AvgCst# = 0
          END IF
          
          REDIM PrintRec(1) AS BillOutRecType

          PrintRec(1).AcctNo = FUsing(STR$(CustAcctNo&), "########")
          PrintRec(1).LocationNum = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
          RSET PrintRec(1).CUSTNAME = QPTrim$(UBCustRec(1).CUSTNAME)
          RSET PrintRec(1).ADDR1 = QPTrim$(UBCustRec(1).ADDR1)
          RSET PrintRec(1).ADDR2 = QPTrim$(UBCustRec(1).ADDR2)
          RSET PrintRec(1).SERVADDR = QPTrim$(UBCustRec(1).SERVADDR)
          RSET PrintRec(1).CITY = QPTrim$(UBCustRec(1).CITY)
          RSET PrintRec(1).STATE = QPTrim$(UBCustRec(1).STATE)
          RSET PrintRec(1).ZIPCODE = QPTrim$(UBCustRec(1).ZIPCODE)
          PrintRec(1).BillType = "N"
          PrintRec(1).DepAppAmt = ""
          PrintRec(1).PrevDue = FUsing(STR$(Previous#), Fmt15$)
          PrintRec(1).CurrDue = FUsing(STR$(UBBillRec(1).TransAmt), Fmt15$)
          PrintRec(1).TotalDue = FUsing(STR$(TotalAmt#), Fmt15$)
          PrintRec(1).CurrDate = DateRead$
          PrintRec(1).PrevDate = PrevDate$
          
          IF WFoundMtr THEN
            PrintRec(1).CurrRead = FUsing(STR$(WCurrRead&), Fmt10$)
            PrintRec(1).PrevRead = FUsing(STR$(WPrevRead&), Fmt10$)
            PrintRec(1).Consump = FUsing(STR$(WUsageAmt&), Fmt10$)
            PrintRec(1).ServDays = FUsing(STR$(DaysINRead), "####")
          ELSE
            PrintRec(1).CurrRead = ""
            PrintRec(1).PrevRead = ""
            PrintRec(1).Consump = ""
            PrintRec(1).ServDays = ""
          END IF

          IF IFoundMtr = False THEN
            PrintRec(1).ICurrRead = ""
            PrintRec(1).IPrevRead = ""
            PrintRec(1).IConsump = ""
            PrintRec(1).IServDays = ""
          ELSE
            PrintRec(1).ICurrRead = FUsing(STR$(ICurrRead&), Fmt10$)
            PrintRec(1).IPrevRead = FUsing(STR$(IPrevRead&), Fmt10$)
            PrintRec(1).IConsump = FUsing(STR$(IUsageAmt&), Fmt10$)
            PrintRec(1).IServDays = FUsing(STR$(DaysINRead), "####")
          END IF
          
          FOR Serv = 1 TO 15
            PrintRec(1).ServInfo(Serv).ServText = QPTrim$(UBSetUpRec(1).Revenues(Serv).REVNAME)
            PrintRec(1).ServInfo(Serv).ServAmt = FUsing(STR$(UBBillRec(1).RevAmt(Serv)), Fmt10a$)
          NEXT

          PrintRec(1).MtrType = MtrType$
          IF CubicMtr THEN
            PrintRec(1).MtrUnit = "C"
          ELSE
            PrintRec(1).MtrUnit = "G"
          END IF

          PrintRec(1).BillDate = BillDate$
          PrintRec(1).PastDueDate = PastDueDate$
          IF DidADraftFlag THEN
            PrintRec(1).DraftDate = DraftDate$
          ELSE
            PrintRec(1).DraftDate = ""
          END IF
          RSET PrintRec(1).MsgLine1 = Message$
          RSET PrintRec(1).MsgLine2 = Msg2$
          RSET PrintRec(1).MsgLine3 = Msg3$
          RSET PrintRec(1).MsgLine4 = Msg4$
'unrem
'          LSET PrintRec(1).MtrNumb = MtrNumb$

          PrintRec(1).CrLf = CrLf$
          PUT #UBRpt, , PrintRec(1)
          
          ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
          
        END IF
      ELSE      'mod for non-profits
        IF UBBillRec(1).NONProfit = "Y" THEN
          GOSUB DOChurchTrans
        END IF
        '***
      END IF
    END IF
    
    'IF AskAbandonPrint% THEN
    '  AbortFlag = True
    '  EXIT FOR
    'END IF
    ShowPctComp Cnt, NumOfRecs
SkipEm:
  NEXT
  'IF LPIFlag = -2 THEN
  '  PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  'END IF
  
  'IF FFFlag THEN
  '  PRINT #UBRpt, CHR$(12);
  'END IF
  

  CLOSE
  
  IF AbortFlag THEN
    UBLog "ABORTED: Bill printing, AFTER START."
    GOTO ExitPrintBill
  ELSE
    UBLog "Finished printing to disk."
  END IF
  
  BillInfoRec(1).FrstBill = 1
  BillInfoRec(1).LastBill = PrintedCnt
  BillInfoRec(1).BillDate = Date2Num(Form$(BillDFld, 0))
  BillInfoRec(1).PastDate = Date2Num(Form$(PastDFld, 0))
  BillInfoRec(1).PRDate = Date2Num(Form$(PRDateFld, 0))
  BillInfoRec(1).CRDate = Date2Num(Form$(CRDateFld, 0))
  BillInfoRec(1).DrftDate = Date2Num(Form$(DraftDFld, 0))
  BillInfoRec(1).PrnOrder = QPTrim$(Form$(BillOFld, 0))
  
  BillInfoRec(1).MsgLine1 = Form$(Msgf1, 0)
  BillInfoRec(1).MsgLine2 = Form$(Msgf2, 0)
  BillInfoRec(1).MsgLine3 = Form$(Msgf3, 0)
  BillInfoRec(1).MsgLine4 = Form$(Msgf4, 0)
  
  UBFile = FREEFILE
  OPEN "UBPINFON.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  PUT #UBFile, 1, BillInfoRec(1)
  CLOSE
  UBLog "Updated: Bill Information File."
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, BillInfoRec
  
  IF NOT AbortFlag THEN
   ' PrintRptFile "Utility Bill Printing ", SCSFileName$, 1, RetCode, 1
    BlockClear
    DisplayUBScrn "UPDATEOK"
    QPrintRC "File Name:  " + SCSFileName$, 10, 29, 112
    WaitForAction
  END IF
  
  '**************************************************************************
  'end bill printing
  GOTO ExitPrintBill:
  '******************
  
  '******************
CheckReqFields:
  
  BillDate = Date2Num(Form$(BillDFld, 0))
  PastDate = Date2Num(Form$(PastDFld, 0))
  DraftDate = Date2Num(Form$(DraftDFld, 0))
  DueDate$ = Form$(PastDFld, 0)
  BillOrder$ = QPTrim$(LEFT$(Form$(BillOFld, 0), 1))
  
  Message$ = QPTrim$(Form$(Msgf1, 0))
  Msg2$ = QPTrim$(Form$(Msgf2, 0))
  Msg3$ = QPTrim$(Form$(Msgf3, 0))
  Msg4$ = QPTrim$(Form$(Msgf4, 0))
  
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
    Frm(1).FldNo = DraftDFld
  ELSEIF LEN(BillOrder$) = 0 THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "NOBORDER"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 6
  ELSE
    ReqFldsOK = True
  END IF
  
  IF UseDraftFlag THEN
    DraftDate$ = Form$(DraftDFld, 0)
    DFFileName$ = "DF" + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 4, 2) + RIGHT$(DraftDate$, 2) + ".DAT"
  END IF
  
  RETURN
  
DOChurchTrans:
  'mod for cleveland***
  FOR MtrCnt = 1 TO 7
    IF UBBillRec(1).MtrTypes(MtrCnt) > 0 THEN
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
  
  UBBillRec(1).BillDate = BillDate
  UBBillRec(1).PastDueDate = PastDate
  UBBillRec(1).DraftDate = DraftDate
  UBBillRec(1).BillMsg = Message$
  
  'UBBillRec(1)CustLocation = CustAcctNo&
  UBBillRec(1).CustAcctNo = CustAcctNo&
  
  PUT UBBill, CustAcctNo&, UBBillRec(1)
  
  RETURN
  '***
GetOut:

ExitPrintBill:
  'STOP
  UBLog "OUT: Bill Printing." + CrLf$
  
END SUB

SUB PrintZipReport
  
  UBLog " IN: Zipcode Report."
  
  IF NOT ChkBillFile% THEN
    UBLog "ERROR: NO BILL FILE."
    CursorOff
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "No Billing Information File Found!", 11, 24, -1
    WaitForAction
    GOTO ExitZipReport
  END IF
  
  MaxLine = 50
  
  ShowProcessingScrn "Reading Billing Information."
  Dash80$ = STRING$(80, "-")
  
  '  REDIM UBSetUpRec(1) AS UBSetupRecType
  '  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  '
  '  TownName$ = UBSetUpRec(1).UTILNAME
  '  IF INSTR(TownName$, "HAMLET") > 0 THEN
  '    HamFlag = True
  '  END IF
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  
  BHandle = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS BHandle LEN = UBBillRecLen
  NumBillRecs& = LOF(BHandle) \ UBBillRecLen
  
  REDIM ZipIndex(1 TO NumBillRecs&)  AS UBPostalIndexType
  IdxRecLen = LEN(ZipIndex(1))
  
  FOR Cnt& = 1 TO NumBillRecs&
    GET BHandle, Cnt&, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
      GET CHandle, UBBillRec(1).CustAcctNo, UBCustRec(1)
      ZipCnt = ZipCnt + 1
      'IF HamFlag THEN
      RSET ZipIndex(ZipCnt).ZIPCODE = QPTrim$(LEFT$(UBCustRec(1).ZIPCODE, 5))
      'ELSE
      '  RSET ZipIndex(ZipCnt).ZipCode = QPTrim$(UBCustRec(1).ZipCode)
      'END IF
    END IF
    ShowPctComp Cnt&, NumBillRecs&
  NEXT
  CLOSE
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ZipIndex(1), ZipCnt, 0, 16, 0, 10
  
  PrnCnt = 1
  REDIM PRESERVE PrnArray(1 TO PrnCnt) AS UBPostalIndexType
  PrnArray(PrnCnt).ZIPCODE = ZipIndex(1).ZIPCODE
  PrnArray(PrnCnt).RecNum = 1
  
  FOR ZCnt = 2 TO ZipCnt
    HadIt = False
    FOR PCnt = 1 TO PrnCnt
      IF PrnArray(PCnt).ZIPCODE = ZipIndex(ZCnt).ZIPCODE THEN
        PrnArray(PCnt).RecNum = PrnArray(PCnt).RecNum + 1
        HadIt = True
        EXIT FOR
      END IF
    NEXT
    IF NOT HadIt THEN
      PrnCnt = PrnCnt + 1
      REDIM PRESERVE PrnArray(1 TO PrnCnt) AS UBPostalIndexType
      PrnArray(PrnCnt).ZIPCODE = ZipIndex(ZCnt).ZIPCODE
      PrnArray(PrnCnt).RecNum = 1
    END IF
    ShowPctComp ZCnt, ZipCnt
  NEXT
  ERASE ZipIndex, UBBillRec, UBCustRec
  
  ShowProcessingScrn "Processing Zipcode Report."
  
  UBRpt = FREEFILE
  OPEN "UBZIPRPT.RPT" FOR OUTPUT AS UBRpt
  GOSUB ZipHeader
  FOR Cnt = 1 TO PrnCnt
    PRINT #UBRpt, PrnArray(Cnt).ZIPCODE, TAB(30); USING "#####"; PrnArray(Cnt).RecNum
    LineCnt = LineCnt + 1
    IF LineCnt > MaxLine THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB ZipHeader
    END IF
    
    ShowPctComp Cnt, PrnCnt
  NEXT
  
  GOSUB ZipFooter
  
  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Billing Zipcode Report.", "UBZIPRPT.RPT", 1, RetCode, EntryPoint
  END IF
  
ExitZipReport:
  UBLog "OUT: Zipcode Report."
  EXIT SUB
  
ZipHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(24); "Billing Zipcode Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, TAB(4); "Zip Code"; TAB(30); "Count"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
  RETURN
  
ZipFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, " Unique Zipcodes:"; PrnCnt; "   Bills Printed:"; ZipCnt
  PRINT #UBRpt, CHR$(12)
  RETURN
  
END SUB

