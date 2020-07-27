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
  
  CONST False = 0, True = NOT False
  
  STACK 8000
  
  CRLF$ = CHR$(13) + CHR$(10)
  q$ = CHR$(34)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Print ALL Utility Bills "
  MChoice$(2) = "Reprint Selected Bills"
  MChoice$(3) = "Print " + q$ + "B" + q$ + " Status Bills"
  MChoice$(4) = "Print Zipcode Report"
  MChoice$(5) = "Print Postal Route Report"
  MChoice$(6) = "Posted Bill Reprinting"
  MChoice$(7) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Bill Printing Menu ", Cnf
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
      RePrintUtilBills
    CASE 3
      BStatusBills
    CASE 4
      PrintZipReport
    CASE 5
      PrintPostalRtReport
    CASE 6
      BlockClear
      OK = MsgBox%("UBPRNBIL", "REPRNTOK")
      IF OK = 2 THEN   'Continue button was pressed
        BlockClear
        LookUp RecNo&, "Posted Bill Reprint", 2, True, True
        IF RecNo& > 0 THEN
          PostedReprint RecNo&
        END IF
      END IF
      Choice = 1
      GOTO Top
    CASE 7
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

SUB BStatusBills
  
  UBLog " IN: B-Status Bill printing."
  
  SHARED Choice$()
  
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
    CASE F5KEY
      PrintAlignMask
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
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen
  
  NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBILLB.PRN" FOR OUTPUT AS UBRpt
  
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

  IF UBCustRec(1).BillTo = "O" THEN
    GET UBOwn, CustAcctNo&, UBOwnerRec(1)
    OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
    UBCustRec(1).CustName = OName$
    UBCustRec(1).Addr1 = UBOwnerRec(1).Addr1
    UBCustRec(1).Addr2 = UBOwnerRec(1).Addr2
    UBCustRec(1).City = UBOwnerRec(1).City
    UBCustRec(1).State = UBOwnerRec(1).State
    UBCustRec(1).ZipCode = UBOwnerRec(1).ZipCode
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
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRType)) > 0 THEN
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

'NOTE: Maccelsfield bill has best code to parse out meter readings
  FOR BillCopies = 1 TO Num2Print
    ''$INCLUDE: 'NORWODBL.BI'        'Norwood
    ''$INCLUDE: 'UBALBERT.BI'        'ALBERTA
    ''$INCLUDE: 'UBALEXML.BI'        'ALEXANDER MILLS
    ''$INCLUDE: 'UBANSON.BI'         'Ansonville NC
    ''$INCLUDE: 'UBANDREW.BI'        'Andrews
    ''$INCLUDE: 'UBAPPLCH.BI'        'APPALACHIA VA
    ''$INCLUDE: 'UBAURORA.BI'        'like
    ''$INCLUDE: 'UBAUTRYV.BI'        'Autryville NC
    ''$INCLUDE: 'UBBEECH.BI2'        'like
    ''$INCLUDE: 'UBBETHEL.BI'        'like
    ''$INCLUDE: 'UBBLADEN.BI'        'Bladen
    ''$INCLUDE: 'UBBRANCH.BI'        'Branchville SC
    ''$INCLUDE: 'UBBROKNL.BI'        'Brookneal
    ''$INCLUDE: 'UBBROKFR.BI'        'BrookFord
    ''$INCLUDE: 'UBBRUNDG.BI'        'Brundage
    ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
    ''$INCLUDE: 'UBCALDWL.BI'        'Caldwell
    ''$INCLUDE: 'UBCASHON.BI'        'Cashion
    ''$INCLUDE: 'UBCELINA.BI'        'Celina
    ''$INCLUDE: 'UBCHESTR.BI'        'Chester
    ''$INCLUDE: 'UBCHILHO.BI'        'Chilhowe
    ''$INCLUDE: 'UBCONCRD.BI'        'Concord
    ''$INCLUDE: 'UBCONWAY.BI'        'Conway
    ''$INCLUDE: 'UBCREWBL.BI'        'Crew
    ''$INCLUDE: 'UBCedar.BI'         'Cedar Bluff VA
    ''$INCLUDE: 'UBDECATR.BI'        'Decatur
    ''$INCLUDE: 'UBDENTON.BI'        'Denton
    ''$INCLUDE: 'UBDUBLIN.BI'        'Dublin NC
    ''$INCLUDE: 'UBEHRHRT.BI'         'Ehrhardt SC
    ''$INCLUDE: 'UBELKTON.BI'        'Elkton
    ''$INCLUDE: 'UBELLEN.BI'         'Ellenboro NC
    ''$INCLUDE: 'UBELLORE.BI'        'Elloree
    ''$INCLUDE: 'UBFAITH.BI'         'Faith
    ''$INCLUDE: 'UBFOUROK.BI'        'Four Oaks NC
    ''$INCLUDE: 'UBGATECT.BI'        'Gate City
    ''$INCLUDE: 'UBGEORTN.BI'        'Georgetown
    ''$INCLUDE: 'UBGILBRT.BI'        'Gilbert Summit
    ''$INCLUDE: 'UBGILES.BI'         'GILES VA
    ''$INCLUDE: 'UBGORDON.BI'        'Gordonsville VA
    ''$INCLUDE: 'UBGRAYSN.BI'        'Grayson VA
    ''$INCLUDE: 'UBGRENVR.BI'        'Greenevers NC
    ''$INCLUDE: 'UBGROVER.BI'        'Grover NC  24 Line Special
    ''$INCLUDE: 'UBHALFAX.BI'        'Halifax VA
    ''$INCLUDE: 'UBHAMLET.BI'        'Hamlet
    ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
    ''$INCLUDE: 'UBHILBIL.BI'        'Hillsville VA
    ''$INCLUDE: 'UBHILDBR.BI'        'Hildabran
    ''$INCLUDE: 'UBHOLLY.BI'         'Holly springs
    ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
    ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
    ''$INCLUDE: 'UBJAMEST.BI'        'Jamestown
    ''$INCLUDE: 'UBJONSVL.BI'        'Jonesville NC
    '$INCLUDE: 'UBJVILLE.BI'        'Johnsonville
    ''$INCLUDE: 'UBKELFRD.BI'        'Kelford NC
    ''$INCLUDE: 'UBKERSHA.bi'        'Kershaw
    ''$INCLUDE: 'UBLANDIS.BI'        'Landis,nc
    ''$INCLUDE: 'UBLINCOL.BI'        'Lincoln County GA
    ''$INCLUDE: 'UBLINDEN.BI'        'Linden,nc
    ''$INCLUDE: 'UBLONGVW.BI'        'Long View nc
    ''$INCLUDE: 'UBLOUISA.BI'        'Louisa
    ''$INCLUDE: 'UBMACCFL.BI'        'Maccelfield
    ''$INCLUDE: 'UBMARSH.BI'         'Marshville, nc
    ''$INCLUDE: 'UBMANTEO.BI'         'Manteo NC
    ''$INCLUDE: 'UBMAYVIL.BI'        'Maysville
    ''$INCLUDE: 'UBMCMICK.BI'        'McCormick sc
    ''$INCLUDE: 'UBMEANVL.BI'        'Meanville SC
    ''$INCLUDE: 'UBMEBANE.BI'        'Mebane Orange Alamance
    ''$INCLUDE: 'UBMIDTWN.BI'        'Middletown VA
    ''$INCLUDE: 'UBMINCO.BI'         'Minco OK
    ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
    ''$INCLUDE: 'UBMOWB.BI'          'Mowasa use this one here
    ''$INCLUDE: 'UBMTGILD.BI'        'Mt Gilead
    ''$INCLUDE: 'UBMTPLES.BI'        'Mt Plesent
    ''$INCLUDE: 'UBMURPHY.BI'        'Murphy
    ''$INCLUDE: 'UBNEWPRT.BI'        'Newport NC
    ''$INCLUDE: 'UBNLENOR.BI'        'North Lenoir
    ''$INCLUDE: 'UBNORLIN.BI'        'Norlina
    ''$INCLUDE: 'UBOAKCTY.BI'        'Oak City
    ''$INCLUDE: 'UBOAKBOR.BI'        'Oakboro
    ''$INCLUDE: 'UBOLDSDS.BI'        'Cleveland NC
    ''$INCLUDE: 'UBOLDSTD.BI'        'McBee SC    Old Standard 21 Line Bill
    ''$INCLUDE: 'UBPAGELN.BI'        'Pageland
    ''$INCLUDE: 'UBPEACHB.BI'        'Peachland
    ''$INCLUDE: 'UBPEMBRK.BI'        'Pembroke VA
    ''$INCLUDE: 'UBPENDTN.BI'        'Pendelton
    ''$INCLUDE: 'UBPINEBF.BI'        'Pinebluff
    ''$INCLUDE: 'UBPLYMTH.BI'        'Plymouth
    ''$INCLUDE: 'UBPOLKTN.BI'        'Polkton
    ''$INCLUDE: 'UBPOUND.BI'         'Pound VA
    ''$INCLUDE: 'UBPRINCE.BI'        'Princeton
    ''$INCLUDE: 'UBREMTON.BI'        'Remington
    ''$INCLUDE: 'UBROLES.BI'         'Rolesville
    ''$INCLUDE: 'UBRUSLCO.BI'        'Russell Co.
    ''$INCLUDE: 'UBROSEBL.BI'        'Roseboro
    ''$INCLUDE: 'UBRYAN.BI'        'Ryan OK
    ''$INCLUDE: 'UBSALEM.BI'         'Salem SC
    ''$INCLUDE: 'UBSALTVL.BI'        'Saltville
    ''$INCLUDE: 'UBSALUCO.BI'        'Saluda CO
    ''$INCLUDE: 'UBSAWMIL.BI'        'Sawmills
    ''$INCLUDE: 'UBSCOTTS.BI'        'Scottsburg
    ''$INCLUDE: 'UBSPENNC.BI'        'Spencer NC
    ''$INCLUDE: 'UBSPENTN.BI'        'Spencer TN
    ''$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
    ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
    ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
    ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
    '$INCLUDE: 'UBSUNSET.BI'        'Sunset Beech
    ''$INCLUDE: 'UBTALOGA.BI'        'Taloga
    ''$INCLUDE: 'UBTENNRD.BI'        'Tenn Ridge
    ''$INCLUDE: 'UBTIMSON.BI'        'Timmonsville SC
    ''$INCLUDE: 'UBTROUTM.BI'        'Troutman, nc
    ''$INCLUDE: 'UBTROY.BI'          'Troy
    ''$INCLUDE: 'UBTUSK.BI'          'Tuskaseigee NC
    ''$INCLUDE: 'UBTWINCT.BI'        'Twin City
    ''$INCLUDE: 'UBWADEBL.BI'        'Wade NC
    ''$INCLUDE: 'UBWAGRAM.BI'        'Wagram
    ''$INCLUDE: 'UBWHITKR.BI'        'Whitakers
    ''$INCLUDE: 'UBWHITLK.BI'        'White Lake
    ''$INCLUDE: 'UBWINDSR.BI'        'Windsor
    ''$INCLUDE: 'UBWINGAT.BI'        'Wingate
    ''$INCLUDE: 'UBWRIGHT.BI'        'Wrightsville Beech
    ''$INCLUDE: 'UBYONGBL.BI'        'Youngsville
    ''$INCLUDE: 'UBjoston.BI'        'Johnston
    ''$INCLUDE: 'UBlawrnc.BI'        'Lawrenceville
    ''$INCLUDE: 'UBnewstd.BI'        '
    ''$INCLUDE: 'UBYADKIN.BI'        'Yadkinville
    ''$INCLUDE: 'UBLUGOFW.bi'        '
  NEXT
  
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
    PrintRptFile "Utility Bill Printing ", "UBBILLB.PRN", 1, RetCode, 1
  END IF
  
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
  
  UBLog "OUT: B-Status Bill Printing." + CRLF$
  
  
END SUB

SUB DisplayPrnScrn (ScrnName$)
  LibFile2Scrn "UBPRNBIL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB PostedReprint (CustAcctNo&)

  SHARED Choice$()

  UBLog " IN: Posted Bill Reprinting."

  REDIM Choice$(0 TO 2, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Customer"
  Choice$(2, 0) = "Owner"

  LibName$ = "UBPRNBIL"
  ScrnName$ = "PRNTBBIL"
  
  NoUpDate = True

  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, CustAcctNo&, UBCustRec(1)
  CLOSE UBCust

  SELECT CASE UBCustRec(1).BillTo
  CASE "O"
    OFlag = True
  END SELECT

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  TownName$ = UBSetUpRec(1).UTILNAME

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen

  FoundBill = False

  ThisTrans& = UBCustRec(1).LastTrans
  DO WHILE ThisTrans& > 0
    GET #UBTran, ThisTrans&, UBBillRec(1)
    IF UBBillRec(1).TransType = TranUtilityBill THEN
      FoundBill = True
      EXIT DO
    END IF
    ThisTrans& = UBBillRec(1).PrevTrans
  LOOP
  CLOSE UBTran

  IF FoundBill = False THEN   'show can't find a bill error
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "This Accounts Has NO Bill Transaction", 10, 22, -1
    QPrintRC "Nothing to Reprint!", 12, 31, -1
    WaitForAction
    GOTO Reprint1Exit
  END IF
  
  LibName$ = "UBPRNBIL"
  ScrnName$ = "PRNTPBIL"
  
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
  FirstTime = True
  MActionRow = 19

  BlockClear

  DisplayPrnScrn ScrnName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      IF OFlag THEN
        LSET Form$(1, 0) = Choice$(2, 0)
      ELSE
        LSET Form$(1, 0) = Choice$(1, 0)
      END IF
      QPrintRC LEFT$(QPTrim$(UBCustRec(1).CustName), 30), 7, 34, 14
      QPrintRC Num2Date$(UBBillRec(1).BillDate), 9, 34, 14
      Action = 1
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      ExitFlag = True
    CASE F5KEY
      PrintAlignMask
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
    UBLog "ABORTED: Posted Bill Reprinting."
    GOTO Reprint1Exit
  END IF

  UBLog "PBR: Reprinting Bill for ACCT:" + STR$(CustAcctNo&)

  Message$ = QPTrim$(Form$(2, 0))
  Msg2$ = QPTrim$(Form$(3, 0))
  Msg3$ = QPTrim$(Form$(4, 0))
  Msg4$ = QPTrim$(Form$(5, 0))

  BillInfoRec(1).MsgLine1 = Message$
  BillInfoRec(1).MsgLine2 = Msg2$
  BillInfoRec(1).MsgLine3 = Msg3$
  BillInfoRec(1).MsgLine4 = Msg4$

  PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
  PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
  DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
  DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate

  BillDate$ = Num2Date$(UBBillRec(1).BillDate)

'This section attempts to get the previous balance at the time the bill
'was originally printed it is only temporary.
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).CurrBalance = 0
  IF Round#(UBBillRec(1).RunBalance) <> Round#(UBBillRec(1).TransAmt) THEN
    UBCustRec(1).CurrBalance = Round#(UBBillRec(1).RunBalance - UBBillRec(1).TransAmt)
  END IF
'*********************************
  IF LEFT$(Form$(1, 0), 1) = "O" THEN
    UBOwn = FREEFILE
    OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen
    GET UBOwn, CustAcctNo&, UBOwnerRec(1)
    OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
    UBCustRec(1).CustName = OName$
    UBCustRec(1).Addr1 = UBOwnerRec(1).Addr1
    UBCustRec(1).Addr2 = UBOwnerRec(1).Addr2
    UBCustRec(1).City = UBOwnerRec(1).City
    UBCustRec(1).State = UBOwnerRec(1).State
    UBCustRec(1).ZipCode = UBOwnerRec(1).ZipCode
    CLOSE UBOwn
  END IF
  
  UBRpt = FREEFILE
  OPEN "UBBILLPP.PRN" FOR OUTPUT AS UBRpt
  PrintedCnt = 1
  ''$INCLUDE: 'NORWODBL.BI'        'Norwood
  ''$INCLUDE: 'UBALBERT.BI'        'ALBERTA
  ''$INCLUDE: 'UBALEXML.BI'        'ALEXANDER MILLS
  ''$INCLUDE: 'UBANSON.BI'         'Ansonville NC
  ''$INCLUDE: 'UBANDREW.BI'         'Andrews
  ''$INCLUDE: 'UBAPPLCH.BI'        'APPALACHIA VA
  ''$INCLUDE: 'UBAURORA.BI'        'like
  ''$INCLUDE: 'UBAUTRYV.BI'        'Autryville NC
  ''$INCLUDE: 'UBBEECH.BI2'        'like
  ''$INCLUDE: 'UBBETHEL.BI'        'like
  ''$INCLUDE: 'UBBLADEN.BI'        'Bladen
  ''$INCLUDE: 'UBBRANCH.BI'        'Branchville SC
  ''$INCLUDE: 'UBBROKNL.BI'        'Brookneal
  ''$INCLUDE: 'UBBROKFR.BI'        'BrookFord
  ''$INCLUDE: 'UBBRUNDG.BI'        'Brundage
  ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
  ''$INCLUDE: 'UBCALDWL.BI'        'Caldwell
  ''$INCLUDE: 'UBCASHON.BI'        'Cashion
  ''$INCLUDE: 'UBCELINA.BI'        'Celina
  ''$INCLUDE: 'UBCHESTR.BI'        'Chester
  ''$INCLUDE: 'UBCHILHO.BI'        'Chilhowe
  ''$INCLUDE: 'UBCONCRD.BI'        'Concord
  ''$INCLUDE: 'UBCONWAY.BI'        'Conway
  ''$INCLUDE: 'UBCREWBL.BI'        'Crew
  ''$INCLUDE: 'UBCedar.BI'         'Cedar Bluff VA
  ''$INCLUDE: 'UBDECATR.BI'        'Decatur
  ''$INCLUDE: 'UBDENTON.BI'        'Denton
  ''$INCLUDE: 'UBDUBLIN.BI'        'Dublin NC
  ''$INCLUDE: 'UBEHRHRT.BI'         'Ehrhardt SC
  ''$INCLUDE: 'UBELKTON.BI'         'Elkton
  ''$INCLUDE: 'UBELLEN.BI'         'Ellenboro NC
  ''$INCLUDE: 'UBELLORE.BI'        'Elloree
  ''$INCLUDE: 'UBFAITH.BI'         'Faith
  ''$INCLUDE: 'UBFOUROK.BI'        'Four Oaks NC
  ''$INCLUDE: 'UBGATECT.BI'        'Gate City
  ''$INCLUDE: 'UBGEORTN.BI'        'Georgetown
  ''$INCLUDE: 'UBGILBRT.BI'        'Gilbert Summit
  ''$INCLUDE: 'UBGILES.BI'         'GILES VA
  ''$INCLUDE: 'UBGORDON.BI'        'Gordonsville VA
  ''$INCLUDE: 'UBGRAYSN.BI'        'Grayson VA
  ''$INCLUDE: 'UBGRENVR.BI'        'Greenevers NC
  ''$INCLUDE: 'UBGROVER.BI'        'Grover NC  24 Line Special
  ''$INCLUDE: 'UBHALFAX.BI'        'Halifax VA
  ''$INCLUDE: 'UBHAMLET.BI'        'Hamlet
  ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
  ''$INCLUDE: 'UBHILBIL.BI'        'Hillsville VA
  ''$INCLUDE: 'UBHILDBR.BI'        'Hildabran
  ''$INCLUDE: 'UBHOLLY.BI'         'Holly springs
  ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
  ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
  ''$INCLUDE: 'UBJAMEST.BI'        'Jamestown
  ''$INCLUDE: 'UBJONSVL.BI'        'Jonesville NC
  ''$INCLUDE: 'UBJVILLE.BI'        'Johnsonville
  ''$INCLUDE: 'UBKELFRD.BI'        'Kelford NC
  ''$INCLUDE: 'UBKERSHA.bi'        'Kershaw
  ''$INCLUDE: 'UBLANDIS.BI'        'Landis,nc
  ''$INCLUDE: 'UBLINCOL.BI'        'Lincoln County GA
  ''$INCLUDE: 'UBLINDEN.BI'        'Linden,nc
  ''$INCLUDE: 'UBLONGVW.BI'        'Long View nc
  ''$INCLUDE: 'UBLOUISA.BI'        'Louisa
  ''$INCLUDE: 'UBMACCFL.BI'        'Maccelfield
  ''$INCLUDE: 'UBMARSH.BI'         'Marshville, nc
  ''$INCLUDE: 'UBMANTEO.BI'         'Manteo NC
  ''$INCLUDE: 'UBMAYVIL.BI'        'Maysville
  ''$INCLUDE: 'UBMCMICK.BI'        'McCormick sc
  ''$INCLUDE: 'UBMEANVL.BI'        'Meanville SC
  ''$INCLUDE: 'UBMEBANE.BI'        'Mebane Orange Alamance
  ''$INCLUDE: 'UBMIDTWN.BI'        'Middletown VA
  ''$INCLUDE: 'UBMINCO.BI'         'Minco OK
  ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
  ''$INCLUDE: 'UBMTGILD.BI'        'Mt Gilead
  ''$INCLUDE: 'UBMTPLES.BI'        'Mt Plesent
  ''$INCLUDE: 'UBMURPHY.BI'        'Murphy
  ''$INCLUDE: 'UBNEWPRT.BI'        'Newport NC
  ''$INCLUDE: 'UBNLENOR.BI'        'North Lenoir
  ''$INCLUDE: 'UBNORLIN.BI'        'Norlina
  ''$INCLUDE: 'UBOAKCTY.BI'        'Oak City
  ''$INCLUDE: 'UBOAKBOR.BI'        'Oakboro
  ''$INCLUDE: 'UBOLDSDS.BI'        'Cleveland NC
  ''$INCLUDE: 'UBOLDSTD.BI'        'McBee SC    Old Standard 21 Line Bill
  ''$INCLUDE: 'UBPAGELN.BI'        'Pageland
  ''$INCLUDE: 'UBPEACHB.BI'        'Peachland
  ''$INCLUDE: 'UBPEMBRK.BI'        'Pembroke VA
  ''$INCLUDE: 'UBPENDTN.BI'        'Pendelton
  ''$INCLUDE: 'UBPINEBF.BI'        'Pinebluff
  ''$INCLUDE: 'UBPLYMTH.BI'        'Plymouth
  ''$INCLUDE: 'UBPOLKTN.BI'        'Polkton
  ''$INCLUDE: 'UBPOUND.BI'         'Pound VA
  ''$INCLUDE: 'UBPRINCE.BI'        'Princeton
  ''$INCLUDE: 'UBREMTON.BI'        'Remington
  ''$INCLUDE: 'UBROLES.BI'         'Rolesville
  ''$INCLUDE: 'UBRUSLCO.BI'        'Russell Co.
  ''$INCLUDE: 'UBROSEBL.BI'        'Roseboro
  ''$INCLUDE: 'UBRYAN.BI'        'Ryan OK
  ''$INCLUDE: 'UBSALEM.BI'         'Salem SC
  ''$INCLUDE: 'UBSALTVL.BI'        'Saltville
  ''$INCLUDE: 'UBSALUCO.BI'        'Saluda CO
  ''$INCLUDE: 'UBSAWMIL.BI'        'Sawmills
  ''$INCLUDE: 'UBSCOTTS.BI'        'Scottsburg
  ''$INCLUDE: 'UBSPENNC.BI'        'Spencer NC
  ''$INCLUDE: 'UBSPENTN.BI'        'Spencer TN
  ''$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
  ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
  ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
  ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
  '$INCLUDE: 'UBSUNSET.BI'        'Sunset Beech
  ''$INCLUDE: 'UBTALOGA.BI'        'Taloga
  ''$INCLUDE: 'UBTENNRD.BI'        'Tenn Ridge
  ''$INCLUDE: 'UBTIMSON.BI'        'Timmonsville SC
  ''$INCLUDE: 'UBTROUTM.BI'        'Troutman, nc
  ''$INCLUDE: 'UBTROY.BI'          'Troy
  ''$INCLUDE: 'UBTUSK.BI'          'Tuskaseigee NC
  ''$INCLUDE: 'UBTWINCT.BI'        'Twin City
  ''$INCLUDE: 'UBWADEBL.BI'        'Wade NC
  ''$INCLUDE: 'UBWAGRAM.BI'        'Wagram
  ''$INCLUDE: 'UBWHITKR.BI'        'Whitakers
  ''$INCLUDE: 'UBWHITLK.BI'        'White Lake
  ''$INCLUDE: 'UBWINDSR.BI'        'Windsor
  ''$INCLUDE: 'UBWINGAT.BI'        'Wingate
  ''$INCLUDE: 'UBWRIGHT.BI'        'Wrightsville Beech
  ''$INCLUDE: 'UBYONGBL.BI'        'Youngsville
  ''$INCLUDE: 'UBjoston.BI'        'Johnston
  ''$INCLUDE: 'UBlawrnc.BI'        'Lawrenceville
  ''$INCLUDE: 'UBnewstd.BI'        '
  ''$INCLUDE: 'UBYADKIN.BI'        'Yadkinville
  ''$INCLUDE: 'UBLUGOFW.bi'        '
  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Posted Bill Reprinting ", "UBBILLPP.PRN", 1, RetCode, 1
  END IF

Reprint1Exit:
  UBLog "OUT: Posted Bill Reprinting."
END SUB

SUB PrintAlignMask
  REDIM ScrnArray(0)
  CursorOff
  SaveScrn ScrnArray()
  BlockClear
  PrintRptFile "Utility Bill Alignment Mask ", "UBBLMASK.DAT", 1, RetCode, 4
  RestScrn ScrnArray()
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
    IF INSTR(UBSetUpRec(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
    END IF
  NEXT
  
  'Section to check for customer modifications
  TownName$ = UBSetUpRec(1).UTILNAME
  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF
  IF INSTR(TownName$, "GILES") THEN
    PSAFlag = True
  END IF
  IF INSTR(TownName$, "MOWAS") THEN
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
    CASE F5KEY
      PrintAlignMask
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
  
  PastDate$ = Form$(PastDFld, 0)
  
  'do bill printing here
  '**************************************************************************
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
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBILLS.PRN" FOR OUTPUT AS UBRpt
  
  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF
  
  UBLog "Printing utility bills to disk."
  ShowProcessingScrn "Creating Utility Bills."
  
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
    IF UBCustRec(1).BillTo = "O" THEN   'if they want to send the bill
      'temp copy the owners info to the customers rec
      GET UBOwn, CustAcctNo&, UBOwnerRec(1)
      OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
      UBCustRec(1).CustName = OName$
      UBCustRec(1).Addr1 = UBOwnerRec(1).Addr1
      UBCustRec(1).Addr2 = UBOwnerRec(1).Addr2
      UBCustRec(1).City = UBOwnerRec(1).City
      UBCustRec(1).State = UBOwnerRec(1).State
      UBCustRec(1).ZipCode = UBOwnerRec(1).ZipCode
    END IF
    
    Num2Print = UBCustRec(1).BILLCOPY
    IF Num2Print < 1 THEN Num2Print = 1
    
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
          
'IF DaysINRead > 59 THEN STOP

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
'NOTE: Maccelsfield bill has best code to parse out meter readings
          FOR BillCopies = 1 TO Num2Print
            ''$INCLUDE: 'NORWODBL.BI'        'Norwood
            ''$INCLUDE: 'UBALBERT.BI'        'ALBERTA
            ''$INCLUDE: 'UBALEXML.BI'        'ALEXANDER MILLS
            ''$INCLUDE: 'UBANDREW.BI'        'Andrews
            ''$INCLUDE: 'UBANSON.BI'         'Ansonville NC
            ''$INCLUDE: 'UBAPPLCH.BI'        'APPALACHIA VA
            ''$INCLUDE: 'UBAURORA.BI'        'like
            ''$INCLUDE: 'UBAUTRYV.BI'        'Autryville NC
            ''$INCLUDE: 'UBBEECH.BI2'        'like
            ''$INCLUDE: 'UBBETHEL.BI'        'like
            ''$INCLUDE: 'UBBLADEN.BI'        'Bladen
            ''$INCLUDE: 'UBBRANCH.BI'        'Branchville SC
            ''$INCLUDE: 'UBBROKNL.BI'        'Brookneal
            ''$INCLUDE: 'UBBROKFR.BI'        'BrookFord
            ''$INCLUDE: 'UBBRUNDG.BI'        'Brundage
            ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
            ''$INCLUDE: 'UBCALDWL.BI'        'Caldwell
            ''$INCLUDE: 'UBCASHON.BI'        'Cashion
            ''$INCLUDE: 'UBCELINA.BI'        'Celina
            ''$INCLUDE: 'UBCHESTR.BI'        'Chester
            ''$INCLUDE: 'UBCHILHO.BI'        'Chilhowe
            ''$INCLUDE: 'UBCONCRD.BI'        'Concord
            ''$INCLUDE: 'UBCONWAY.BI'        'Conway
            ''$INCLUDE: 'UBCREWBL.BI'        'Crew
            ''$INCLUDE: 'UBCedar.BI'         'Cedar Bluff VA
            ''$INCLUDE: 'UBDECATR.BI'        'Decatur
            ''$INCLUDE: 'UBDENTON.BI'        'Denton
            ''$INCLUDE: 'UBDUBLIN.BI'        'Dublin NC
            ''$INCLUDE: 'UBEHRHRT.BI'         'Ehrhardt SC
            ''$INCLUDE: 'UBELKTON.BI'         'Elkton
            ''$INCLUDE: 'UBELLEN.BI'         'Ellenboro NC
            ''$INCLUDE: 'UBELLORE.BI'        'Elloree
            ''$INCLUDE: 'UBFAITH.BI'         'Faith
            ''$INCLUDE: 'UBFOUROK.BI'        'Four Oaks NC
            ''$INCLUDE: 'UBGATECT.BI'        'Gate City
            ''$INCLUDE: 'UBGEORTN.BI'        'Georgetown
            ''$INCLUDE: 'UBGILBRT.BI'        'Gilbert Summit
            ''$INCLUDE: 'UBGILES.BI'         'GILES VA
            ''$INCLUDE: 'UBGORDON.BI'        'Gordonsville VA
            ''$INCLUDE: 'UBGRAYSN.BI'        'Grayson VA
            ''$INCLUDE: 'UBGRENVR.BI'        'Greenevers NC
            ''$INCLUDE: 'UBGROVER.BI'        'Grover NC  24 Line Special
            ''$INCLUDE: 'UBHALFAX.BI'        'Halifax VA
            ''$INCLUDE: 'UBHAMLET.BI'        'Hamlet
            ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
            ''$INCLUDE: 'UBHILBIL.BI'        'Hillsville VA
            ''$INCLUDE: 'UBHILDBR.BI'        'Hildabran
            ''$INCLUDE: 'UBHOLLY.BI'         'Holly springs
            ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
            ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
            ''$INCLUDE: 'UBJAMEST.BI'        'Jamestown
            ''$INCLUDE: 'UBJONSVL.BI'        'Jonesville NC
            ''$INCLUDE: 'UBJVILLE.BI'        'Johnsonville
            ''$INCLUDE: 'UBKELFRD.BI'        'Kelford NC
            ''$INCLUDE: 'UBKERSHA.bi'        'Kershaw
            ''$INCLUDE: 'UBLANDIS.BI'        'Landis,nc
            '$INCLUDE: 'UBLINCOL.BI'        'Lincoln County GA
            ''$INCLUDE: 'UBLINDEN.BI'        'Linden,nc
            ''$INCLUDE: 'UBLONGVW.BI'        'Long View nc
            ''$INCLUDE: 'UBLOUISA.BI'        'Louisa
            ''$INCLUDE: 'UBMACCFL.BI'        'Maccelfield
            ''$INCLUDE: 'UBMARSH.BI'         'Marshville, nc
            ''$INCLUDE: 'UBMANTEO.BI'         'Manteo NC
            ''$INCLUDE: 'UBMAYVIL.BI'        'Maysville
            ''$INCLUDE: 'UBMCMICK.BI'        'McCormick sc
            ''$INCLUDE: 'UBMEANVL.BI'        'Meanville SC
            ''$INCLUDE: 'UBMEBANE.BI'        'Mebane Orange Alamance
            ''$INCLUDE: 'UBMIDTWN.BI'        'Middletown VA
            ''$INCLUDE: 'UBMINCO.BI'         'Minco OK
            ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
            ''$INCLUDE: 'UBMTGILD.BI'        'Mt Gilead
            ''$INCLUDE: 'UBMTPLES.BI'        'Mt Plesent
            ''$INCLUDE: 'UBMURPHY.BI'        'Murphy
            ''$INCLUDE: 'UBNEWPRT.BI'        'Newport NC
            ''$INCLUDE: 'UBNLENOR.BI'        'North Lenoir
            ''$INCLUDE: 'UBNORLIN.BI'        'Norlina
            ''$INCLUDE: 'UBOAKCTY.BI'        'Oak City
            ''$INCLUDE: 'UBOAKBOR.BI'        'Oakboro
            ''$INCLUDE: 'UBOLDSDS.BI'        'Cleveland NC
            ''$INCLUDE: 'UBOLDSTD.BI'        'McBee SC    Old Standard 21 Line Bill
            ''$INCLUDE: 'UBPAGELN.BI'        'Pageland
            ''$INCLUDE: 'UBPEACHB.BI'        'Peachland
            ''$INCLUDE: 'UBPEMBRK.BI'        'Pembroke VA
            ''$INCLUDE: 'UBPENDTN.BI'        'Pendelton
            ''$INCLUDE: 'UBPINEBF.BI'        'Pinebluff
            ''$INCLUDE: 'UBPLYMTH.BI'        'Plymouth
            ''$INCLUDE: 'UBPOLKTN.BI'        'Polkton
            ''$INCLUDE: 'UBPOUND.BI'         'Pound VA
            ''$INCLUDE: 'UBPRINCE.BI'        'Princeton
            ''$INCLUDE: 'UBREMTON.BI'        'Remington
            ''$INCLUDE: 'UBROLES.BI'         'Rolesville
            ''$INCLUDE: 'UBRUSLCO.BI'        'Russell Co.
            ''$INCLUDE: 'UBROSEBL.BI'        'Roseboro
            ''$INCLUDE: 'UBRYAN.BI'        'Ryan OK
            ''$INCLUDE: 'UBSALEM.BI'         'Salem SC
            ''$INCLUDE: 'UBSALTVL.BI'        'Saltville
            ''$INCLUDE: 'UBSALUCO.BI'        'Saluda CO
            ''$INCLUDE: 'UBSAWMIL.BI'        'Sawmills
            ''$INCLUDE: 'UBSCOTTS.BI'        'Scottsburg
            ''$INCLUDE: 'UBSPENNC.BI'        'Spencer NC
            ''$INCLUDE: 'UBSPENTN.BI'        'Spencer TN
            ''$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
            ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
            ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
            ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
            ''$INCLUDE: 'UBSUNSET.BI'        'Sunset Beech
            ''$INCLUDE: 'UBTALOGA.BI'        'Taloga
            ''$INCLUDE: 'UBTENNRD.BI'        'Tenn Ridge
            ''$INCLUDE: 'UBTIMSON.BI'        'Timmonsville SC
            ''$INCLUDE: 'UBTROUTM.BI'        'Troutman, nc
            ''$INCLUDE: 'UBTROY.BI'          'Troy
            ''$INCLUDE: 'UBTUSK.BI'          'Tuskaseigee NC
            ''$INCLUDE: 'UBTWINCT.BI'        'Twin City
            ''$INCLUDE: 'UBWADEBL.BI'        'Wade NC
            ''$INCLUDE: 'UBWAGRAM.BI'        'Wagram
            ''$INCLUDE: 'UBWHITKR.BI'        'Whitakers
            ''$INCLUDE: 'UBWHITLK.BI'        'White Lake
            ''$INCLUDE: 'UBWINDSR.BI'        'Windsor
            ''$INCLUDE: 'UBWINGAT.BI'        'Wingate
            ''$INCLUDE: 'UBWRIGHT.BI'        'Wrightsville Beech
            ''$INCLUDE: 'UBYONGBL.BI'        'Youngsville
            ''$INCLUDE: 'UBjoston.BI'        'Johnston
            ''$INCLUDE: 'UBlawrnc.BI'        'Lawrenceville
            ''$INCLUDE: 'UBnewstd.BI'        '
            ''$INCLUDE: 'UBYADKIN.BI'        'Yadkinville
            ''$INCLUDE: 'UBLUGOFW.bi'        '
          NEXT
        END IF
      ELSE
        'mod for cleveland***
        IF UBBillRec(1).NONProfit = "Y" THEN
          GOSUB DOChurchTrans
        END IF
        '***
      END IF
    END IF

'    IF DidCnt > 1 THEN EXIT FOR

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
SkipEm:
  NEXT
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF
  
  IF FFFlag THEN
    PRINT #UBRpt, CHR$(12);
  END IF
  
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
    PrintRptFile "Utility Bill Printing ", "UBBILLS.PRN", 1, RetCode, 1
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
  
  Message$ = Form$(Msgf1, 0)
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
  UBLog "OUT: Bill Printing." + CRLF$
  
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
        RSET ZipIndex(ZipCnt).ZipCode = QPTrim$(LEFT$(UBCustRec(1).ZipCode, 5))
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
  PrnArray(PrnCnt).ZipCode = ZipIndex(1).ZipCode
  PrnArray(PrnCnt).RecNum = 1
  
  FOR ZCnt = 2 TO ZipCnt
    HadIt = False
    FOR PCnt = 1 TO PrnCnt
      IF PrnArray(PCnt).ZipCode = ZipIndex(ZCnt).ZipCode THEN
        PrnArray(PCnt).RecNum = PrnArray(PCnt).RecNum + 1
        HadIt = True
        EXIT FOR
      END IF
    NEXT
    IF NOT HadIt THEN
      PrnCnt = PrnCnt + 1
      REDIM PRESERVE PrnArray(1 TO PrnCnt) AS UBPostalIndexType
      PrnArray(PrnCnt).ZipCode = ZipIndex(ZCnt).ZipCode
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
    PRINT #UBRpt, PrnArray(Cnt).ZipCode, TAB(30); USING "#####"; PrnArray(Cnt).RecNum
    
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

SUB RePrintUtilBills
  
  UBLog " IN: Bill Reprinting."
  
  LPIFlag = False
  
  REDIM TempScrn(0)
  
  IF NOT Exist("UBBILLS.PRN") THEN
    UBLog "ERROR: NO PRN FILE."
    GOSUB ShowErrScrn
    GOTO ExitRePrintBill
  END IF
  
  IF NOT ChkBillFile% THEN
    UBLog "ERROR: NO BILL FILE."
    GOSUB ShowErrScrn
    GOTO ExitRePrintBill
  END IF
  
  NotBeenDone = True
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF
  
  FOR ThisRevCnt = 1 TO 15
    IF INSTR(UBSetUpRec(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
    END IF
  NEXT
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  'Section to check for customer modifications
  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF
  
  ShowProcessingScrn "Indexing Billing Information."
  QPrintRC "      Reading Bill Records       ", 11, 25, -1
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  NumOfRecs = LOF(UBBill) \ UBBillRecLen
  REDIM RePrintIdx(1 TO NumOfRecs) AS RePrintIndexType
  FOR Cnt = 1 TO NumOfRecs
    GET UBBill, Cnt, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
      RePrintIdx(Cnt).BillNum = UBBillRec(1).BillNumber
      RePrintIdx(Cnt).BillRec = Cnt
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  CLOSE UBBill
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  UBFile = FREEFILE
  OPEN "UBPINFON.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  GET #UBFile, 1, BillInfoRec(1)
  CLOSE
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  SortT RePrintIdx(1), NumOfRecs, 0, 8, 0, -2
  
  '--Initialize the form name array
  'LibName$ = "UB"
  LibName$ = "UBPRNBIL"
  ScrnName$ = "PRNTBIL2"
  MActionRow = 20
  
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
  BillOFld = FldNum%("PRNORDER", Fld())
  
  Msgf1 = FldNum%("MSGLINE1", Fld())
  Msgf2 = FldNum%("MSGLINE2", Fld())
  Msgf3 = FldNum%("MSGLINE3", Fld())
  Msgf4 = FldNum%("MSGLINE4", Fld())
  
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  FirstTime = True
  
  BlockClear
  
  DisplayPrnScrn ScrnName$
  
  'mod for Indian Trail
  IF IndianFlag THEN
    FOR Cnt = 1 TO 3
      QPrintRC " Description " + QPTrim$(STR$(Cnt)) + ":", Cnt + 15, 20, -1
    NEXT
  END IF
  'end Indian Trail mod
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      GOSUB FillForm
      FirstTime = False
    END IF
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE F5KEY
      PrintAlignMask
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
    UBLog "Aborted BILL Reprinting."
    GOTO ExitRePrintBill
  END IF
  
  ShowProcessingScrn "Creating Utility Bills."
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBILLSR.PRN" FOR OUTPUT AS UBRpt
  
  Cnt = 0
  DO
    Cnt = Cnt + 1
    IF Cnt > NumOfRecs THEN EXIT DO
    ThisBillNum = RePrintIdx(Cnt).BillNum
    
    IF ThisBillNum >= FirstBill AND ThisBillNum <= LastBill THEN
      PrintedCnt = ThisBillNum
      CustAcctNo& = RePrintIdx(Cnt).BillRec

      GET UBBill, CustAcctNo&, UBBillRec(1)
      GET UBCust, CustAcctNo&, UBCustRec(1)
'121598 Added bill to owner
      IF UBCustRec(1).BillTo = "O" THEN   'if bill owner, temp copy info
        GET UBOwn, CustAcctNo&, UBOwnerRec(1)
        OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
        UBCustRec(1).CustName = OName$
        UBCustRec(1).Addr1 = UBOwnerRec(1).Addr1
        UBCustRec(1).Addr2 = UBOwnerRec(1).Addr2
        UBCustRec(1).City = UBOwnerRec(1).City
        UBCustRec(1).State = UBOwnerRec(1).State
        UBCustRec(1).ZipCode = UBOwnerRec(1).ZipCode
      END IF
      
      BillDate$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
      PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)

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
      
      TotalTax# = 0
      FOR TaxCnt = 1 TO MaxRevsCnt
        TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
      NEXT

      DidADraftFlag = False
      IF UseDraftFlag AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag THEN
        DidADraftFlag = True
      END IF
      
      Num2Print = UBCustRec(1).BILLCOPY
      IF Num2Print < 1 THEN Num2Print = 1
'NOTE: Maccelsfield bill has best code to parse out meter readings
      FOR BillCopies = 1 TO Num2Print
        ''$INCLUDE: 'NORWODBL.BI'        'Norwood
        ''$INCLUDE: 'UBALBERT.BI'        'ALBERTA
        ''$INCLUDE: 'UBALEXML.BI'        'ALEXANDER MILLS
        ''$INCLUDE: 'UBANDREW.BI'        'Andrews
        ''$INCLUDE: 'UBANSON.BI'         'Ansonville NC
        ''$INCLUDE: 'UBAPPLCH.BI'        'APPALACHIA VA
        ''$INCLUDE: 'UBAURORA.BI'        'like
        ''$INCLUDE: 'UBAUTRYV.BI'        'Autryville NC
        ''$INCLUDE: 'UBBEECH.BI2'        'like
        ''$INCLUDE: 'UBBETHEL.BI'        'like
        ''$INCLUDE: 'UBBLADEN.BI'        'Bladen
        ''$INCLUDE: 'UBBRANCH.BI'        'Branchville SC
        ''$INCLUDE: 'UBBROKNL.BI'        'Brookneal
        ''$INCLUDE: 'UBBROKFR.BI'        'BrookFord
        ''$INCLUDE: 'UBBRUNDG.BI'        'Brundage
        ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
        ''$INCLUDE: 'UBCALDWL.BI'        'Caldwell
        ''$INCLUDE: 'UBCASHON.BI'        'Cashion
        ''$INCLUDE: 'UBCELINA.BI'        'Celina
        ''$INCLUDE: 'UBCHESTR.BI'        'Chester
        ''$INCLUDE: 'UBCHILHO.BI'        'Chilhowe
        ''$INCLUDE: 'UBCONCRD.BI'        'Concord
        ''$INCLUDE: 'UBCONWAY.BI'        'Conway
        ''$INCLUDE: 'UBCREWBL.BI'        'Crew
        ''$INCLUDE: 'UBCedar.BI'         'Cedar Bluff VA
        ''$INCLUDE: 'UBDECATR.BI'        'Decatur
        ''$INCLUDE: 'UBDENTON.BI'        'Denton
        ''$INCLUDE: 'UBDUBLIN.BI'        'Dublin NC
        ''$INCLUDE: 'UBEHRHRT.BI'         'Ehrhardt SC
        ''$INCLUDE: 'UBELKTON.BI'         'Elkton
        ''$INCLUDE: 'UBELLEN.BI'         'Ellenboro NC
        ''$INCLUDE: 'UBELLORE.BI'        'Elloree
        ''$INCLUDE: 'UBFAITH.BI'         'Faith
        ''$INCLUDE: 'UBFOUROK.BI'        'Four Oaks NC
        ''$INCLUDE: 'UBGATECT.BI'        'Gate City
        ''$INCLUDE: 'UBGEORTN.BI'        'Georgetown
        ''$INCLUDE: 'UBGILBRT.BI'        'Gilbert Summit
        ''$INCLUDE: 'UBGILES.BI'         'GILES VA
        ''$INCLUDE: 'UBGORDON.BI'        'Gordonsville VA
        ''$INCLUDE: 'UBGRAYSN.BI'        'Grayson VA
        ''$INCLUDE: 'UBGRENVR.BI'        'Greenevers NC
        ''$INCLUDE: 'UBGROVER.BI'        'Grover NC  24 Line Special
        ''$INCLUDE: 'UBHALFAX.BI'        'Halifax VA
        ''$INCLUDE: 'UBHAMLET.BI'        'Hamlet
        ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
        ''$INCLUDE: 'UBHILBIL.BI'        'Hillsville VA
        ''$INCLUDE: 'UBHILDBR.BI'        'Hildabran
        ''$INCLUDE: 'UBHOLLY.BI'         'Holly springs
        ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
        ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
        ''$INCLUDE: 'UBJAMEST.BI'        'Jamestown
        ''$INCLUDE: 'UBJONSVL.BI'        'Jonesville NC
        ''$INCLUDE: 'UBJVILLE.BI'        'Johnsonville
        ''$INCLUDE: 'UBKELFRD.BI'        'Kelford NC
        ''$INCLUDE: 'UBKERSHA.bi'        'Kershaw
        ''$INCLUDE: 'UBLANDIS.BI'        'Landis,nc
        '$INCLUDE: 'UBLINCOL.BI'        'Lincoln County GA
        ''$INCLUDE: 'UBLINDEN.BI'        'Linden,nc
        ''$INCLUDE: 'UBLONGVW.BI'        'Long View nc
        ''$INCLUDE: 'UBLOUISA.BI'        'Louisa
        ''$INCLUDE: 'UBMACCFL.BI'        'Maccelfield
        ''$INCLUDE: 'UBMARSH.BI'         'Marshville, nc
        ''$INCLUDE: 'UBMANTEO.BI'         'Manteo NC
        ''$INCLUDE: 'UBMAYVIL.BI'        'Maysville
        ''$INCLUDE: 'UBMCMICK.BI'        'McCormick sc
        ''$INCLUDE: 'UBMEANVL.BI'        'Meanville SC
        ''$INCLUDE: 'UBMEBANE.BI'        'Mebane Orange Alamance
        ''$INCLUDE: 'UBMIDTWN.BI'        'Middletown VA
        ''$INCLUDE: 'UBMINCO.BI'         'Minco OK
        ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
        ''$INCLUDE: 'UBMTGILD.BI'        'Mt Gilead
        ''$INCLUDE: 'UBMTPLES.BI'        'Mt Plesent
        ''$INCLUDE: 'UBMURPHY.BI'        'Murphy
        ''$INCLUDE: 'UBNEWPRT.BI'        'Newport NC
        ''$INCLUDE: 'UBNLENOR.BI'        'North Lenoir
        ''$INCLUDE: 'UBNORLIN.BI'        'Norlina
        ''$INCLUDE: 'UBOAKCTY.BI'        'Oak City
        ''$INCLUDE: 'UBOAKBOR.BI'        'Oakboro
        ''$INCLUDE: 'UBOLDSDS.BI'        'Cleveland NC
        ''$INCLUDE: 'UBOLDSTD.BI'        'McBee SC    Old Standard 21 Line Bill
        ''$INCLUDE: 'UBPAGELN.BI'        'Pageland
        ''$INCLUDE: 'UBPEACHB.BI'        'Peachland
        ''$INCLUDE: 'UBPEMBRK.BI'        'Pembroke VA
        ''$INCLUDE: 'UBPENDTN.BI'        'Pendelton
        ''$INCLUDE: 'UBPINEBF.BI'        'Pinebluff
        ''$INCLUDE: 'UBPLYMTH.BI'        'Plymouth
        ''$INCLUDE: 'UBPOLKTN.BI'        'Polkton
        ''$INCLUDE: 'UBPOUND.BI'         'Pound VA
        ''$INCLUDE: 'UBPRINCE.BI'        'Princeton
        ''$INCLUDE: 'UBREMTON.BI'        'Remington
        ''$INCLUDE: 'UBROLES.BI'         'Rolesville
        ''$INCLUDE: 'UBRUSLCO.BI'        'Russell Co.
        ''$INCLUDE: 'UBROSEBL.BI'        'Roseboro
        ''$INCLUDE: 'UBRYAN.BI'        'Ryan OK
        ''$INCLUDE: 'UBSALEM.BI'         'Salem SC
        ''$INCLUDE: 'UBSALTVL.BI'        'Saltville
        ''$INCLUDE: 'UBSALUCO.BI'        'Saluda CO
        ''$INCLUDE: 'UBSAWMIL.BI'        'Sawmills
        ''$INCLUDE: 'UBSCOTTS.BI'        'Scottsburg
        ''$INCLUDE: 'UBSPENNC.BI'        'Spencer NC
        ''$INCLUDE: 'UBSPENTN.BI'        'Spencer TN
        ''$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
        ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
        ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
        ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
        ''$INCLUDE: 'UBSUNSET.BI'        'Sunset Beech
        ''$INCLUDE: 'UBTALOGA.BI'        'Taloga
        ''$INCLUDE: 'UBTENNRD.BI'        'Tenn Ridge
        ''$INCLUDE: 'UBTIMSON.BI'        'Timmonsville SC
        ''$INCLUDE: 'UBTROUTM.BI'        'Troutman, nc
        ''$INCLUDE: 'UBTROY.BI'          'Troy
        ''$INCLUDE: 'UBTUSK.BI'          'Tuskaseigee NC
        ''$INCLUDE: 'UBTWINCT.BI'        'Twin City
        ''$INCLUDE: 'UBWADEBL.BI'        'Wade NC
        ''$INCLUDE: 'UBWAGRAM.BI'        'Wagram
        ''$INCLUDE: 'UBWHITKR.BI'        'Whitakers
        ''$INCLUDE: 'UBWHITLK.BI'        'White Lake
        ''$INCLUDE: 'UBWINDSR.BI'        'Windsor
        ''$INCLUDE: 'UBWINGAT.BI'        'Wingate
        ''$INCLUDE: 'UBWRIGHT.BI'        'Wrightsville Beech
        ''$INCLUDE: 'UBYONGBL.BI'        'Youngsville
        ''$INCLUDE: 'UBjoston.BI'        'Johnston
        ''$INCLUDE: 'UBlawrnc.BI'        'Lawrenceville
        ''$INCLUDE: 'UBnewstd.BI'        '
        ''$INCLUDE: 'UBYADKIN.BI'        'Yadkinville
        ''$INCLUDE: 'UBLUGOFW.bi'        '
      NEXT
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT DO
    END IF
    ShowPctComp Cnt, NumOfRecs
  LOOP
  
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF
  
zzgetout:
  
  IF FFFlag THEN
    PRINT #UBRpt, CHR$(12);
  END IF
  
  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, RePrintIdx
  
  IF NOT AbortFlag THEN
    PrintRptFile "Utility Bill Printing ", "UBBILLSR.PRN", 1, RetCode, 0
  ELSE
    UBLog "ABORTED: Bill Reprinting."
  END IF
  
ExitRePrintBill:
  UBLog "OUT: Bill Reprinting." + CRLF$
  EXIT SUB
  
CheckFields:
  FirstBill = Value#(Form$(FrstBFld, 0), ErrCode%)
  LastBill = Value#(Form$(LastBFld, 0), ErrCode%)
  IF FirstBill > LastBill THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADBILNO"
    WaitForAction
    RestScrn TempScrn()
  ELSE
    ReqFldsOK = True
  END IF
  RETURN
  
FillForm:
  
  LSET Form$(FrstBFld, 0) = QPTrim$(STR$(BillInfoRec(1).FrstBill))
  LSET Form$(LastBFld, 0) = QPTrim$(STR$(BillInfoRec(1).LastBill))
  LSET Form$(BillDFld, 0) = Num2Date$(BillInfoRec(1).BillDate)
  LSET Form$(PastDFld, 0) = Num2Date$(BillInfoRec(1).PastDate)
  
  IF BillInfoRec(1).PRDate > 0 THEN
    LSET Form$(PRDateFld, 0) = Num2Date$(BillInfoRec(1).PRDate)
  END IF
  IF BillInfoRec(1).CRDate > 0 THEN
    LSET Form$(CRDateFld, 0) = Num2Date$(BillInfoRec(1).CRDate)
  END IF
  IF BillInfoRec(1).DrftDate > 0 THEN
    LSET Form$(DraftDFld, 0) = Num2Date$(BillInfoRec(1).DrftDate)
  END IF
  
  LSET Form$(BillOFld, 0) = BillInfoRec(1).PrnOrder
  LSET Form$(Msgf1, 0) = BillInfoRec(1).MsgLine1
  LSET Form$(Msgf2, 0) = BillInfoRec(1).MsgLine2
  LSET Form$(Msgf3, 0) = BillInfoRec(1).MsgLine3
  LSET Form$(Msgf4, 0) = BillInfoRec(1).MsgLine4
  
  Message$ = QPTrim$(BillInfoRec(1).MsgLine1)
  Msg2$ = QPTrim$(BillInfoRec(1).MsgLine2)
  Msg3$ = QPTrim$(BillInfoRec(1).MsgLine3)
  Msg4$ = QPTrim$(BillInfoRec(1).MsgLine4)
  
  Action = 2
  RETURN
  
ShowErrScrn:
  CursorOff
  BlockClear
  DisplayUBScrn "NOTPRNTD"
  WaitForAction
  RETURN
END SUB

