DECLARE SUB PrintAlignMask ()
DEFINT A-Z

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
    '$INCLUDE: 'UBDENTON.BI'        'Denton
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
    '$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
    ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
    ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
    ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
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

SUB PrintAlignMask
  REDIM ScrnArray(0)
  CursorOff
  SaveScrn ScrnArray()
  BlockClear
  PrintRptFile "Utility Bill Alignment Mask ", "UBBLMASK.DAT", 1, RetCode, 4
  RestScrn ScrnArray()
END SUB

