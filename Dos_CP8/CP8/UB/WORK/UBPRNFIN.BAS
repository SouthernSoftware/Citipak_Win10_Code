DEFINT A-Z
DECLARE SUB DisplayPrnScrn (ScrnName$)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB MakeZipCodeIndex (IdxText$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubowner.bi'
  '$INCLUDE: 'ubpinfo.bi'
  
  CONST False = 0, True = NOT False
  
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Print ALL Final Utility Bills "
  MChoice$(2) = "Print-Reprint Selected Bills"
  MChoice$(3) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Final Bill Printing Menu ", Cnf
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
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  RUN "UBFINBIL"

SUB DisplayPrnScrn (ScrnName$)
  LibFile2Scrn "UBPRNBIL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB PrintAlignMask
  REDIM ScrnArray(0)
  CursorOff
  SaveScrn ScrnArray()
  BlockClear
  PrintRptFile "Utility Bill Alignment Mask ", "UBBLMASK.DAT", 1, RetCode, 4
  RestScrn ScrnArray()
END SUB

SUB PrintUtilBills
  
  OPEN "UBDEPFLG.DAT" FOR RANDOM SHARED AS #1 LEN = 2
  GET #1, , UseDepositFlag
  CLOSE #1
  
  SHARED Choice$()
  
  LPIFlag = False
  
  IF NOT Exist(UBFinBillsFile) THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NON2PRNT"
    WaitForAction
    GOTO ExitPrintBill
  END IF
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  REDIM IndexArray(1) AS UBCustIndexRecType
  
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
  IF INSTR(TownName$, "MOWAS") THEN
    MowFlag = True
  END IF
  
  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF
  
  REDIM Choice$(0 TO 6, 0)
  
  Choice$(1, 0) = "Customer Name Order"
  Choice$(2, 0) = "Account Number Order"
  'Choice$(3, 0) = "Location Number Order"
  Choice$(3, 0) = "Postal Carrier Route Order"
  Choice$(4, 0) = "ZipCode Order"
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    Choice$(6, 0) = "Sequence Number Order"
  END IF
  
  LibName$ = "UBPRNBIL"
  
  ScrnName$ = "PRNBILF1"
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
  
'  FrstBFld = FldNum%("FRSTBILL", Fld())
'  LastBFld = FldNum%("LASTBILL", Fld())
  
  BillDFld = FldNum%("BILLDATE", Fld())
  PastDFld = FldNum%("PASTDATE", Fld())
  
  PRDateFld = FldNum%("PRDATE", Fld())
  CRDateFld = FldNum%("CRDATE", Fld())
  
  DraftDFld = FldNum%("DRFTDATE", Fld())
  BillOFld = FldNum%("PRNORDER", Fld())
  DepositFld = FldNum%("APPLYDEP", Fld())
  
  MsgF1 = FldNum%("MSGLINE1", Fld())
  MsgF2 = FldNum%("MSGLINE2", Fld())
  MsgF3 = FldNum%("MSGLINE3", Fld())
  MsgF4 = FldNum%("MSGLINE4", Fld())
  
  Fld(MsgF1).Protected = True
  
  Today$ = DATE$
  
  Action = 1
  Frm(1).StayOnField = True
  
  FirstTime = True
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
      LSET Form$(BillDFld, 0) = Today$
      LSET Form$(PastDFld, 0) = Today$
      LSET Form$(DraftDFld, 0) = Today$
      SELECT CASE UseDepositFlag
      CASE True
        LSET Form$(DepositFld, 0) = "Y"
      CASE False
        LSET Form$(DepositFld, 0) = "N"
      END SELECT
      
      LSET Form$(MsgF1, 0) = "Final Billing"
      
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
  
  IF AbortFlag THEN GOTO ExitPrintBill
  
  SELECT CASE BillOrder$
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    OKFlag = True
  CASE "A"
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  CASE "P", "Z"
    IF BillOrder$ = "P" THEN
      IdxTypeText$ = "Postal Route"
      MakePostalIndex IdxTypeText$
    ELSE
      IdxTypeText$ = "Zip-Code"
      MakeMowZipCodeIndex IdxTypeText$
    END IF
    IndexName$ = TempIndexName
    OKFlag = True
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
  FinalFlag = True
  
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
  
  BillInfoRec(1).MsgLine1 = Form$(MsgF1, 0)
  BillInfoRec(1).MsgLine2 = Form$(MsgF2, 0)
  BillInfoRec(1).MsgLine3 = Form$(MsgF3, 0)
  BillInfoRec(1).MsgLine4 = Form$(MsgF4, 0)
  
  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen

  UBRpt = FREEFILE
  OPEN "UBFBILLS.PRN" FOR OUTPUT AS UBRpt
  
  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF
  
  ShowProcessingScrn "Creating Utility Bills."
  
  '-----------------------------------------
  PrintedCnt = 0
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      CustAcctNo& = Cnt
    ELSE
      CustAcctNo& = IndexArray(Cnt).RecNum
    END IF
    
    GET UBCust, CustAcctNo&, UBCustRec(1)
    
    Num2Print = UBCustRec(1).BILLCOPY
    IF Num2Print < 1 THEN Num2Print = 1

    GET UBBill, CustAcctNo&, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
'121598 Finished adding bill to owner
      IF UBCustRec(1).BillTo = "O" THEN   'if they want to send bill to owner
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
        
        IF ApplyDepFlag$ = "Y" THEN
          CDeposit# = UBCustRec(1).DepositAmt
        ELSE
          CDeposit# = 0
        END IF
        
        UBBillRec(1).ApplyDepFlag = ApplyDepFlag$
        
        PUT UBBill, CustAcctNo&, UBBillRec(1)
        
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
          ''$INCLUDE: 'NORWODBL.BI'        'Norwood
          ''$INCLUDE: 'UBALBERT.BI'        'ALBERTA
          ''$INCLUDE: 'UBALEXML.BI'        'ALEXANDER MILLS
          ''$INCLUDE: 'UBANDREW.BI'         'Andrews
          ''$INCLUDE: 'UBANSON.BI'         'Ansonville NC
          ''$INCLUDE: 'UBAPPLCH.BI'        'APPALACHIA VA
          ''$INCLUDE: 'UBARLNGT.BI'        'Arlington
          ''$INCLUDE: 'UBAURORA.BI'        'like
          ''$INCLUDE: 'UBAUTRYV.BI'        'Autryville NC
          ''$INCLUDE: 'UBBEAVER.BI'        'Beaver OK
          ''$INCLUDE: 'UBBEECH.BI2'        'like
          ''$INCLUDE: 'UBBETHEL.BI'        'like
          ''$INCLUDE: 'UBBLADEN.BI'        'Bladen
          ''$INCLUDE: 'UBBRANCH.BI'        'Branchville SC
          ''$INCLUDE: 'UBBROKNL.BI'        'Brookneal
          ''$INCLUDE: 'UBBROKFR.BI'        'BrookFord
          ''$INCLUDE: 'UBBRUNDG.BI'        'Brundage
          ''$INCLUDE: 'UBBRUNWK.BI'        'Brunswick
          ''$INCLUDE: 'UBBUCKSP.BI'        'Bucksport Water
          ''$INCLUDE: 'UBBURKVL.BI'        'Burkville VA
          ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
          ''$INCLUDE: 'UBCALDWL.BI'        'Caldwell
          '$INCLUDE: 'UBCANDOR.BI'        'Candor
          ''$INCLUDE: 'UBCARBCH.BI'        'Carolina Beech
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
          ''$INCLUDE: 'UBFOUNTN.BI'        'Fountain
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
          ''$INCLUDE: 'UBHAMTON.BI'        'Hamilton
          ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
          ''$INCLUDE: 'UBHILBIL.BI'        'Hillsville VA
          ''$INCLUDE: 'UBHILDBR.BI'        'Hildabran
          ''$INCLUDE: 'UBHOLLY.BI'         'Holly springs
          ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
          ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
          ''$INCLUDE: 'UBJAMEST.BI'        'Jamestown
          ''$INCLUDE: 'UBJONSVL.BI'        'Jonesville NC
          ''$INCLUDE: 'UBJVILLE.BI'        'Johnsonville,  Morvan
          ''$INCLUDE: 'UBKELFRD.BI'        'Kelford NC
          ''$INCLUDE: 'UBKERSHA.bi'        'Kershaw
          ''$INCLUDE: 'UBLANDIS.BI'        'Landis,nc
          ''$INCLUDE: 'UBLEEPSA.BI'        'Lee PSA
          ''$INCLUDE: 'UBLINCOL.BI'        'Lincoln County GA
          ''$INCLUDE: 'UBLINDEN.BI'        'Linden,nc
          ''$INCLUDE: 'UBLONGVW.BI'        'Long View nc
          ''$INCLUDE: 'UBLOUISA.BI'        'Louisa
          ''$INCLUDE: 'UBMACCFL.BI'        'Maccelfield
          ''$INCLUDE: 'UBMAGNOL.BI'         'Magnolia NC
          ''$INCLUDE: 'UBMARSH.BI'         'Marshville, nc
          ''$INCLUDE: 'UBMANTEO.BI'         'Manteo NC
          ''$INCLUDE: 'UBMAYVIL.BI'        'Maysville
          ''$INCLUDE: 'UBMCMICK.BI'        'McCormick sc
          ''$INCLUDE: 'UBMEANVL.BI'        'Meanville SC
          ''$INCLUDE: 'UBMEBANE.BI'        'Mebane Orange Alamance
          ''$INCLUDE: 'UBMIDTWN.BI'        'Middletown VA
          ''$INCLUDE: 'UBMINCO.BI'         'Minco OK
          ''$INCLUDE: 'UBMONTRS.BI'         'Montross
          ''$INCLUDE: 'UBMORVAN.BI'        'Morven
          ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
          ''$INCLUDE: 'UBMTGILD.BI'        'Mt Gilead
          ''$INCLUDE: 'UBMTPLES.BI'        'Mt Plesent
          ''$INCLUDE: 'UBMURPHY.BI'        'Murphy
          ''$INCLUDE: 'UBNEWPRT.BI'        'Newport NC
          ''$INCLUDE: 'UBNLENOR.BI'         'North Lenoir
          ''$INCLUDE: 'UBNORLIN.BI'        'Norlina
          ''$INCLUDE: 'UBNORWST.BI'        'Northwest
          ''$INCLUDE: 'UBOAKCTY.BI'        'Oak City
          ''$INCLUDE: 'UBOAKBOR.BI'        'Oakboro
          ''$INCLUDE: 'UBOLDSDS.BI'        'Cleveland NC
          ''$INCLUDE: 'UBOLDSTD.BI'        'McBee SC    Old Standard 21 Line Bill
          ''$INCLUDE: 'UBPEACHB.BI'        'Peachland
          ''$INCLUDE: 'UBPEMBRK.BI'        'Pembroke VA
          ''$INCLUDE: 'UBPAGELN.BI'        'Pageland
          ''$INCLUDE: 'UBPENDTN.BI'        'Pendelton
          ''$INCLUDE: 'UBPINEBF.BI'        'Pinebluff
          ''$INCLUDE: 'UBPINKHL.BI'        'PinkHill
          ''$INCLUDE: 'UBPLYMTH.BI'        'Plymouth
          ''$INCLUDE: 'UBPOLKTN.BI'        'Polkton
          ''$INCLUDE: 'UBPOUND.BI'         'Pound VA
          ''$INCLUDE: 'UBPRINCE.BI'        'Princeton
          ''$INCLUDE: 'UBREMTON.BI'        'Remington
          ''$INCLUDE: 'UBRICHLN.BI'        'Richland
          ''$INCLUDE: 'UBROLES.BI'         'Rolesville
          ''$INCLUDE: 'UBRUSLCO.BI'        'Russell Co.
          ''$INCLUDE: 'UBROSEBL.BI'        'Roseboro
          ''$INCLUDE: 'UBRYAN.BI'        'Ryan OK
          ''$INCLUDE: 'UBSALEM.BI'         'Salem SC
          ''$INCLUDE: 'UBSALTVL.BI'        'Saltville
          ''$INCLUDE: 'UBSALUCO.BI'        'Saluda CO
          ''$INCLUDE: 'UBSAWMIL.BI'        'Sawmills
          ''$INCLUDE: 'UBSCOTTS.BI'        'Scottsburg
          ''$INCLUDE: 'UBSEDGE.BI'        'Sedgefield
          ''$INCLUDE: 'UBSPENNC.BI'        'Spencer NC
          ''$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
          ''$INCLUDE: 'UBSTEDBL.BI'        'Stedman
          ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
          ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
          ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
          ''$INCLUDE: 'UBSUNSET.BI'        'Sunset Beech
          ''$INCLUDE: 'UBSUMTON.BI'        'Summerton
          ''$INCLUDE: 'UBTALOGA.BI'        'Taloga
          ''$INCLUDE: 'UBTENNRD.BI'        'Tenn Ridge
          ''$INCLUDE: 'UBTENNIL.BI'        'Tennil GA
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
          ''$INCLUDE: 'UBWARSAW.BI'        '
          ''$INCLUDE: 'UBYADKIN.BI'        'Yadkinville
          ''$INCLUDE: 'UBLUGOFW.bi'        '
          ''$INCLUDE: 'UBNLASER.BI'        '(new standard style) 17" laser
        NEXT
      END IF
    END IF
    
SkipEm:
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF

  IF FFFlag THEN
    PRINT #UBRpt, CHR$(12);
  END IF

  CLOSE
  
  BillInfoRec(1).FrstBill = 1
  BillInfoRec(1).LastBill = PrintedCnt
  BillInfoRec(1).BillDate = Date2Num(Form$(BillDFld, 0))
  BillInfoRec(1).PastDate = Date2Num(Form$(PastDFld, 0))
  BillInfoRec(1).PRDate = Date2Num(Form$(PRDateFld, 0))
  BillInfoRec(1).CRDate = Date2Num(Form$(CRDateFld, 0))
  BillInfoRec(1).DrftDate = Date2Num(Form$(DraftDFld, 0))
  BillInfoRec(1).PrnOrder = QPTrim$(Form$(BillOFld, 0))
  
  BillInfoRec(1).MsgLine1 = Form$(MsgF1, 0)
  BillInfoRec(1).MsgLine2 = Form$(MsgF2, 0)
  BillInfoRec(1).MsgLine3 = Form$(MsgF3, 0)
  BillInfoRec(1).MsgLine4 = Form$(MsgF4, 0)
  
  UBFile = FREEFILE
  OPEN "UBPINFOF.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  PUT #UBFile, 1, BillInfoRec(1)
  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, BillInfoRec, IndexArray
  
  IF NOT AbortFlag THEN
    PrintRptFile "Final Utility Bill Printing ", "UBFBILLS.PRN", 1, RetCode, 1
  END IF
  
  '**************************************************************************
  'end bill printing
  GOTO ExitPrintBill:
  '******************
  
  '******************
CheckReqFields:
  
  'FBillNO& = VAL(Form$(FrstBFld, 0))
  'LBillNO& = VAL(Form$(LastBFld, 0))

  BillDate = Date2Num(Form$(BillDFld, 0))
  BillDate$ = Num2Date$(BillDate)
  
  '  BilledDate$ = Form$(BillDFld, 0)
  PastDate = Date2Num(Form$(PastDFld, 0))
  
  DraftDate = Date2Num(Form$(DraftDFld, 0))
  DueDate$ = Form$(PastDFld, 0)
  BillOrder$ = QPTrim$(LEFT$(Form$(BillOFld, 0), 1))
  ApplyDepFlag$ = Form$(DepositFld, 0)
  
  Message$ = Form$(MsgF1, 0)
  
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
  ELSEIF LEN(BillOrder$) = 0 THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "NOBORDER"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 6
    '  ELSEIF ((FBillNO& <= 0) OR (LBillNO& < FBillNO&)) THEN
    '    SaveScrn ScrnArray()
    '    DisplayUBScrn "NOSTRTNO"
    '    WaitForAction
    '    RestScrn ScrnArray()
    '    Frm(1).FldNo = FrstBFld
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

SUB RePrintUtilBills
  'CustAcctNo&
  LPIFlag = False
  FinalFlag = True
  REDIM TempScrn(0)
  
  IF NOT Exist("UBFBILLS.PRN") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NOTPRNTD"
    WaitForAction
    GOTO ExitRePrintBill
  END IF
  
  OPEN "UBDEPFLG.DAT" FOR RANDOM SHARED AS #1 LEN = 2
  GET #1, , UseDepositFlag
  CLOSE #1
  
  SELECT CASE UseDepositFlag
  CASE True
    Deposit$ = "Y"
    ApplyDepFlag$ = "Y"
  CASE ELSE
    Deposit$ = "N"
    ApplyDepFlag$ = " "
  END SELECT
  NotBeenDone = True
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
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
  
  ShowProcessingScrn "Indexing Billing File."
  QPrintRC "      Reading Bill Records       ", 11, 25, -1
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
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
  OPEN "UBPINFOF.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  GET #UBFile, 1, BillInfoRec(1)
  CLOSE
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT RePrintIdx(1), NumOfRecs, 0, 8, 0, -2
  
  '--Initialize the form name array
  LibName$ = "UBPRNBIL"
  ScrnName$ = "PRNBILF2"
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
  DepositFld = FldNum%("APPLYDEP", Fld())
  
  MsgF1 = FldNum%("MSGLINE1", Fld())
  MsgF2 = FldNum%("MSGLINE2", Fld())
  MsgF3 = FldNum%("MSGLINE3", Fld())
  MsgF4 = FldNum%("MSGLINE4", Fld())
  
  Message$ = "FINAL BILLING"
  
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
  
  IF AbortFlag GOTO ExitRePrintBill:
  
  ShowProcessingScrn "Creating Utility Bills."
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen
  
  UBRpt = FREEFILE
  OPEN "UBFBILLR.PRN" FOR OUTPUT AS UBRpt

  Cnt = 0
  DO
    Cnt = Cnt + 1
    IF Cnt > NumOfRecs THEN EXIT DO
    ThisBillNum = RePrintIdx(Cnt).BillNum
    IF ThisBillNum >= FirstBill AND ThisBillNum <= LastBill THEN
      PrintedCnt = ThisBillNum
      GET UBBill, RePrintIdx(Cnt).BillRec, UBBillRec(1)
      GET UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
      CustAcctNo& = UBBillRec(1).CustAcctNo
'121598 Finished adding bill to owner
      IF UBCustRec(1).BillTo = "O" THEN   'if they want to send bill to owner
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
      
      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)
      
      DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
      
      IF ApplyDepFlag$ = "Y" THEN
        CDeposit# = UBCustRec(1).DepositAmt
      ELSE
        CDeposit# = 0
      END IF
      
      Num2Print = UBCustRec(1).BILLCOPY
      IF Num2Print < 1 THEN Num2Print = 1
      FOR BillCopies = 1 TO Num2Print
        ''$INCLUDE: 'NORWODBL.BI'        'Norwood
        ''$INCLUDE: 'UBALBERT.BI'        'ALBERTA
        ''$INCLUDE: 'UBALEXML.BI'        'ALEXANDER MILLS
        ''$INCLUDE: 'UBANDREW.BI'         'Andrews
        ''$INCLUDE: 'UBANSON.BI'         'Ansonville NC
        ''$INCLUDE: 'UBAPPLCH.BI'        'APPALACHIA VA
        ''$INCLUDE: 'UBARLNGT.BI'        'Arlington
        ''$INCLUDE: 'UBAURORA.BI'        'like
        ''$INCLUDE: 'UBAUTRYV.BI'        'Autryville NC
        ''$INCLUDE: 'UBBEAVER.BI'        'Beaver OK
        ''$INCLUDE: 'UBBEECH.BI2'        'like
        ''$INCLUDE: 'UBBETHEL.BI'        'like
        ''$INCLUDE: 'UBBLADEN.BI'        'Bladen
        ''$INCLUDE: 'UBBRANCH.BI'        'Branchville SC
        ''$INCLUDE: 'UBBROKNL.BI'        'Brookneal
        ''$INCLUDE: 'UBBROKFR.BI'        'BrookFord
        ''$INCLUDE: 'UBBRUNDG.BI'        'Brundage
        ''$INCLUDE: 'UBBRUNWK.BI'        'Brunswick
        ''$INCLUDE: 'UBBUCKSP.BI'        'Bucksport Water
        ''$INCLUDE: 'UBBURKVL.BI'        'Burkville VA
        ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
        ''$INCLUDE: 'UBCALDWL.BI'        'Caldwell
        ''$INCLUDE: 'UBCANDOR.BI'        'Candor
        ''$INCLUDE: 'UBCARBCH.BI'        'Carolina Beech
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
        ''$INCLUDE: 'UBFOUNTN.BI'        'Fountain
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
        ''$INCLUDE: 'UBHAMTON.BI'        'Hamilton
        ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
        ''$INCLUDE: 'UBHILBIL.BI'        'Hillsville VA
        ''$INCLUDE: 'UBHILDBR.BI'        'Hildabran
        ''$INCLUDE: 'UBHOLLY.BI'         'Holly springs
        ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
        ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
        ''$INCLUDE: 'UBJAMEST.BI'        'Jamestown
        ''$INCLUDE: 'UBJONSVL.BI'        'Jonesville NC
        ''$INCLUDE: 'UBJVILLE.BI'        'Johnsonville,  Morvan
        ''$INCLUDE: 'UBKELFRD.BI'        'Kelford NC
        ''$INCLUDE: 'UBKERSHA.bi'        'Kershaw
        ''$INCLUDE: 'UBLANDIS.BI'        'Landis,nc
        ''$INCLUDE: 'UBLEEPSA.BI'        'Lee PSA
        ''$INCLUDE: 'UBLINCOL.BI'        'Lincoln County GA
        ''$INCLUDE: 'UBLINDEN.BI'        'Linden,nc
        ''$INCLUDE: 'UBLONGVW.BI'        'Long View nc
        ''$INCLUDE: 'UBLOUISA.BI'        'Louisa
        ''$INCLUDE: 'UBMACCFL.BI'        'Maccelfield
        ''$INCLUDE: 'UBMAGNOL.BI'         'Magnolia NC
        ''$INCLUDE: 'UBMARSH.BI'         'Marshville, nc
        ''$INCLUDE: 'UBMANTEO.BI'         'Manteo NC
        ''$INCLUDE: 'UBMAYVIL.BI'        'Maysville
        ''$INCLUDE: 'UBMCMICK.BI'        'McCormick sc
        ''$INCLUDE: 'UBMEANVL.BI'        'Meanville SC
        ''$INCLUDE: 'UBMEBANE.BI'        'Mebane Orange Alamance
        ''$INCLUDE: 'UBMIDTWN.BI'        'Middletown VA
        ''$INCLUDE: 'UBMINCO.BI'         'Minco OK
        ''$INCLUDE: 'UBMONTRS.BI'         'Montross
        ''$INCLUDE: 'UBMORVAN.BI'        'Morven
        ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
        ''$INCLUDE: 'UBMTGILD.BI'        'Mt Gilead
        ''$INCLUDE: 'UBMTPLES.BI'        'Mt Plesent
        ''$INCLUDE: 'UBMURPHY.BI'        'Murphy
        ''$INCLUDE: 'UBNEWPRT.BI'        'Newport NC
        ''$INCLUDE: 'UBNLENOR.BI'         'North Lenoir
        ''$INCLUDE: 'UBNORLIN.BI'        'Norlina
        ''$INCLUDE: 'UBNORWST.BI'        'Northwest
        ''$INCLUDE: 'UBOAKCTY.BI'        'Oak City
        ''$INCLUDE: 'UBOAKBOR.BI'        'Oakboro
        ''$INCLUDE: 'UBOLDSDS.BI'        'Cleveland NC
        ''$INCLUDE: 'UBOLDSTD.BI'        'McBee SC    Old Standard 21 Line Bill
        ''$INCLUDE: 'UBPAGELN.BI'        'Pageland
        ''$INCLUDE: 'UBPEACHB.BI'        'Peachland
        ''$INCLUDE: 'UBPEMBRK.BI'        'Pembroke VA
        ''$INCLUDE: 'UBPENDTN.BI'        'Pendelton
        ''$INCLUDE: 'UBPINEBF.BI'        'Pinebluff
        ''$INCLUDE: 'UBPINKHL.BI'        'PinkHill
        ''$INCLUDE: 'UBPLYMTH.BI'        'Plymouth
        ''$INCLUDE: 'UBPOLKTN.BI'        'Polkton
        ''$INCLUDE: 'UBPOUND.BI'         'Pound VA
        ''$INCLUDE: 'UBPRINCE.BI'        'Princeton
        ''$INCLUDE: 'UBREMTON.BI'        'Remington
        ''$INCLUDE: 'UBRICHLN.BI'        'Richland
        ''$INCLUDE: 'UBROLES.BI'         'Rolesville
        ''$INCLUDE: 'UBROSEBL.BI'        'Roseboro
        ''$INCLUDE: 'UBRUSLCO.BI'        'Russell Co.
        ''$INCLUDE: 'UBRYAN.BI'        'Ryan OK
        ''$INCLUDE: 'UBSALEM.BI'         'Salem SC
        ''$INCLUDE: 'UBSALTVL.BI'        'Saltville
        ''$INCLUDE: 'UBSALUCO.BI'        'Saluda CO
        ''$INCLUDE: 'UBSAWMIL.BI'        'Sawmills
        ''$INCLUDE: 'UBSCOTTS.BI'        'Scottsburg
        ''$INCLUDE: 'UBSEDGE.BI'        'Sedgefield
        ''$INCLUDE: 'UBSPENNC.BI'        'Spencer NC
        ''$INCLUDE: 'UBSTDBIL.BI'        'Lilesville, nc
        ''$INCLUDE: 'UBSTEDBL.BI'        'Stedman
        ''$INCLUDE: 'UBSTEPHN.BI'        'Stephens City
        ''$INCLUDE: 'UBSTPAUL.BI'        'St. Paul VA
        ''$INCLUDE: 'UBSUMDAL.BI'        'Summer Dale
        ''$INCLUDE: 'UBSUNSET.BI'        'Sunset Beech
        ''$INCLUDE: 'UBSUMTON.BI'        'Summerton
        ''$INCLUDE: 'UBTALOGA.BI'        'Taloga
        ''$INCLUDE: 'UBTENNRD.BI'        'Tenn Ridge
        ''$INCLUDE: 'UBTENNIL.BI'        'Tennil GA
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
        ''$INCLUDE: 'UBWARSAW.BI'        '
        ''$INCLUDE: 'UBYADKIN.BI'        'Yadkinville
        ''$INCLUDE: 'UBLUGOFW.bi'        '
        '$INCLUDE: 'UBNLASER.BI'        '(new standard style) 17" laser
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

  IF FFFlag THEN
    PRINT #UBRpt, CHR$(12);
  END IF
  
  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, RePrintIdx
  
  IF NOT AbortFlag THEN
    PrintRptFile "Utility Bill Printing ", "UBFBILLR.PRN", 1, RetCode, 1
  END IF
  
ExitRePrintBill:
  
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
  LSET Form$(MsgF1, 0) = BillInfoRec(1).MsgLine1
  LSET Form$(MsgF2, 0) = BillInfoRec(1).MsgLine2
  LSET Form$(MsgF3, 0) = BillInfoRec(1).MsgLine3
  LSET Form$(MsgF4, 0) = BillInfoRec(1).MsgLine4
  
  Message$ = QPTrim$(BillInfoRec(1).MsgLine1)
  Msg2$ = QPTrim$(BillInfoRec(1).MsgLine2)
  Msg3$ = QPTrim$(BillInfoRec(1).MsgLine3)
  Msg4$ = QPTrim$(BillInfoRec(1).MsgLine4)
  
  LSET Form$(DepositFld, 0) = Deposit$
  
  BillDate$ = Num2Date$(BillInfoRec(1).BillDate)
  
  
  Action = 2
  RETURN
  
END SUB

