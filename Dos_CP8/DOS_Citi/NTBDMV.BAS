  DEFINT A-Z
DECLARE SUB DMVRedo ()
DECLARE SUB DMVTest ()
DECLARE SUB DMVLive ()
DECLARE SUB DelList (Items() AS ANY, Picked%(), NPicked%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE SUB PostTaxPayments ()
DECLARE SUB PrintPaymentJournal ()
DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
DECLARE SUB WPickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag%, Cnf AS ANY)
DECLARE SUB DisplayAPScrn (ScrnName$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayOperEntry ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB EnterPayments (taxtype%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%, taxtype%)
DECLARE SUB OperatorEntry ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB ButtonPress (WhatBtn, N, MooseButton, MCol, MRow)   ' ----- Check Mouse activity
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ValidateOperator ()
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Unique$ (Path$)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION WEnvTest ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION IsCustDeleted% (AcctNum&, taxtype%)
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE FUNCTION Exist% (FileName$)
  
  
  TYPE FLen2
    v AS STRING * 64
  END TYPE
  
  TYPE VACustPayListType
    CustAcct     AS LONG
    LastPayRec  AS LONG
    NumPayRec   AS LONG
    BillType AS STRING * 1      'R  or P
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'LCTax.BI'
  '$INCLUDE: 'cmfiles.BI'
  
  CONST False = 0, True = NOT False
  
  STACK 5000
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  DIM SHARED OperNum AS INTEGER, OperOKFlag AS INTEGER
  DIM SHARED PostDate AS STRING, OperPassword AS STRING
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Prepare DMV Transmission File"
  MChoice$(2) = "Prepare DMV Test File"
  MChoice$(3) = "Re-Process a Batch"
  MChoice$(4) = "Return to Main Menu"
  MChoice$(5) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((23 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 3, Col, MaxLen + 3, "DMV Processing Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      DMVLive
    CASE 2
      DMVTest
    CASE 3
      DMVRedo
    CASE 4
      RUN "tbtmenu"
    CASE 5
      CLS
      END
    END SELECT
  LOOP
  
TaxPaymentExit:
  
  RUN "TbtMenu"
  
  END

SUB DMVLive
  
  TBPath$ = "F:\"               ' change to "F:\" Before compiling
  
  
  REDIM TaxCustRec(1) AS TBPPCust
  REDIM TBPPCustBal(1) AS TBPPBal
  REDIM VehRec(1) AS PPVehType
  REDIM DMVLiveIF(1) AS DMVInformationType
  
  
  DIM SSN1$(200), LastName1$(200), FirstName1$(200), Addr1$(200), Addr2$(200), City$(200), State$(200), Zip$(200), VIN$(200), VehValue#(200), PPTaxPd$(200), PPTaxReimb$(200)
  ClearBack
  Early = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VADMVLIV"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Form$(1, 0) = DATE$
  'Fld(1).Protected = True
  
  'Check for Existing Information
  OPEN "TAXDMVIF.DAT" FOR RANDOM AS #15 LEN = LEN(DMVLiveIF(1))
  IF LOF(15) > 0 THEN
    GET 15, 1, DMVLiveIF(1)
    Form$(4, 0) = STR$(DMVLiveIF(1).PerRate)
    Form$(5, 0) = STR$(DMVLiveIF(1).Batch + 1)
    Form$(6, 0) = DMVLiveIF(1).JCode
    CLOSE 15
  ELSE
    CLOSE 15
  END IF
  Form$(2, 0) = DATE$
ReProcess:
  Action = 1
  ClearBack
  
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  
  OPEN "TAXDMVIF.DAT" FOR RANDOM AS #15 LEN = LEN(DMVLiveIF(1))
  DMVLiveIF(1).PerRate = Value#(Form$(4, 0), Ecode)
  DMVLiveIF(1).Batch = Value#(Form$(5, 0), Ecode)
  DMVLiveIF(1).JCode = Form$(6, 0)
  'PUT 15, 1, DMVLiveIF(1)
  CLOSE 15
  
  
  ' Date Calculations Here
  Batch$ = RTRIM$(LTRIM$(Form$(5, 0)))
  IF LEN(Batch$) < 3 THEN
    Batch$ = STRING$(3 - LEN(Batch$), "0") + Batch$
  END IF
  
  
  ProcessDate = Date2Num%(Form$(1, 0))
  Jury$ = Form$(6, 0)
  
  'Check Tax Year For Accuracy
  TaxYear$ = Form$(3, 0)
  'Do Not Let Years > 2004 to Process Now
  IF VAL(TaxYear$) > 2005 THEN
    CLOSE
    CLS : PRINT "Call Southern Software For 2006 DMV Submission Update "
    SLEEP 20
    EXIT SUB
  END IF
  
  IF VAL(TaxYear$) < 1999 OR VAL(TaxYear$) > 2005 THEN
    Form$(3, 0) = ""
    Frm(1).FldNo = 3
    ExitFlag = False
    GOTO ReProcess
  END IF
  GOSUB pptra   'Get %
  
  TaxRate! = Value#(Form$(4, 0), Ecode)
  Today = Date2Num%(Form$(2, 0))
  Today$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 4)
  Cdate = Date2Num%("12-31-2002")
  JulianDate = Today - Cdate

jjj1:
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365: GOTO jjj1
  
  
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$
  
  ' Create Report
  RptHandle = FREEFILE
  ReportFile$ = "DMVFILE.RPT"
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  GOSUB LiveReportHeading
  
  PayJourName$ = "T" + Jury$ + JulianDate$ + "." + Batch$
  Header$ = "Creating DMV Data File"
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  DF$ = "DEL " + PayJourName$
  SHELL (DF$)
  OPEN PayJourName$ FOR OUTPUT AS #10
  
  ShowProcessingScrn Header$
  GOSUB 60000
  NumOfTaxRecs = LOF(TBFile) / LEN(TaxCustRec(1))
  
  
  FOR cnt = 1 TO NumOfTaxRecs
    custrecno = cnt
    
    GET TBFile, custrecno, TaxCustRec(1)
    GET TBBalFile, custrecno, TBPPCustBal(1)
    
    
    LastPaidDate = 0
    
    'IF TaxCustRec(1).FirstTrans > 0 THEN
    '  STOP
    'END IF
    
    'IF TaxCustRec(1).DMVSubmitted = "Y" THEN
    '  STOP
    'ELSEIF TaxCustRec(1).DMVSubmitted1999 = "Y" THEN
    '  STOP
    'ELSEIF TaxCustRec(1).DMVSubmitted2000 = "Y" THEN
    '  STOP
    'ELSEIF TaxCustRec(1).DMVSubmitted2001 = "Y" THEN
    '  STOP
    'ELSEIF TaxCustRec(1).DMVSubmitted2002 = "Y" THEN
    '  STOP
    'ELSEIF TaxCustRec(1).DMVSubmitted2003 = "Y" AND TaxCustRec(1).Sup2003 <> "Y" THEN
    '  STOP
    'ELSEIF TaxCustRec(1).DMVSubmitted2004 = "Y" AND TaxCustRec(1).Sup2004 <> "Y" THEN
    '  STOP
    'END IF
    
    IF TaxCustRec(1).Deleted <> "Y" AND TaxCustRec(1).FirstTrans > 0 THEN
      
      'Now Check For Submissions This Date
      IF VAL(TaxYear$) = 1998 AND TaxCustRec(1).DMVSubmitted = "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 1999 AND TaxCustRec(1).DMVSubmitted1999 = "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 2000 AND TaxCustRec(1).DMVSubmitted2000 = "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 2001 AND TaxCustRec(1).DMVSubmitted2001 = "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 2002 AND TaxCustRec(1).DMVSubmitted2002 = "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 2003 AND TaxCustRec(1).DMVSubmitted2003 = "Y" AND TaxCustRec(1).Sup2003 <> "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 2004 AND TaxCustRec(1).DMVSubmitted2004 = "Y" AND TaxCustRec(1).Sup2004 <> "Y" THEN
        GOTO NoProcess
      END IF
      IF VAL(TaxYear$) = 2005 AND TaxCustRec(1).DMVSubmitted2005 = "Y" AND TaxCustRec(1).Sup2005 <> "Y" THEN
        GOTO NoProcess
      END IF
        'ELSE
        '  STOP
      'END IF
      
      'Check for Zero Balance in Current Balance Field
      'Set For Tax Year 2003 Must Redo Each Year Before It Will Work
      
      CurBal# = 0
      
      IF VAL(TaxYear$) = 2005 THEN
        CurBal# = TaxCustRec(1).CurYrPersPropTax + TaxCustRec(1).CurYrFarmEquipTax + TaxCustRec(1).CurYrMachToolsTax
        CurBal# = CurBal# + TaxCustRec(1).CurYrMerchCapTax + TaxCustRec(1).CurYrMobileHomeTax + TaxCustRec(1).CurYrLateFeeTax
        CurBal# = CurBal# + TaxCustRec(1).CurYrIntAmount + TaxCustRec(1).CurYrPenAmount
        CurBal# = INT((CurBal# * 100) + .5) / 100
      END IF
      
      IF VAL(TaxYear$) = 2004 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(1) + TBPPCustBal(1).FarmEquipAmt(1)
        CurBal# = CurBal# + TBPPCustBal(1).MachToolsAmt(1) + TBPPCustBal(1).MerchCapAmt(1)
        CurBal# = CurBal# + TBPPCustBal(1).MobileHomeAmt(1) + TBPPCustBal(1).LateFeeAmt(1) + TBPPCustBal(1).IntAmt(1) + TBPPCustBal(1).PenAmt(1)
      END IF
      
      IF VAL(TaxYear$) = 2003 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(2) + TBPPCustBal(1).FarmEquipAmt(2)
        CurBal# = CurBal# + TBPPCustBal(1).MachToolsAmt(2) + TBPPCustBal(1).MerchCapAmt(2)
        CurBal# = CurBal# + TBPPCustBal(1).MobileHomeAmt(2) + TBPPCustBal(1).LateFeeAmt(2) + TBPPCustBal(1).IntAmt(2) + TBPPCustBal(1).PenAmt(2)
      END IF
      
      IF VAL(TaxYear$) = 2002 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(3) + TBPPCustBal(1).FarmEquipAmt(3)
        CurBal# = CurBal# + TBPPCustBal(1).MachToolsAmt(3) + TBPPCustBal(1).MerchCapAmt(3)
        CurBal# = CurBal# + TBPPCustBal(1).MobileHomeAmt(3) + TBPPCustBal(1).LateFeeAmt(3) + TBPPCustBal(1).IntAmt(3) + TBPPCustBal(1).PenAmt(3)
      END IF
      
      IF VAL(TaxYear$) = 2001 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(4) + TBPPCustBal(1).FarmEquipAmt(4)
        CurBal# = CurBal# + TBPPCustBal(1).MachToolsAmt(4) + TBPPCustBal(1).MerchCapAmt(4)
        CurBal# = CurBal# + TBPPCustBal(1).MobileHomeAmt(4) + TBPPCustBal(1).LateFeeAmt(4) + TBPPCustBal(1).IntAmt(4) + TBPPCustBal(1).PenAmt(4)
      END IF
      
      IF VAL(TaxYear$) = 2000 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(5) + TBPPCustBal(1).FarmEquipAmt(5)
        CurBal# = CurBal# + TBPPCustBal(1).MachToolsAmt(5) + TBPPCustBal(1).MerchCapAmt(5)
        CurBal# = CurBal# + TBPPCustBal(1).MobileHomeAmt(5) + TBPPCustBal(1).LateFeeAmt(5) + TBPPCustBal(1).IntAmt(5) + TBPPCustBal(1).PenAmt(5)
      END IF
      
      IF VAL(TaxYear$) = 1999 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(6) + TBPPCustBal(1).FarmEquipAmt(6)
        CurBal# = CurBal# + TBPPCustBal(1).MachToolsAmt(6) + TBPPCustBal(1).MerchCapAmt(6)
        CurBal# = CurBal# + TBPPCustBal(1).MobileHomeAmt(6) + TBPPCustBal(1).LateFeeAmt(6) + TBPPCustBal(1).IntAmt(6) + TBPPCustBal(1).PenAmt(6)
      END IF
      
      
      IF CurBal# <= 0 THEN
        
        'Set Name
        LastName$ = ""
        FirstName$ = ""
        Nme$ = QPTrim$(TaxCustRec(1).CustName)
        kk = INSTR(Nme$, " ")
        IF kk > 0 THEN
          LastName$ = LEFT$(Nme$, kk - 1)
          FirstName$ = MID$(Nme$, kk + 1, (LEN(Nme$) - kk))
        END IF
        
        IF LEN(RTRIM$(LastName$)) = 0 THEN GOTO NoProcess
        
        'Personal Property Here
        PersValue# = 0
        IF TaxCustRec(1).FirstVeh > 0 THEN
          
          PropertyRecord! = TaxCustRec(1).FirstVeh
          WHILE PropertyRecord! <> 0
            
            GET TBVFile, PropertyRecord!, VehRec(1)
            
            IF LEFT$(VehRec(1).VehQ, 1) = "Y" AND VehRec(1).VehValue > 0 AND VehRec(1).VehTaxYr = VAL(TaxYear$) THEN
              
              '2004 - added to process sup bills  When bills were posted
              'if client had a zero balance then account veh were flagged  with a y in supp field
              'if client had a balance then we want to process all vehicles when balance
              'goes to zero therefore we did not flag them at all since they are to prcess
              'normally
              
              IF TaxCustRec(1).Sup2005 = "Y" AND TaxCustRec(1).DMVSubmitted2005 = "Y" AND VehRec(1).VehSupFlag <> "Y" THEN GOTO NoProcess
              IF TaxCustRec(1).Sup2004 = "Y" AND TaxCustRec(1).DMVSubmitted2004 = "Y" AND VehRec(1).VehSupFlag <> "Y" THEN GOTO NoProcess
              IF TaxCustRec(1).Sup2003 = "Y" AND TaxCustRec(1).DMVSubmitted2003 = "Y" AND VehRec(1).VehSupFlag <> "Y" THEN GOTO NoProcess
              
              VehRec(1).VehSupFlag = "P"        ' Set flag to Processed to disallow duplicate sup processing fixed 2-25-05 njp
              PUT TBVFile, PropertyRecord!, VehRec(1)
              GET TBVFile, PropertyRecord!, VehRec(1)
              
              v = v + 1
              y = 1
              SSN$ = LEFT$(TaxCustRec(1).SocSec, 3) + MID$(TaxCustRec(1).SocSec, 5, 2) + MID$(TaxCustRec(1).SocSec, 8, 4)
              SSN1$(y) = SSN$
              LastName1$(y) = LastName$
              
              kk = INSTR(FirstName$, " ")
              IF kk > 0 THEN
                FirstName$ = LEFT$(FirstName$, kk)
              END IF
              
              
              PersValue# = 0
              ProcessThisCustomer$ = "N"
              TransRecord& = TaxCustRec(1).FirstTrans
              FirstName1$(y) = FirstName$
              Addr1$(y) = TaxCustRec(1).Address2
              Addr2$(y) = ""
              City$(y) = TaxCustRec(1).City
              State$(y) = TaxCustRec(1).State
              Zip$(y) = TaxCustRec(1).Zip
              VIN$(y) = VehRec(1).VehVin
              VehValue#(y) = VehRec(1).VehValue
              
              IF VehValue#(y) > 20000 THEN VehValue#(y) = 20000 'Maximum of 20,000
              
              ' Calculate Tax Paid
              TaxPaid@ = (VehValue#(y) / 100) * TaxRate!
              TaxPaid@ = INT((TaxPaid@ * 100) + .5) / 100
              TaxPaid$ = LTRIM$(STR$(TaxPaid@ * 100))
              TaxPaid$ = LEFT$(TaxPaid$, LEN(TaxPaid$) - 2) + "." + RIGHT$(TaxPaid$, 2)
              
              PPTaxPd$(y) = TaxPaid$
              
              ' Calculate Reimbursement
              IF VehValue#(y) <= 1000 THEN
                Reimbursement@ = VAL(TaxPaid$)
                Reimbursement$ = TaxPaid$
              ELSE
                IF VehValue#(y) <= 20000 THEN
                  Reimbursement@ = TaxPaid@ * PERC!
                  Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
                  Reimbursement$ = LTRIM$(STR$(Reimbursement@))
                END IF
              END IF
              
              Reimbursement$ = LTRIM$(STR$(Reimbursement@ * 100))
              Reimbursement$ = LEFT$(Reimbursement$, LEN(Reimbursement$) - 2) + "." + RIGHT$(Reimbursement$, 2)
              
              PPTaxReimb$(y) = Reimbursement$
              TotalReimb# = TotalReimb# + Reimbursement@
              TotalReimb# = INT((TotalReimb# * 100) + .5) / 100
              GOSUB LiveDMVLines
              veh = v
              
              'Set Submit Flag and Batch

              'IF VAL(TaxYear$) = 1998 THEN
              '  TaxCustRec(1).DMVSubmitted = "Y"
              '  TaxCustRec(1).DMVBatch = Batch$
              'END IF

              IF VAL(TaxYear$) = 1999 THEN
                TaxCustRec(1).DMVSubmitted1999 = "Y"
                TaxCustRec(1).DMVBatch1999 = Batch$
              END IF
              IF VAL(TaxYear$) = 2000 THEN
                TaxCustRec(1).DMVSubmitted2000 = "Y"
                TaxCustRec(1).DMVBatch2000 = Batch$
              END IF
              IF VAL(TaxYear$) = 2001 THEN
                TaxCustRec(1).DMVSubmitted2001 = "Y"
                TaxCustRec(1).DMVBatch2001 = Batch$
              END IF
              IF VAL(TaxYear$) = 2002 THEN
                TaxCustRec(1).DMVSubmitted2002 = "Y"
                TaxCustRec(1).DMVBatch2002 = Batch$
              END IF
              
              IF VAL(TaxYear$) = 2003 THEN
                TaxCustRec(1).DMVSubmitted2003 = "Y"
                TaxCustRec(1).DMVBatch2003 = Batch$
              END IF
              
              IF VAL(TaxYear$) = 2004 THEN
                TaxCustRec(1).DMVSubmitted2004 = "Y"
                TaxCustRec(1).DMVBatch2004 = Batch$
              END IF

              IF VAL(TaxYear$) = 2005 THEN
                TaxCustRec(1).DMVSubmitted2005 = "Y"
                TaxCustRec(1).DMVBatch2005 = Batch$
              END IF

              
              PUT TBFile, custrecno, TaxCustRec(1)
            END IF
            
            PropertyRecord! = VehRec(1).VehNext
            
          WEND  'finish loop
          
          
        END IF
      END IF
      
    END IF      'Check for Customer Deleteded
    
    
NoProcess:
    ShowPctComp cnt, NumOfTaxRecs
  NEXT
  
LiveEndDMVProcess:
  'Write Header
  GOSUB LiveDMVHeaderWrite
  'Write Records
  'GOSUB LiveDMVLines
  'Show FileName
  
  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "File Name to Send is "; PayJourName$
  PRINT #RptHandle, "Total Vehicles: "; USING "######,#"; VAL(Vehicles$)
  PRINT #RptHandle, "Total of Reimbursement: "; USING "#####,#.##"; VAL(TotalAmt$)
  IF Early = 1 THEN PRINT #RptHandle, "More Vehicles to Process, Please Create an Additional Batch for This Date"
  CLOSE
  CLS
  
  EntryPoint = 1
  LPTPORT% = 1
  
  PrintRptFile Header$, ReportFile$, LPTPORT%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
LiveDMVExitJournal:
  
  EXIT SUB
  
LiveDMVLines:
  Records = 1   'One at a time
  PRINT #10, "D@";
  PRINT #10, RTRIM$(LTRIM$(STR$(Records))) + "@";
  PRINT #10, RTRIM$(SSN1$(Records)) + "@";
  PRINT #10, RTRIM$(LastName1$(Records)) + "@";
  PRINT #10, RTRIM$(FirstName1$(Records)) + "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, RTRIM$(Addr1$(Records)) + "@";
  PRINT #10, RTRIM$(Addr2$(Records)) + "@";
  PRINT #10, RTRIM$(City$(Records)) + "@";
  PRINT #10, RTRIM$(State$(Records)) + "@";
  PRINT #10, RTRIM$(Zip$(Records)) + "@";
  PRINT #10, RTRIM$(VIN$(Records)) + "@";
  VehValue$ = LTRIM$(STR$(VehValue#(Records)))
  PRINT #10, RTRIM$(VehValue$) + "@";
  PRINT #10, RTRIM$(PPTaxPd$(Records)) + "@";
  PRINT #10, RTRIM$(PPTaxReimb$(Records)) + "@";
  PRINT #10, WhatYear$ + "01" + "@";
  PRINT #10, WhatYear$ + "12" + "@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RTRIM$(Today$) + "@"
  'Added Credit Indicator Here Leave Blank For Now
  
  
  Nme$ = RTRIM$(FirstName1$(Records)) + " " + RTRIM$(LastName1$(Records))
  PRINT #RptHandle, LEFT$(Nme$, 38);
  PRINT #RptHandle, TAB(40); LEFT$(VIN$(Records), 20);
  PRINT #RptHandle, TAB(61); USING "####,#.##"; VAL(PPTaxPd$(Records));
  PRINT #RptHandle, TAB(71); USING "####.##"; VAL(PPTaxReimb$(Records))
  lc = lc + 1
  IF lc >= 56 THEN
    PRINT #RptHandle, CHR$(12);
    GOSUB LiveReportHeading
  END IF
  
  
  RETURN
  
LiveDMVHeaderWrite:
  CLOSE 10
  OPEN "TBDMV.TMP" FOR OUTPUT AS #10
  TotalAmt$ = LTRIM$(STR$(TotalReimb# * 100))
  TLen = LEN(TotalAmt$)
  IF VAL(TotalAmt$) = 0 THEN
    TotalAmt$ = "000"
  ELSE
    TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
  END IF
  TotalCrCnt$ = ""
  TotalCrAmt$ = ""
  Vehicles$ = LTRIM$(STR$(veh))
  PRINT #10, "H@";
  PRINT #10, "1@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RIGHT$(DATE$, 4) + "@";
  PRINT #10, RTRIM$(Today$) + "@";
  PRINT #10, RTRIM$(Vehicles$) + "@";
  PRINT #10, TotalAmt$ + "@";
  PRINT #10, TotalCrCnt$ + "@";
  PRINT #10, TotalCrAmt$
  CLOSE #10
  OPEN "TBDMV.TMP" FOR APPEND AS #10
  OPEN PayJourName$ FOR INPUT AS #11
  WHILE NOT EOF(11)
    LINE INPUT #11, a$
    PRINT #10, a$
  WEND
  CLOSE 10
  CLOSE 11
  DF$ = "DEL " + PayJourName$
  SHELL (DF$)
  DF$ = "RENAME TBDMV.TMP " + PayJourName$
  SHELL (DF$)
  RETURN
  
  
  
LiveReportHeading:
  PRINT #RptHandle, "DMV Processing : Data File Contents"
  PRINT #RptHandle, "Submission Date: "; Form$(2, 0); TAB(60); "Batch #"; Batch$
  PRINT #RptHandle,
  PRINT #RptHandle, "Name"; TAB(40); "VIN #"; TAB(58); "Tax Amount"; TAB(70); "PPTRA Amt"
  PRINT #RptHandle, STRING$(79, "-")
  lc = 5
  RETURN
  
60000  CLOSE TBFile
  TBFile = FREEFILE
  OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TaxCustRec(1))
  CLOSE TBVFile
  TBVFile = FREEFILE
  OPEN TBPath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(VehRec(1))
  CLOSE TBBalFile
  TBBalFile = FREEFILE
  OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
  RETURN
  
pptra:
  WhatYear = VAL(TaxYear$)
  IF WhatYear = 1998 THEN PERC! = .125: WhatYear$ = "1998"
  IF WhatYear = 1999 THEN PERC! = .275: WhatYear$ = "1999"
  IF WhatYear = 2000 THEN PERC! = .475: WhatYear$ = "2000"
  IF WhatYear = 2001 THEN PERC! = .7: WhatYear$ = "2001"
  IF WhatYear = 2002 THEN PERC! = .7: WhatYear$ = "2002"
  IF WhatYear = 2003 THEN PERC! = .7: WhatYear$ = "2003"
  IF WhatYear = 2004 THEN PERC! = .7: WhatYear$ = "2004"
  IF WhatYear = 2005 THEN PERC! = .7: WhatYear$ = "2005"
  IF WhatYear = 2006 THEN PERC! = .7: WhatYear$ = "2006"
  IF WhatYear = 2007 THEN PERC! = .7: WhatYear$ = "2007"
  IF WhatYear = 2008 THEN PERC! = .7: WhatYear$ = "2008"
  IF WhatYear = 2009 THEN PERC! = .7: WhatYear$ = "2009"
  RETURN
  
END SUB

SUB DMVRedo
  
  TBPath$ = "F:\"               ' change to "F:\" Before compiling
  REDIM TaxCustRec(1) AS TBPPCust
  REDIM TBPPCustBal(1) AS TBPPBal
  REDIM VehRec(1) AS PPVehType
  REDIM DMVLiveIF(1) AS DMVInformationType
  
  
  DIM SSN1$(200), LastName1$(200), FirstName1$(200), Addr1$(200), Addr2$(200), City$(200), State$(200), Zip$(200), VIN$(200), VehValue#(200), PPTaxPd$(200), PPTaxReimb$(200)
  ClearBack
  Early = 0
  
  LibName$ = "TAX"
  ScrnName$ = "VADMVLIV"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Form$(1, 0) = DATE$
  'Fld(1).Protected = True
  
  'Check for Existing Information
  OPEN "TAXDMVIF.DAT" FOR RANDOM AS #15 LEN = LEN(DMVLiveIF(1))
  IF LOF(15) > 0 THEN
    GET 15, 1, DMVLiveIF(1)
    Form$(4, 0) = STR$(DMVLiveIF(1).PerRate)
    Form$(5, 0) = STR$(DMVLiveIF(1).Batch + 1)
    Form$(6, 0) = DMVLiveIF(1).JCode
    CLOSE 15
  ELSE
    CLOSE 15
  END IF
  Form$(2, 0) = DATE$
ReProcess1:
  Action = 1
  ClearBack
  
  ShowCursor
  
  
  DisplayTaxScrn ScrnName$
  LOCATE 4, 28: COLOR 12
  PRINT "REDO"
  
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB
  
  
  
  ' Date Calculations Here
  Batch$ = RTRIM$(LTRIM$(Form$(5, 0)))
  IF LEN(Batch$) < 3 THEN
    Batch$ = STRING$(3 - LEN(Batch$), "0") + Batch$
  END IF
  
  
  ProcessDate = Date2Num%(Form$(1, 0))
  Jury$ = Form$(6, 0)
  
  'Check Tax Year For Accuracy
  TaxYear$ = Form$(3, 0)
  IF VAL(TaxYear$) < 1998 OR VAL(TaxYear$) > 2005 THEN
    Form$(3, 0) = ""
    Frm(1).FldNo = 3
    ExitFlag = False
    GOTO ReProcess1
  END IF
  GOSUB pptra1  'Get %
  
  TaxRate! = Value#(Form$(4, 0), Ecode)
  Today = Date2Num%(Form$(2, 0))
  Today$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 4)
  Cdate = Date2Num%("12-31-2002")
  JulianDate = Today - Cdate
jj1:
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365: GOTO jj1
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$
  
  ' Create Report
  RptHandle = FREEFILE
  ReportFile$ = "DMVFILE.RPT"
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  GOSUB LiveReportHeading1
  
  PayJourName$ = "T" + Jury$ + JulianDate$ + "." + Batch$
  Header$ = "Creating DMV Data File"
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  'DF$ = "DEL " + PayJourName$
  'SHELL (DF$)
  OPEN PayJourName$ FOR OUTPUT AS #10
  
  ShowProcessingScrn Header$
  GOSUB 60000.1
  NumOfTaxRecs = LOF(TBFile) / LEN(TaxCustRec(1))
  
  
  FOR cnt = 1 TO NumOfTaxRecs
    custrecno = cnt
    
    GET TBFile, custrecno, TaxCustRec(1)
    GET TBBalFile, custrecno, TBPPCustBal(1)
    
    LastPaidDate = 0
    
    IF TaxCustRec(1).Deleted <> "Y" AND TaxCustRec(1).FirstTrans > 0 THEN
      
      'Now Check For Submissions This Date
      IF VAL(TaxYear$) = 1998 AND TaxCustRec(1).DMVSubmitted = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch) = RTRIM$(Batch$) THEN
        GOTO Process
      END IF
      IF VAL(TaxYear$) = 1999 AND TaxCustRec(1).DMVSubmitted1999 = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch1999) = Batch$ THEN
        GOTO Process
      END IF
      IF VAL(TaxYear$) = 2000 AND TaxCustRec(1).DMVSubmitted2000 = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch2000) = Batch$ THEN
        GOTO Process
      END IF
      IF VAL(TaxYear$) = 2001 AND TaxCustRec(1).DMVSubmitted2001 = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch2001) = Batch$ THEN
        GOTO Process
      END IF
      IF VAL(TaxYear$) = 2002 AND TaxCustRec(1).DMVSubmitted2002 = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch2002) = Batch$ THEN
        GOTO Process
      END IF
      IF VAL(TaxYear$) = 2003 AND TaxCustRec(1).DMVSubmitted2003 = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch2003) = Batch$ THEN
        GOTO Process
      END IF
      IF VAL(TaxYear$) = 2004 AND TaxCustRec(1).DMVSubmitted2004 = "Y" AND RTRIM$(TaxCustRec(1).DMVBatch2004) = Batch$ THEN
        GOTO Process
      END IF
      
      
      GOTO NoProcess1
      
Process:
      'Check for Zero Balance in Current Balance Field
      'Set For Tax Year 1999 Must Redo Each Year Before It Will Work
      
      IF VAL(TaxYear$) = 2004 THEN
        CurBal# = TaxCustRec(1).CurYrPersPropTax + TaxCustRec(1).CurYrFarmEquipTax + TaxCustRec(1).CurYrMachToolsTax + TaxCustRec(1).CurYrMerchCapTax + TaxCustRec(1).CurYrMobileHomeTax + TaxCustRec(1).CurYrLateFeeTax + TaxCustRec(1).CurYrIntAmount + _
 TaxCustRec(1).CurYrPenAmount
        CurBal# = INT((CurBal# * 100) + .5) / 100
      END IF
      
      IF VAL(TaxYear$) = 2003 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(1) + TBPPCustBal(1).FarmEquipAmt(1) + TBPPCustBal(1).MachToolsAmt(1) + TBPPCustBal(1).MerchCapAmt(1) + TBPPCustBal(1).MobileHomeAmt(1) + TBPPCustBal(1).LateFeeAmt(1) + TBPPCustBal(1).IntAmt(1) +  _
TBPPCustBal(1).PenAmt(1)
      END IF
      
      IF VAL(TaxYear$) = 2002 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(2) + TBPPCustBal(1).FarmEquipAmt(2) + TBPPCustBal(1).MachToolsAmt(2) + TBPPCustBal(1).MerchCapAmt(2) + TBPPCustBal(1).MobileHomeAmt(2) + TBPPCustBal(1).LateFeeAmt(2) + TBPPCustBal(1).IntAmt(2) +  _
TBPPCustBal(1).PenAmt(2)
      END IF
      
      IF VAL(TaxYear$) = 2001 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(3) + TBPPCustBal(1).FarmEquipAmt(3) + TBPPCustBal(1).MachToolsAmt(3) + TBPPCustBal(1).MerchCapAmt(3) + TBPPCustBal(1).MobileHomeAmt(3) + TBPPCustBal(1).LateFeeAmt(3) + TBPPCustBal(1).IntAmt(3) +  _
TBPPCustBal(1).PenAmt(3)
      END IF
      
      IF VAL(TaxYear$) = 2000 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(4) + TBPPCustBal(1).FarmEquipAmt(4) + TBPPCustBal(1).MachToolsAmt(4) + TBPPCustBal(1).MerchCapAmt(4) + TBPPCustBal(1).MobileHomeAmt(4) + TBPPCustBal(1).LateFeeAmt(4) + TBPPCustBal(1).IntAmt(4) +  _
TBPPCustBal(1).PenAmt(4)
      END IF
      
      IF VAL(TaxYear$) = 1999 THEN
        CurBal# = TBPPCustBal(1).PerPropTaxAmt(5) + TBPPCustBal(1).FarmEquipAmt(5) + TBPPCustBal(1).MachToolsAmt(5) + TBPPCustBal(1).MerchCapAmt(5) + TBPPCustBal(1).MobileHomeAmt(5) + TBPPCustBal(1).LateFeeAmt(5) + TBPPCustBal(1).IntAmt(5) +  _
TBPPCustBal(1).PenAmt(5)
      END IF
      
      
      IF CurBal# <= 0 THEN
        
        'Set Name
        LastName$ = ""
        FirstName$ = ""
        Nme$ = QPTrim$(TaxCustRec(1).CustName)
        kk = INSTR(Nme$, " ")
        IF kk > 0 THEN
          LastName$ = LEFT$(Nme$, kk - 1)
          FirstName$ = MID$(Nme$, kk + 1, (LEN(Nme$) - kk))
        END IF
        
        IF LEN(RTRIM$(LastName$)) = 0 THEN GOTO NoProcess1
        
        'Personal Property Here
        PersValue# = 0
        IF TaxCustRec(1).FirstVeh > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstVeh
          WHILE PropertyRecord! <> 0
            GET TBVFile, PropertyRecord!, VehRec(1)
            
            
            IF LEFT$(VehRec(1).VehQ, 1) = "Y" AND VehRec(1).VehValue > 0 AND VehRec(1).VehTaxYr = VAL(TaxYear$) THEN
              
              
              '2004 - added to process sup bills  When bills were posted
              'if client had a zero balance then account veh were flagged  with a y in supp field
              'if client had a balance then we want to process all vehicles when balance
              'goes to zero therefore we did not flag them at all since they are to prcess
              'normally
              
              IF TaxCustRec(1).Sup2003 = "Y" AND TaxCustRec(1).DMVSubmitted2003 = "Y" AND VehRec(1).VehSupFlag <> "P" THEN GOTO NoProcess1
              
              
              v = v + 1
              y = 1
              SSN$ = LEFT$(TaxCustRec(1).SocSec, 3) + MID$(TaxCustRec(1).SocSec, 5, 2) + MID$(TaxCustRec(1).SocSec, 8, 4)
              SSN1$(y) = SSN$
              LastName1$(y) = LastName$
              
              kk = INSTR(FirstName$, " ")
              IF kk > 0 THEN
                FirstName$ = LEFT$(FirstName$, kk)
              END IF
              
              
              PersValue# = 0
              ProcessThisCustomer$ = "N"
              TransRecord& = TaxCustRec(1).FirstTrans
              FirstName1$(y) = FirstName$
              Addr1$(y) = TaxCustRec(1).Address2
              Addr2$(y) = ""
              City$(y) = TaxCustRec(1).City
              State$(y) = TaxCustRec(1).State
              Zip$(y) = TaxCustRec(1).Zip
              VIN$(y) = VehRec(1).VehVin
              VehValue#(y) = VehRec(1).VehValue
              
              IF VehValue#(y) > 20000 THEN VehValue#(y) = 20000 'Maximum of 20,000
              
              ' Calculate Tax Paid
              TaxPaid@ = (VehValue#(y) / 100) * TaxRate!
              TaxPaid@ = INT((TaxPaid@ * 100) + .5) / 100
              TaxPaid$ = LTRIM$(STR$(TaxPaid@ * 100))
              TaxPaid$ = LEFT$(TaxPaid$, LEN(TaxPaid$) - 2) + "." + RIGHT$(TaxPaid$, 2)
              
              PPTaxPd$(y) = TaxPaid$
              
              ' Calculate Reimbursement
              IF VehValue#(y) <= 1000 THEN
                Reimbursement@ = VAL(TaxPaid$)
                Reimbursement$ = TaxPaid$
              ELSE
                IF VehValue#(y) <= 20000 THEN
                  Reimbursement@ = TaxPaid@ * PERC!
                  Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
                  Reimbursement$ = LTRIM$(STR$(Reimbursement@))
                END IF
              END IF
              
              Reimbursement$ = LTRIM$(STR$(Reimbursement@ * 100))
              Reimbursement$ = LEFT$(Reimbursement$, LEN(Reimbursement$) - 2) + "." + RIGHT$(Reimbursement$, 2)
              
              PPTaxReimb$(y) = Reimbursement$
              TotalReimb# = TotalReimb# + Reimbursement@
              TotalReimb# = INT((TotalReimb# * 100) + .5) / 100
              GOSUB LiveDMVLines1
              veh = v
              
              'Set Submit Flag and Batch
              IF VAL(TaxYear$) = 1998 THEN
                TaxCustRec(1).DMVSubmitted = "Y"
                TaxCustRec(1).DMVBatch = Batch$
              END IF
              IF VAL(TaxYear$) = 1999 THEN
                TaxCustRec(1).DMVSubmitted1999 = "Y"
                TaxCustRec(1).DMVBatch1999 = Batch$
              END IF
              IF VAL(TaxYear$) = 2000 THEN
                TaxCustRec(1).DMVSubmitted2000 = "Y"
                TaxCustRec(1).DMVBatch2000 = Batch$
              END IF
              IF VAL(TaxYear$) = 2001 THEN
                TaxCustRec(1).DMVSubmitted2001 = "Y"
                TaxCustRec(1).DMVBatch2001 = Batch$
              END IF
              IF VAL(TaxYear$) = 2002 THEN
                TaxCustRec(1).DMVSubmitted2002 = "Y"
                TaxCustRec(1).DMVBatch2002 = Batch$
              END IF
              IF VAL(TaxYear$) = 2003 THEN
                TaxCustRec(1).DMVSubmitted2003 = "Y"
                TaxCustRec(1).DMVBatch2003 = Batch$
              END IF
              
              PUT TBFile, custrecno, TaxCustRec(1)
              
            END IF
            PropertyRecord! = VehRec(1).VehNext
            
          WEND  'finish loop
          
          
        END IF
      END IF
      
    END IF      'Check for Customer Deleteded
    
    
NoProcess1:
    ShowPctComp cnt, NumOfTaxRecs
    
  NEXT
  
LiveEndDMVProcess1:
  'Write Header
  GOSUB LiveDMVHeaderWrite1
  'Write Records
  'GOSUB LiveDMVLines1
  'Show FileName
  
  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "File Name to Send is "; PayJourName$
  PRINT #RptHandle, "Total Vehicles: "; USING "######,#"; VAL(Vehicles$)
  PRINT #RptHandle, "Total of Reimbursement: "; USING "#####,#.##"; VAL(TotalAmt$)
  IF Early = 1 THEN PRINT #RptHandle, "More Vehicles to Process, Please Create an Additional Batch for This Date"
  CLOSE
  CLS
  
  EntryPoint = 1
  LPTPORT% = 1
  
  PrintRptFile Header$, ReportFile$, LPTPORT%, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
LiveDMVExitJournal1:
  
  EXIT SUB
  
LiveDMVLines1:
  Records = 1   'One at a time
  PRINT #10, "D@";
  PRINT #10, RTRIM$(LTRIM$(STR$(Records))) + "@";
  PRINT #10, RTRIM$(SSN1$(Records)) + "@";
  PRINT #10, RTRIM$(LastName1$(Records)) + "@";
  PRINT #10, RTRIM$(FirstName1$(Records)) + "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, "@";
  PRINT #10, RTRIM$(Addr1$(Records)) + "@";
  PRINT #10, RTRIM$(Addr2$(Records)) + "@";
  PRINT #10, RTRIM$(City$(Records)) + "@";
  PRINT #10, RTRIM$(State$(Records)) + "@";
  PRINT #10, RTRIM$(Zip$(Records)) + "@";
  PRINT #10, RTRIM$(VIN$(Records)) + "@";
  VehValue$ = LTRIM$(STR$(VehValue#(Records)))
  PRINT #10, RTRIM$(VehValue$) + "@";
  PRINT #10, RTRIM$(PPTaxPd$(Records)) + "@";
  PRINT #10, RTRIM$(PPTaxReimb$(Records)) + "@";
  PRINT #10, WhatYear$ + "01" + "@";
  PRINT #10, WhatYear$ + "12" + "@";
  PRINT #10, Jury$ + "@"; ;
  PRINT #10, RTRIM$(Today$) + "@"
  
  
  Nme$ = RTRIM$(FirstName1$(Records)) + " " + RTRIM$(LastName1$(Records))
  PRINT #RptHandle, LEFT$(Nme$, 38);
  PRINT #RptHandle, TAB(40); LEFT$(VIN$(Records), 20);
  PRINT #RptHandle, TAB(61); USING "####,#.##"; VAL(PPTaxPd$(Records));
  PRINT #RptHandle, TAB(71); USING "####.##"; VAL(PPTaxReimb$(Records))
  lc = lc + 1
  IF lc >= 56 THEN
    PRINT #RptHandle, CHR$(12);
    GOSUB LiveReportHeading1
  END IF
  RETURN
  
LiveDMVHeaderWrite1:
  CLOSE 10
  OPEN "TBDMV.TMP" FOR OUTPUT AS #10
  TotalAmt$ = LTRIM$(STR$(TotalReimb# * 100))
  TLen = LEN(TotalAmt$)
  IF VAL(TotalAmt$) = 0 THEN
    TotalAmt$ = "0.00"
  ELSE
    TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
  END IF
  TotalCrCnt$ = ""
  TotalCrAmt$ = ""
  
  Vehicles$ = LTRIM$(STR$(veh))
  PRINT #10, "H@";
  PRINT #10, "1@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RIGHT$(DATE$, 4) + "@";
  PRINT #10, RTRIM$(Today$) + "@";
  PRINT #10, RTRIM$(Vehicles$) + "@";
  PRINT #10, TotalAmt$ + "@";
  PRINT #10, TotalCrCnt$ + "@";
  PRINT #10, TotalCrAmt$
  CLOSE #10
  OPEN "TBDMV.TMP" FOR APPEND AS #10
  OPEN PayJourName$ FOR INPUT AS #11
  WHILE NOT EOF(11)
    LINE INPUT #11, a$
    PRINT #10, a$
  WEND
  CLOSE 10
  CLOSE 11
  DF$ = "DEL " + PayJourName$
  SHELL (DF$)
  DF$ = "RENAME TBDMV.TMP " + PayJourName$
  SHELL (DF$)
  RETURN
  
LiveReportHeading1:
  PRINT #RptHandle, "DMV Processing : Data File Contents"
  PRINT #RptHandle, "Submission Date: "; Form$(2, 0); TAB(60); "Batch #"; Batch$
  PRINT #RptHandle,
  PRINT #RptHandle, "Name"; TAB(40); "VIN #"; TAB(58); "Tax Amount"; TAB(70); "PPTRA Amt"
  PRINT #RptHandle, STRING$(79, "-")
  lc = 5
  RETURN
  
60000.1  CLOSE TBFile
  TBFile = FREEFILE
  OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TaxCustRec(1))
  CLOSE TBVFile
  TBVFile = FREEFILE
  OPEN TBPath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(VehRec(1))
  CLOSE TBBalFile
  TBBalFile = FREEFILE
  OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
  RETURN
pptra1:
  WhatYear = VAL(TaxYear$)
  IF WhatYear = 1998 THEN PERC! = .125: WhatYear$ = "1998"
  IF WhatYear = 1999 THEN PERC! = .275: WhatYear$ = "1999"
  IF WhatYear = 2000 THEN PERC! = .475: WhatYear$ = "2000"
  IF WhatYear = 2001 THEN PERC! = .7: WhatYear$ = "2001"
  IF WhatYear = 2002 THEN PERC! = .7: WhatYear$ = "2002"
  IF WhatYear = 2003 THEN PERC! = .7: WhatYear$ = "2003"
  IF WhatYear = 2004 THEN PERC! = .7: WhatYear$ = "2004"
  IF WhatYear = 2005 THEN PERC! = .7: WhatYear$ = "2005"
  IF WhatYear = 2006 THEN PERC! = .7: WhatYear$ = "2006"
  IF WhatYear = 2007 THEN PERC! = .7: WhatYear$ = "2007"
  IF WhatYear = 2008 THEN PERC! = .7: WhatYear$ = "2008"
  IF WhatYear = 2009 THEN PERC! = .7: WhatYear$ = "2009"
  RETURN
  
END SUB

SUB DMVTest
  REDIM TaxCustRec(1) AS TBPPCust
  REDIM VehRec(1) AS PPVehType
  
  DIM SSN1$(50), LastName1$(50), FirstName1$(50), Addr1$(50), Addr2$(50), City$(50), State$(50), Zip$(50), VIN$(50), VehValue#(50), PPTaxPd$(50), PPTaxReimb$(50), TaxPaid$(50), Reimbursement$(50)
  
  TBPath$ = "F:\"               ' change to "F:\" Before compiling
  '   TBPath$ = "C:\qb45\lc\"
  
  ' Date Calculations Here
  TaxRate! = 3.5                'Hard Coded for Lunenburg
  Today = Date2Num%(DATE$)
  Today$ = RIGHT$(DATE$, 4) + MID$(DATE$, 4, 2) + LEFT$(DATE$, 2)
  Cdate = Date2Num%("12-31-1997")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$
  PayJourName$ = TBPath$ + "TLUR" + JulianDate$ + ".001"
  
  Header$ = "Creating DMV Test Data"
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  DF$ = "DEL " + PayJourName$
  SHELL (DF$)
  OPEN PayJourName$ FOR OUTPUT AS #10
  
  
  
  
  'ShowProcessingScrn Header$
  
  
  TaxFile = FREEFILE
  OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
  VehFile = FREEFILE
  OPEN TBPath$ + "TBTVeh.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #VehFile LEN = LEN(VehRec(1))
  NumVehRecs& = LOF(VehFile) / LEN(VehRec(1))
  
  
  
  FOR cnt = 1 TO NumOfTaxRecs
    custrecno = cnt
    GET TaxFile, custrecno, TaxCustRec(1)
    
    IF TaxCustRec(1).Deleted <> "Y" THEN
      'Set Name
      LastName$ = ""
      FirstName$ = ""
      Nme$ = QPTrim$(TaxCustRec(1).CustName)
      kk = INSTR(Nme$, " ")
      IF kk > 0 THEN
        LastName$ = LEFT$(Nme$, kk - 1)
        FirstName$ = MID$(Nme$, kk + 1, (LEN(Nme$) - kk))
      END IF
      
      
      
      
      
      'Personal Property Here
      PersValue# = 0
      IF TaxCustRec(1).FirstVeh > 0 THEN
        PropertyRecord! = TaxCustRec(1).FirstVeh
        WHILE PropertyRecord! <> 0
          GET VehFile, PropertyRecord!, VehRec(1)
          
          IF LEFT$(VehRec(1).VehQ, 1) = "Y" THEN
            v = v + 1
            SSN$ = LEFT$(TaxCustRec(1).SocSec, 3) + MID$(TaxCustRec(1).SocSec, 5, 2) + MID$(TaxCustRec(1).SocSec, 8, 4)
            SSN1$(y) = SSN$
            LastName1$(y) = LastName$
            
            kk = INSTR(FirstName$, " ")
            IF kk > 0 THEN
              FirstName$ = LEFT$(FirstName$, kk)
            END IF
            
            
            FirstName1$(y) = FirstName$
            
            Addr1$(y) = TaxCustRec(1).Address2
            Addr2$(y) = ""
            City$(y) = TaxCustRec(1).City
            State$(y) = TaxCustRec(1).State
            Zip$(y) = TaxCustRec(1).Zip
            VIN$(y) = VehRec(1).VehVin
            VehValue#(y) = VehRec(1).VehValue
            ' Calculate Tax Paid
            TaxPaid! = (VehValue#(y) / 100) * TaxRate!
            TaxPaid! = INT((TaxPaid! * 100) + .5) / 100
            TaxPaid$ = LTRIM$(STR$(TaxPaid! * 100))
            TaxPaid$ = LEFT$(TaxPaid$, LEN(TaxPaid$) - 2) + "." + RIGHT$(TaxPaid$, 2)
            TaxPaid$(y) = TaxPaid$
            PPTaxPd$(y) = TaxPaid$
            
            ' Calculate Reimbursement
            IF VehValue#(y) <= 1000 THEN
              Reimbursement@ = VAL(TaxPaid$)
              Reimbursement$ = TaxPaid$
            ELSE
              IF VehValue#(y) < 20000 THEN
                Reimbursement@ = TaxPaid! * .125
                Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
                Reimbursement$ = LTRIM$(STR$(Reimbursement@))
              ELSE
                Reimbursement@ = 20000 * .125
                Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
              END IF
            END IF
            
            Reimbursement$ = LTRIM$(STR$(Reimbursement@ * 100))
            Reimbursement$ = LEFT$(Reimbursement$, LEN(Reimbursement$) - 2) + "." + RIGHT$(Reimbursement$, 2)
            
            PPTaxReimb$(y) = Reimbursement$
            TotalReimb# = TotalReimb# + Reimbursement@
            TotalReimb# = INT((TotalReimb# * 100) + .5) / 100
            veh = veh + 1
          END IF
          
          
          
          
          
          PropertyRecord! = VehRec(1).VehNext
        WEND
      END IF
      
      '
      IF v > 15 THEN
        GOTO EndDMVProcess
      END IF
      
      
      
      
      
    END IF
    'ShowPctComp Cnt, NumOfTaxRecs
  NEXT
  
EndDMVProcess:
  'Write Header
  GOSUB DMVHeaderWrite1
  'Write Records
  GOSUB DMVLines1
  'Show FileName
  
  
  CLOSE
  CLS
  PRINT CHR$(7);
  PRINT "YOUR FILE HAS BEEN CREATED AND NAMED: "; PayJourName$
  INPUT "PLEASE RECORD THIS NAME AND PRESS THE <ENTER> KEY!"; AA
  
  
DMVExitJournal:
  EXIT SUB
DMVLines1:
  FOR Records = 1 TO veh
    PRINT #10, "D@";
    PRINT #10, RTRIM$(LTRIM$(STR$(Records))) + "@";
    PRINT #10, RTRIM$(SSN1$(Records)) + "@";
    PRINT #10, RTRIM$(LastName1$(Records)) + "@";
    PRINT #10, RTRIM$(FirstName1$(Records)) + "@";
    PRINT #10, "@";
    PRINT #10, "@";
    PRINT #10, "@";
    PRINT #10, "@";
    PRINT #10, "@";
    PRINT #10, RTRIM$(Addr1$(Records)) + "@";
    PRINT #10, RTRIM$(Addr2$(Records)) + "@";
    PRINT #10, RTRIM$(City$(Records)) + "@";
    PRINT #10, RTRIM$(State$(Records)) + "@";
    PRINT #10, RTRIM$(Zip$(Records)) + "@";
    PRINT #10, RTRIM$(VIN$(Records)) + "@";
    VehValue$ = LTRIM$(STR$(VehValue#(Records)))
    PRINT #10, RTRIM$(VehValue$) + "@";
    PRINT #10, RTRIM$(PPTaxPd$(Records)) + "@";
    PRINT #10, RTRIM$(PPTaxReimb$(Records)) + "@";
    PRINT #10, "199801" + "@";
    PRINT #10, "199812" + "@";
    PRINT #10, "LURG" + "@"; ;
    PRINT #10, RTRIM$(Today$)
    
    
    '  PUT RptHandle, NextRecord, DMVRecord(1)
  NEXT Records
  
  RETURN
DMVHeaderWrite1:
  TotalAmt$ = LTRIM$(STR$(TotalReimb# * 100))
  TLen = LEN(TotalAmt$)
  TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
  
  Vehicles$ = LTRIM$(STR$(veh))
  PRINT #10, "H@";
  PRINT #10, "1@";
  PRINT #10, "LURG@";
  PRINT #10, "1998@";
  PRINT #10, RTRIM$(Today$) + "@";
  PRINT #10, RTRIM$(Vehicles$) + "@";
  PRINT #10, TotalAmt$
  'PUT RptHandle, 1, DMVHeader(1)
  RETURN
  
  
END SUB

