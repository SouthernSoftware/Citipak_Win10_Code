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
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Prepare DMV Transmission File"
  MChoice$(2) = "Prepare DMV Test File"
  MChoice$(3) = "Re-Process a Batch"
  MChoice$(4) = "Return to Main Menu"
  
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
  Row = ((23 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    ClearBack
    
    TitleBox 3, Col, MaxLen + 3, "DMV Processing Menu - Supp's ", Cnf
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
    END SELECT
  LOOP

TaxPaymentExit:

  RUN "TbtsMenu"

  END

SUB DMVLive

  TBPath$ = "F:\"                    ' change to "F:\" Before compiling
'   TBPath$ = "C:\qb45\lc\"

  REDIM TaxCustRec(1) AS TBPPCust
  REDIM VehRec(1) AS PPVehType
  REDIM DMVLiveIF(1) AS DMVInformationType

  DIM SSN1$(150), LastName1$(150), FirstName1$(150), Addr1$(150), Addr2$(150), City$(150), State$(150), Zip$(150), VIN$(150), VehValue#(150), PPTaxPd$(150), PPTaxReimb$(150)
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
  Fld(1).Protected = True

  'Check for Existing Information
  OPEN "TAXDMVIF.DAT" FOR RANDOM AS #15 LEN = LEN(DMVLiveIF(1))
  IF LOF(15) > 0 THEN
   GET 15, 1, DMVLiveIF(1)
   Form$(3, 0) = STR$(DMVLiveIF(1).PerRate)
   Form$(4, 0) = STR$(DMVLiveIF(1).Batch + 1)
   Form$(5, 0) = DMVLiveIF(1).JCode
   CLOSE 15
   ELSE
   CLOSE 15
  END IF
  Form$(2, 0) = DATE$
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
    DMVLiveIF(1).PerRate = Value#(Form$(3, 0), Ecode)
    DMVLiveIF(1).Batch = Value#(Form$(4, 0), Ecode)
    DMVLiveIF(1).JCode = Form$(5, 0)
   PUT 15, 1, DMVLiveIF(1)
   CLOSE 15


' Date Calculations Here
  Batch$ = RTRIM$(LTRIM$(Form$(4, 0)))
  IF LEN(Batch$) < 3 THEN
   Batch$ = STRING$(3 - LEN(Batch$), "0") + Batch$
  END IF


  ProcessDate = Date2Num%(Form$(1, 0))
  Jury$ = Form$(5, 0)
  TaxRate! = Value#(Form$(3, 0), Ecode)
  Today = Date2Num%(Form$(2, 0))
  Today$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 4)
  Cdate = Date2Num%("12-31-1998")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365
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
  df$ = "DEL " + PayJourName$
  SHELL (df$)
  OPEN PayJourName$ FOR OUTPUT AS #10

  ShowProcessingScrn Header$
  GOSUB 60000
  NumOfTaxRecs = LOF(TBFile) / LEN(TaxCustRec(1))
  

   FOR Cnt = 1 TO NumOfTaxRecs
    CustRecNo = Cnt

    GET TBFile, CustRecNo, TaxCustRec(1)

    LastPaidDate = 0

  IF TaxCustRec(1).Deleted <> "Y" AND TaxCustRec(1).DMVSubmitted <> "Y" AND TaxCustRec(1).FirstTrans > 0 THEN

   'Check for Zero Balance in Current Balance Field
     CurBal# = TaxCustRec(1).CurYrPersPropTax + TaxCustRec(1).CurYrFarmEquipTax + TaxCustRec(1).CurYrMachToolsTax + TaxCustRec(1).CurYrMerchCapTax + TaxCustRec(1).CurYrMobileHomeTax + TaxCustRec(1).CurYrLateFeeTax + TaxCustRec(1).CurYrIntAmount +  _
TaxCustRec(1).CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
  
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

        IF LEFT$(VehRec(1).VehQ, 1) = "Y" AND VehRec(1).VehValue > 0 THEN
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
                Reimbursement@ = TaxPaid@ * .125
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
              TaxCustRec(1).DMVSubmitted = "Y"
              TaxCustRec(1).DMVBatch = Batch$
              PUT TBFile, CustRecNo, TaxCustRec(1)

            END IF
            PropertyRecord! = VehRec(1).VehNext

      WEND                                              'finish loop


      END IF
      END IF

      END IF            'Check for Customer Deleteded


NoProcess:
      ShowPctComp Cnt, NumOfTaxRecs
      IF veh > 2000 THEN
         Early = 1
         GOTO LiveEndDMVProcess
      END IF

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

  KILL ReportFile$

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
  PRINT #10, "199801" + "@";
  PRINT #10, "199812" + "@";
  PRINT #10, Jury$ + "@"; ;
  PRINT #10, RTRIM$(Today$)


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
    TotalAmt$ = "0.00"
   ELSE
    TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
   END IF
   Vehicles$ = LTRIM$(STR$(veh))
   PRINT #10, "H@";
   PRINT #10, "1@";
   PRINT #10, Jury$ + "@";
   PRINT #10, "1998@";
   PRINT #10, RTRIM$(Today$) + "@";
   PRINT #10, RTRIM$(Vehicles$) + "@";
   PRINT #10, TotalAmt$
   CLOSE #10
   OPEN "TBDMV.TMP" FOR APPEND AS #10
   OPEN PayJourName$ FOR INPUT AS #11
   WHILE NOT EOF(11)
   LINE INPUT #11, a$
   PRINT #10, a$
   WEND
   CLOSE 10
   CLOSE 11
    df$ = "DEL " + PayJourName$
   SHELL (df$)
    df$ = "RENAME TBDMV.TMP " + PayJourName$
   SHELL (df$)
   RETURN



LiveReportHeading:
  PRINT #RptHandle, "DMV Processing : Data File Contents"
  PRINT #RptHandle, "Submission Date: "; Form$(2, 0); TAB(60); "Batch #"; Batch$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Name"; TAB(40); "VIN #"; TAB(60); "Tax Paid"; TAB(70); "Reimb Amt"
  PRINT #RptHandle, STRING$(79, "-")
  lc = 5
  RETURN

60000  CLOSE TBFile
       TBFile = FREEFILE
       OPEN TBPath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TaxCustRec(1))
       CLOSE TBVFile
       TBVFile = FREEFILE
       OPEN TBPath$ + "TTSVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(VehRec(1))
       RETURN

END SUB

SUB DMVRedo

  TBPath$ = "F:\"                    ' change to "F:\" Before compiling
'   TBPath$ = "C:\qb45\lc\"

  REDIM TaxCustRec(1) AS TBPPCust
  REDIM VehRec(1) AS PPVehType
  REDIM DMVLiveIF(1) AS DMVInformationType

  DIM SSN1$(150), LastName1$(150), FirstName1$(150), Addr1$(150), Addr2$(150), City$(150), State$(150), Zip$(150), VIN$(150), VehValue#(150), PPTaxPd$(150), PPTaxReimb$(150)
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
  Fld(1).Protected = True

  'Check for Existing Information
  OPEN "TAXDMVIF.DAT" FOR RANDOM AS #15 LEN = LEN(DMVLiveIF(1))
  IF LOF(15) > 0 THEN
   GET 15, 1, DMVLiveIF(1)
   Form$(3, 0) = STR$(DMVLiveIF(1).PerRate)
   Form$(4, 0) = STR$(DMVLiveIF(1).Batch + 1)
   Form$(5, 0) = DMVLiveIF(1).JCode
   CLOSE 15
   ELSE
   CLOSE 15
  END IF
  Form$(2, 0) = DATE$
  Action = 1
  ClearBack

  ShowCursor

  DisplayTaxScrn ScrnName$
  COLOR 12
  LOCATE 6, 28: PRINT "REDO"

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
    DMVLiveIF(1).PerRate = Value#(Form$(3, 0), Ecode)
    DMVLiveIF(1).Batch = Value#(Form$(4, 0), Ecode)
    DMVLiveIF(1).JCode = Form$(5, 0)
   PUT 15, 1, DMVLiveIF(1)
   CLOSE 15


' Date Calculations Here
  Batch$ = RTRIM$(LTRIM$(Form$(4, 0)))
  IF LEN(Batch$) < 3 THEN
   Batch$ = STRING$(3 - LEN(Batch$), "0") + Batch$
  END IF


  ProcessDate = Date2Num%(Form$(1, 0))
  Jury$ = Form$(5, 0)
  TaxRate! = Value#(Form$(3, 0), Ecode)
  Today = Date2Num%(Form$(2, 0))
  Today$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 4)
  Cdate = Date2Num%("12-31-1998")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$

  Batch! = VAL(Batch$)

' Create Report
  RptHandle = FREEFILE
  ReportFile$ = "DMVFILE.RPT"
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  GOSUB RLiveReportHeading

  PayJourName$ = "T" + Jury$ + JulianDate$ + "." + Batch$
  Header$ = "Creating DMV Data File"
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  df$ = "DEL " + PayJourName$
  SHELL (df$)
  OPEN PayJourName$ FOR OUTPUT AS #10

  ShowProcessingScrn Header$
  GOSUB 60000.1
  NumOfTaxRecs = LOF(TBFile) / LEN(TaxCustRec(1))


   FOR Cnt = 1 TO NumOfTaxRecs
    CustRecNo = Cnt

    GET TBFile, CustRecNo, TaxCustRec(1)

    LastPaidDate = 0

  IF TaxCustRec(1).DMVSubmitted = "Y" AND VAL(TaxCustRec(1).DMVBatch) = Batch! THEN

   'Check for Zero Balance in Current Balance Field
     CurBal# = TaxCustRec(1).CurYrPersPropTax + TaxCustRec(1).CurYrFarmEquipTax + TaxCustRec(1).CurYrMachToolsTax + TaxCustRec(1).CurYrMerchCapTax + TaxCustRec(1).CurYrMobileHomeTax + TaxCustRec(1).CurYrLateFeeTax + TaxCustRec(1).CurYrIntAmount +  _
TaxCustRec(1).CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100

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

      IF LEN(RTRIM$(LastName$)) = 0 THEN GOTO RNoProcess

       'Personal Property Here
        PersValue# = 0
   IF TaxCustRec(1).FirstVeh > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstVeh
          WHILE PropertyRecord! <> 0
        GET TBVFile, PropertyRecord!, VehRec(1)

        IF LEFT$(VehRec(1).VehQ, 1) = "Y" AND VehRec(1).VehValue > 0 THEN
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
                Reimbursement@ = TaxPaid@ * .125
                Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
                Reimbursement$ = LTRIM$(STR$(Reimbursement@))
               END IF
              END IF

              Reimbursement$ = LTRIM$(STR$(Reimbursement@ * 100))
              Reimbursement$ = LEFT$(Reimbursement$, LEN(Reimbursement$) - 2) + "." + RIGHT$(Reimbursement$, 2)

              PPTaxReimb$(y) = Reimbursement$
              TotalReimb# = TotalReimb# + Reimbursement@
              TotalReimb# = INT((TotalReimb# * 100) + .5) / 100
              GOSUB RLiveDMVLines
              veh = v
              TaxCustRec(1).DMVSubmitted = "Y"
              TaxCustRec(1).DMVBatch = Batch$
              PUT TBFile, CustRecNo, TaxCustRec(1)

            END IF
            PropertyRecord! = VehRec(1).VehNext

      WEND                                              'finish loop


      END IF
      END IF

      END IF            'Check for Existing Batch


RNoProcess:
      ShowPctComp Cnt, NumOfTaxRecs
      IF veh > 2000 THEN
         Early = 1
         GOTO RLiveEndDMVProcess
      END IF

  NEXT

RLiveEndDMVProcess:
  'Write Header
   GOSUB RLiveDMVHeaderWrite
  'Write Records
  'GOSUB LiveDMVLines
  'Show FileName

  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "File Name to Send is "; PayJourName$
  PRINT #RptHandle, "Total Vehicles: "; USING "######,#"; VAL(Vehicles$)
  PRINT #RptHandle, "Total of Reimbursement: "; USING "#####,#.##"; VAL(TotalAmt$)
  CLOSE
   CLS

  EntryPoint = 1
  LPTPORT% = 1

  PrintRptFile Header$, ReportFile$, LPTPORT%, RetCode%, EntryPoint

  KILL ReportFile$

RLiveDMVExitJournal:

  EXIT SUB

RLiveDMVLines:
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
  PRINT #10, "199801" + "@";
  PRINT #10, "199812" + "@";
  PRINT #10, Jury$ + "@"; ;
  PRINT #10, RTRIM$(Today$)


  Nme$ = RTRIM$(FirstName1$(Records)) + " " + RTRIM$(LastName1$(Records))
  PRINT #RptHandle, LEFT$(Nme$, 38);
  PRINT #RptHandle, TAB(40); LEFT$(VIN$(Records), 20);
  PRINT #RptHandle, TAB(61); USING "####,#.##"; VAL(PPTaxPd$(Records));
  PRINT #RptHandle, TAB(71); USING "####.##"; VAL(PPTaxReimb$(Records))
  lc = lc + 1
  IF lc >= 56 THEN
   PRINT #RptHandle, CHR$(12);
   GOSUB RLiveReportHeading
  END IF


  RETURN

RLiveDMVHeaderWrite:
   CLOSE 10
   OPEN "TBDMV.TMP" FOR OUTPUT AS #10
   TotalAmt$ = LTRIM$(STR$(TotalReimb# * 100))
   TLen = LEN(TotalAmt$)
   IF VAL(TotalAmt$) = 0 THEN
    TotalAmt$ = "0.00"
   ELSE
    TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
   END IF
   Vehicles$ = LTRIM$(STR$(veh))
   PRINT #10, "H@";
   PRINT #10, "1@";
   PRINT #10, Jury$ + "@";
   PRINT #10, "1998@";
   PRINT #10, RTRIM$(Today$) + "@";
   PRINT #10, RTRIM$(Vehicles$) + "@";
   PRINT #10, TotalAmt$
   CLOSE #10
   OPEN "TBDMV.TMP" FOR APPEND AS #10
   OPEN PayJourName$ FOR INPUT AS #11
   WHILE NOT EOF(11)
   LINE INPUT #11, a$
   PRINT #10, a$
   WEND
   CLOSE 10
   CLOSE 11
    df$ = "DEL " + PayJourName$
   SHELL (df$)
    df$ = "RENAME TBDMV.TMP " + PayJourName$
   SHELL (df$)
   RETURN



RLiveReportHeading:
  PRINT #RptHandle, "DMV Processing : Data File Contents"
  PRINT #RptHandle, "Submission Date: "; Form$(2, 0); TAB(60); "Batch #"; Batch$
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Name"; TAB(40); "VIN #"; TAB(60); "Tax Paid"; TAB(70); "Reimb Amt"
  PRINT #RptHandle, STRING$(79, "-")
  lc = 5
  RETURN

60000.1  CLOSE TBFile
       TBFile = FREEFILE
       OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TaxCustRec(1))
       CLOSE TBVFile
       TBVFile = FREEFILE
       OPEN TBPath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(VehRec(1))
       RETURN

END SUB

SUB DMVTest
  REDIM TaxCustRec(1) AS TBPPCust
  REDIM VehRec(1) AS PPVehType

  DIM SSN1$(50), LastName1$(50), FirstName1$(50), Addr1$(50), Addr2$(50), City$(50), State$(50), Zip$(50), VIN$(50), VehValue#(50), PPTaxPd$(50), PPTaxReimb$(50), TaxPaid$(50), Reimbursement$(50)

  TBPath$ = "F:\"                    ' change to "F:\" Before compiling
'   TBPath$ = "C:\qb45\lc\"

' Date Calculations Here
  TaxRate! = 3.5         'Hard Coded for Lunenburg
  Today = Date2Num%(DATE$)
  Today$ = RIGHT$(DATE$, 4) + MID$(DATE$, 4, 2) + LEFT$(DATE$, 2)
  Cdate = Date2Num%("12-31-1997")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$
  PayJourName$ = TBPath$ + "TLURG" + JulianDate$ + ".001"

  Header$ = "Creating DMV Test Data"
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  df$ = "DEL " + PayJourName$
  SHELL (df$)
  OPEN PayJourName$ FOR OUTPUT AS #10




  'ShowProcessingScrn Header$


  TaxFile = FREEFILE
  OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
 
  VehFile = FREEFILE
  OPEN TBPath$ + "TBTVeh.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #VehFile LEN = LEN(VehRec(1))
  NumOfVehRecs = LOF(VehFile) / LEN(VehRec(1))



   FOR Cnt = 1 TO NumOfTaxRecs
    CustRecNo = Cnt
    GET TaxFile, CustRecNo, TaxCustRec(1)

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

