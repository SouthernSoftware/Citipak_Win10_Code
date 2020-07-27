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
DECLARE SUB EnterPayments (TaxType%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%, TaxType%)
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
DECLARE FUNCTION IsCustDeleted% (AcctNum&, TaxType%)
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
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'VApaymnt.BI'
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
  MChoice$(3) = "Re-Process DMV File"
  MChoice$(4) = "Exit to OS"
  
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
    
    TitleBox 3, Col, MaxLen + 3, "DMV Processing Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    


    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      DMVLive
    CASE 2
      'DMVTest
    CASE 3
      DMVRedo
    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

TaxPaymentExit:

  RUN "TaxMenu"

  END

SUB DMVLive
  REDIM DMVHeader(1) AS DMVHeader
  REDIM DMVRecord(1) AS DMVRecord
  REDIM TaxSetUp(1) AS TaxMasterType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM TransRec(1) AS TaxTransactionType
  REDIM PersRec(1) AS PersonalRecType
  REDIM DMVLiveIF(1) AS DMVInformationType

  DIM SSN1$(175), LastName1$(175), FirstName1$(175), Addr1$(175), Addr2$(175), City$(175), State$(175), Zip$(175), Vin$(175), VehValue#(175), PPTaxPd$(175), PPTaxReimb$(175), EMth$(175), SMth$(175)
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
   PUT 15, 1, DMVLiveIF(1)
   CLOSE 15


' Date Calculations Here
  Batch$ = RTRIM$(LTRIM$(Form$(5, 0)))
  IF LEN(Batch$) < 3 THEN
   Batch$ = STRING$(3 - LEN(Batch$), "0") + Batch$
  END IF

  
  ProcessDate = Date2Num%(Form$(1, 0))
  Jury$ = Form$(6, 0)
  TaxYear$ = Form$(3, 0)
  'Check Tax Year For Accuracy
  IF VAL(TaxYear$) < 1998 OR VAL(TaxYear$) > 2010 THEN
   Form$(3, 0) = ""
   Frm(1).FldNo = 3
   ExitFlag = False
   GOTO ReProcess
  END IF
    GOSUB pptra 'Get %


  TaxRate! = Value#(Form$(4, 0), Ecode)

  Today = Date2Num%(Form$(2, 0))
  Today$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 4)
  'Cdate = Date2Num%("12-31-2000")
  Cdate = Date2Num%("12-31-2003")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN JulianDate = JulianDate - 365
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$

' Create Report
  RptHandle = FREEFILE
  ReportFile$ = "DMVFILE.RPT"
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  GOSUB ReportHeading

  PayJourName$ = "T" + Jury$ + JulianDate$ + "." + Batch$
  Header$ = "Creating DMV Data File"
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  DF$ = "DEL " + PayJourName$
  SHELL (DF$)
  OPEN PayJourName$ FOR OUTPUT AS #10
  
  ShowProcessingScrn Header$
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))

  TransFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TransFile LEN = LEN(TransRec(1))



   FOR Cnt = 1 TO NumOfTaxRecs
    custrecno = Cnt
    
    GET TaxFile, custrecno, TaxCustRec(1)
 
    LastPaidDate = 0

    IF NOT TaxCustRec(1).Deleted THEN

      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name


     'Personal Property Here
      PersValue# = 0
      ProcessThisCustomer$ = "N"
      TransRecord& = TaxCustRec(1).LastTrans
      

      WHILE TransRecord& <> 0
        GET TransFile, TransRecord&, TransRec(1)
        IF TransRec(1).BillType <> "P" THEN GOTO NextOne
        IF TransRec(1).TranType <> 1 THEN GOTO NextOne
        IF TransRec(1).DMVSubmitted = "Y" THEN GOTO NextOne
        IF TransRec(1).TaxYear <> VAL(TaxYear$) THEN GOTO NextOne       'Add Processing to Pull by Proper Tax Year
        
        Balance# = 0
        Balance# = Round#(TransRec(1).Revenue.Principle1 + TransRec(1).Revenue.Principle2 + TransRec(1).Revenue.Principle3 + TransRec(1).Revenue.Principle4 + TransRec(1).Revenue.Principle5)
        Balance# = Round#(Balance# - (TransRec(1).Revenue.Principle1Pd + TransRec(1).Revenue.Principle2Pd + TransRec(1).Revenue.Principle3Pd + TransRec(1).Revenue.Principle4Pd + TransRec(1).Revenue.Principle5Pd))

        IF Balance# <= 0 THEN
         ProcessThisCustomer$ = "Y"
         updaterecord& = TransRecord&
         GOTO ProcessMe
        END IF

NextOne:
       IF TransRec(1).TranType = 2 THEN
         LastPaidDate = TransRec(1).TransDate
       END IF
         TransRecord& = TransRec(1).LastTrans
     WEND

ProcessMe:

  IF ProcessThisCustomer$ = "Y" AND LastPaidDate <= ProcessDate THEN


          IF TaxCustRec(1).FirstPersRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPersRec
          WHILE PropertyRecord! <> 0
            GET PersTaxFile, PropertyRecord!, PersRec(1)

             IF LEFT$(PersRec(1).Desc5, 1) = "Y" AND PersRec(1).PersVal > .01 AND MID$(PersRec(1).Desc5, 3, 2) <> "BP" AND PersRec(1).DMVSubmitted <> "Y" AND (VAL(RIGHT$(PersRec(1).Desc5, 5)) = VAL(TaxYear$)) THEN
              PersRec(1).DMVSubmitted = "Y"
              PUT PersTaxFile, PropertyRecord!, PersRec(1)

        'Update Batch Here When you have a good vehicle
         GET TransFile, updaterecord&, TransRec(1)
         TransRec(1).DMVSubmitted = "Y"
         TransRec(1).DMVBatch = VAL(Batch$)
         PUT TransFile, updaterecord&, TransRec(1)

             
              v = v + 1
              SSN1$(v) = TaxCustRec(1).CSSN
              LastName1$(v) = TaxCustRec(1).LName

              FirstName$ = TaxCustRec(1).FNAME
              kk = INSTR(FirstName$, " ")
              IF kk > 0 THEN
              FirstName$ = LEFT$(FirstName$, kk)
              END IF


              FirstName1$(v) = FirstName$
              'Normal
              Addr1$(v) = TaxCustRec(1).Addr1
              Addr2$(v) = TaxCustRec(1).Addr2
              'Victoria
              'Addr1$(v) = TaxCustRec(1).Addr2
              'Addr2$(v) = ""


              City$(v) = TaxCustRec(1).City
              State$(v) = TaxCustRec(1).State$
              Zip$(v) = TaxCustRec(1).Zip
              Vin$(v) = QPTrim$(PersRec(1).DESC1)
              VehValue#(v) = PersRec(1).PersVal
              SMth$ = MID$(PersRec(1).Desc5, 5, 2)
              EMth$ = MID$(PersRec(1).Desc5, 12, 2)

              IF VAL(SMth$) < 1 OR VAL(SMth$) > 12 THEN SMth$ = "01"
              IF VAL(EMth$) < 1 OR VAL(EMth$) > 12 THEN EMth$ = "12"
              SMth$(v) = SMth$
              EMth$(v) = EMth$
             
              IF VehValue#(v) > 20000 THEN VehValue#(v) = 20000 'Maximum of 20,000

            ' Calculate Tax Paid
              TaxPaid@ = (VehValue#(v) / 100) * TaxRate!
              TaxPaid@ = INT((TaxPaid@ * 100) + .5) / 100
              TaxPaid$ = LTRIM$(STR$(TaxPaid@ * 100))
              IF LEN(TaxPaid$) = 1 THEN
                TaxPaid$ = TaxPaid$ + ".00"
              ELSE
                TaxPaid$ = LEFT$(TaxPaid$, LEN(TaxPaid$) - 2) + "." + RIGHT$(TaxPaid$, 2)
              END IF
              PPTaxPd$(v) = TaxPaid$

            ' Calculate Reimbursement
              IF VehValue#(v) <= 1000 THEN
               Reimbursement@ = VAL(TaxPaid$)
               Reimbursement$ = TaxPaid$
               ELSE
               IF VehValue#(v) <= 20000 THEN
                Reimbursement@ = TaxPaid@ * Perc!
                Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
                Reimbursement$ = LTRIM$(STR$(Reimbursement@))
               END IF
              END IF

              Reimbursement$ = LTRIM$(STR$(Reimbursement@ * 100))
              Reimbursement$ = LEFT$(Reimbursement$, LEN(Reimbursement$) - 2) + "." + RIGHT$(Reimbursement$, 2)

              PPTaxReimb$(v) = Reimbursement$
              TotalReimb# = TotalReimb# + Reimbursement@
              TotalReimb# = INT((TotalReimb# * 100) + .5) / 100
              veh = v
         
             END IF

            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF

       END IF
      END IF
      ShowPctComp Cnt, NumOfTaxRecs
      IF veh > 170 THEN
         Early = 1
         GOTO EndDMVProcess
      END IF

  NEXT

EndDMVProcess:
  'Write Header
   GOSUB DMVHeaderWrite
  'Write Records
   GOSUB DMVLines
  'Show FileName

  PRINT #RptHandle, STRING$(79, "-")
  PRINT #RptHandle, "File Name to Send is "; PayJourName$
  PRINT #RptHandle, "Total Vehicles: "; USING "######,#"; VAL(Vehicles$)
  PRINT #RptHandle, "Total of Reimbursement: "; USING "#####,#.##"; VAL(TotalAmt$)
  IF Early = 1 THEN PRINT #RptHandle, "More Vehicles to Process, Please Create an Additional Batch for This Date"
  PRINT #RptHandle, CHR$(12);
  CLOSE
   CLS

  EntryPoint = 1
  LPTPORT% = 1

  PrintRptFile Header$, ReportFile$, LPTPORT%, RetCode%, EntryPoint

  KILL ReportFile$

DMVExitJournal:

  EXIT SUB

DMVLines:
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
  PRINT #10, RTRIM$(Vin$(Records)) + "@";
   VehValue$ = LTRIM$(STR$(VehValue#(Records)))
  PRINT #10, RTRIM$(VehValue$) + "@";
  PRINT #10, RTRIM$(PPTaxPd$(Records)) + "@";
  PRINT #10, RTRIM$(PPTaxReimb$(Records)) + "@";
  PRINT #10, LTRIM$(TaxYear$) + SMth$(Records) + "@";
  PRINT #10, LTRIM$(TaxYear$) + EMth$(Records) + "@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RTRIM$(Today$) + "@"       'ADDED NEW CODE FOR 2001
 
  Nme$ = RTRIM$(FirstName1$(Records)) + " " + RTRIM$(LastName1$(Records))
  PRINT #RptHandle, LEFT$(Nme$, 38);
  PRINT #RptHandle, TAB(40); RTRIM$(Vin$(Records));
  PRINT #RptHandle, TAB(61); USING "####,#.##"; VAL(PPTaxPd$(Records));
  PRINT #RptHandle, TAB(71); USING "####.##"; VAL(PPTaxReimb$(Records))
  lc = lc + 1
  IF lc >= 56 THEN
   PRINT #RptHandle, CHR$(12);
   GOSUB ReportHeading
  END IF

  NEXT Records
  RETURN

DMVHeaderWrite:
  TotalAmt$ = LTRIM$(STR$(TotalReimb# * 100))
  TLen = LEN(TotalAmt$)

  TotalCrCnt$ = ""
  TotalCrAmt$ = ""

  IF VAL(TotalAmt$) = 0 THEN
    TotalAmt$ = "0.00"
   ELSE
    TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
  END IF
  Vehicles$ = LTRIM$(STR$(veh))
  PRINT #10, "H@";
  PRINT #10, "1@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RIGHT$(DATE$, 4) + "@";
  'PRINT #10, WhatYear$ + "@";
  PRINT #10, RTRIM$(Today$) + "@";
  PRINT #10, RTRIM$(Vehicles$) + "@";
  PRINT #10, TotalAmt$ + "@";
  PRINT #10, TotalCrCnt$ + "@";
  PRINT #10, TotalCrAmt$
  RETURN


ReportHeading:
 PRINT #RptHandle, "DMV Processing : Data File Contents"
 PRINT #RptHandle, "Submission Date: "; Form$(2, 0); TAB(60); "Batch #"; Batch$
 PRINT #RptHandle, ""
 PRINT #RptHandle,
 PRINT #RptHandle, "Name"; TAB(40); "VIN #"; TAB(58); "Tax Amount"; TAB(70); "PPTRA Amt"
 PRINT #RptHandle, STRING$(79, "-")
 lc = 5
 RETURN

pptra:
 WhatYear = VAL(TaxYear$)
 IF WhatYear = 1998 THEN Perc! = .125: WhatYear$ = "1998"
 IF WhatYear = 1999 THEN Perc! = .275: WhatYear$ = "1999"
 IF WhatYear = 2000 THEN Perc! = .475: WhatYear$ = "2000"
 IF WhatYear = 2001 THEN Perc! = .7: WhatYear$ = "2001"
 IF WhatYear >= 2002 THEN Perc! = .7: WhatYear$ = QPTrim$(STR$(WhatYear))
 RETURN

END SUB

SUB DMVRedo
  REDIM DMVHeader(1) AS DMVHeader
  REDIM DMVRecord(1) AS DMVRecord
  REDIM TaxSetUp(1) AS TaxMasterType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM TransRec(1) AS TaxTransactionType
  REDIM PersRec(1) AS PersonalRecType
  REDIM DMVLiveIF(1) AS DMVInformationType

  DIM SSN1$(175), LastName1$(175), FirstName1$(175), Addr1$(175), Addr2$(175), City$(175), State$(175), Zip$(175), Vin$(175), VehValue#(175), PPTaxPd$(175), PPTaxReimb$(175), SMth$(175), EMth$(175)
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

RDReProcess:

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
  TaxYear$ = Form$(3, 0)
  'Check Tax Year For Accuracy
  IF VAL(TaxYear$) < 1998 OR VAL(TaxYear$) > 2002 THEN
   Form$(3, 0) = ""
   Frm(1).FldNo = 3
   ExitFlag = False
   GOTO RDReProcess
  END IF
    GOSUB RDpptra 'Get %


  TaxRate! = Value#(Form$(4, 0), Ecode)

  Today = Date2Num%(Form$(2, 0))
  Today$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 4)
  'Cdate = Date2Num%("12-31-2000")
  Cdate = Date2Num%("12-31-2003")
  JulianDate = Today - Cdate
  IF JulianDate > 365 THEN
    JulianDate = JulianDate - 365
  END IF
  JulianDate$ = LTRIM$(STR$(JulianDate))
  IF LEN(JulianDate$) < 3 THEN
    JulianDate$ = STRING$(3 - LEN(JulianDate$), "0") + JulianDate$
  END IF

' Create Report
  RptHandle = FREEFILE
  ReportFile$ = "DMVFILE.RPT"
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  GOSUB RDReportHeading

  PayJourName$ = "T" + Jury$ + JulianDate$ + "." + Batch$
  Header$ = "Creating DMV Data File"
  OPEN PayJourName$ FOR OUTPUT AS #10
  CLOSE 10
  DF$ = "DEL " + PayJourName$
  SHELL (DF$)
  OPEN PayJourName$ FOR OUTPUT AS #10


  ShowProcessingScrn Header$

  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))

  TransFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TransFile LEN = LEN(TransRec(1))



   FOR Cnt = 1 TO NumOfTaxRecs
    custrecno = Cnt

    GET TaxFile, custrecno, TaxCustRec(1)

    LastPaidDate = 0

    IF NOT TaxCustRec(1).Deleted THEN

      Nme$ = QPTrim$(TaxCustRec(1).FNAME) + " " + QPTrim$(TaxCustRec(1).LName)
      Nme$ = QPTrim$(Nme$)      'this one cleans up those with only last name


     'Personal Property Here
      PersValue# = 0
      ProcessThisCustomer$ = "N"
      TransRecord& = TaxCustRec(1).LastTrans


      WHILE TransRecord& <> 0
        GET TransFile, TransRecord&, TransRec(1)
        IF TransRec(1).BillType <> "P" THEN GOTO RDNextOne
        IF TransRec(1).TranType <> 1 THEN GOTO RDNextOne
        IF TransRec(1).TaxYear <> VAL(TaxYear$) THEN GOTO RDNextOne       'Add Processing to Pull by Proper Tax Year
        IF TransRec(1).DMVSubmitted = "Y" AND TransRec(1).DMVBatch = VAL(Batch$) THEN
        Balance# = 0
        Balance# = Round#(TransRec(1).Revenue.Principle1 + TransRec(1).Revenue.Principle2 + TransRec(1).Revenue.Principle3 + TransRec(1).Revenue.Principle4 + TransRec(1).Revenue.Principle5)
        Balance# = Round#(Balance# + TransRec(1).Revenue.Interest + TransRec(1).Revenue.Penalty + TransRec(1).Revenue.Collection)
        Balance# = Round#(Balance# - (TransRec(1).Revenue.Principle1Pd + TransRec(1).Revenue.Principle2Pd + TransRec(1).Revenue.Principle3Pd + TransRec(1).Revenue.Principle4Pd + TransRec(1).Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (TransRec(1).Revenue.InterestPd + TransRec(1).Revenue.PenaltyPd + TransRec(1).Revenue.CollectionPd))

        IF Balance# <= 0 THEN
         ProcessThisCustomer$ = "Y"
         updaterecord& = TransRecord&
        END IF
        
        END IF

RDNextOne:
       IF TransRec(1).TranType = 2 THEN
         LastPaidDate = TransRec(1).TransDate
       END IF
         TransRecord& = TransRec(1).LastTrans
     WEND

RDProcessMe:

  IF ProcessThisCustomer$ = "Y" THEN



          IF TaxCustRec(1).FirstPersRec > 0 THEN
          PropertyRecord! = TaxCustRec(1).FirstPersRec
          WHILE PropertyRecord! <> 0
            GET PersTaxFile, PropertyRecord!, PersRec(1)

             IF LEFT$(PersRec(1).Desc5, 1) = "Y" AND PersRec(1).PersVal > .01 AND MID$(PersRec(1).Desc5, 3, 2) <> "BP" AND (VAL(RIGHT$(PersRec(1).Desc5, 5)) = VAL(TaxYear$)) THEN


              v = v + 1
              SSN1$(v) = TaxCustRec(1).CSSN
              LastName1$(v) = TaxCustRec(1).LName

              FirstName$ = TaxCustRec(1).FNAME
              kk = INSTR(FirstName$, " ")
              IF kk > 0 THEN
              FirstName$ = LEFT$(FirstName$, kk)
              END IF


              FirstName1$(v) = FirstName$
              'Normal
              Addr1$(v) = TaxCustRec(1).Addr1
              Addr2$(v) = TaxCustRec(1).Addr2

              'Victoria
              'Addr1$(v) = TaxCustRec(1).Addr2
              'Addr2$(v) = ""

              City$(v) = TaxCustRec(1).City
              State$(v) = TaxCustRec(1).State$
              Zip$(v) = TaxCustRec(1).Zip
              Vin$(v) = QPTrim$(PersRec(1).DESC1)
              VehValue#(v) = PersRec(1).PersVal

            ' Figure Date Range
              SMth$ = MID$(PersRec(1).Desc5, 5, 2)
              EMth$ = MID$(PersRec(1).Desc5, 12, 2)

              IF VAL(SMth$) < 1 OR VAL(SMth$) > 12 THEN SMth$ = "01"
              IF VAL(EMth$) < 1 OR VAL(EMth$) > 12 THEN EMth$ = "12"
              SMth$(v) = SMth$
              EMth$(v) = EMth$



              IF VehValue#(v) > 20000 THEN VehValue#(v) = 20000 'Maximum of 20,000

            ' Calculate Tax Paid
              TaxPaid@ = (VehValue#(v) / 100) * TaxRate!
              TaxPaid@ = INT((TaxPaid@ * 100) + .5) / 100
              TaxPaid$ = LTRIM$(STR$(TaxPaid@ * 100))
              TaxPaid$ = LEFT$(TaxPaid$, LEN(TaxPaid$) - 2) + "." + RIGHT$(TaxPaid$, 2)

              PPTaxPd$(v) = TaxPaid$

            ' Calculate Reimbursement
              IF VehValue#(v) <= 1000 THEN
               Reimbursement@ = VAL(TaxPaid$)
               Reimbursement$ = TaxPaid$
               ELSE
               IF VehValue#(v) <= 20000 THEN
                Reimbursement@ = TaxPaid@ * Perc!
                Reimbursement@ = INT((Reimbursement@ * 100) + .5) / 100
                Reimbursement$ = LTRIM$(STR$(Reimbursement@))
               END IF
              END IF

              Reimbursement$ = LTRIM$(STR$(Reimbursement@ * 100))
              Reimbursement$ = LEFT$(Reimbursement$, LEN(Reimbursement$) - 2) + "." + RIGHT$(Reimbursement$, 2)

              PPTaxReimb$(v) = Reimbursement$
              TotalReimb# = TotalReimb# + Reimbursement@
              TotalReimb# = INT((TotalReimb# * 100) + .5) / 100
              veh = v

             END IF

            PropertyRecord! = PersRec(1).NextRec
          WEND
        END IF

       END IF
      END IF
      ShowPctComp Cnt, NumOfTaxRecs
      IF veh > 180 THEN
         Early = 1
         GOTO RDEndDMVProcess
      END IF

  NEXT

RDEndDMVProcess:
  'Write Header
   GOSUB RDDMVHeaderWrite
  'Write Records
   GOSUB RDDMVLines
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

RDDMVExitJournal:

  EXIT SUB

RDDMVLines:
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
  PRINT #10, RTRIM$(Vin$(Records)) + "@";
   VehValue$ = LTRIM$(STR$(VehValue#(Records)))
  PRINT #10, RTRIM$(VehValue$) + "@";
  PRINT #10, RTRIM$(PPTaxPd$(Records)) + "@";
  PRINT #10, RTRIM$(PPTaxReimb$(Records)) + "@";
  PRINT #10, LTRIM$(TaxYear$) + SMth$(Records) + "@";
  PRINT #10, LTRIM$(TaxYear$) + EMth$(Records) + "@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RTRIM$(Today$) + "@"

  
  
  Nme$ = RTRIM$(FirstName1$(Records)) + " " + RTRIM$(LastName1$(Records))
  PRINT #RptHandle, LEFT$(Nme$, 38);
  PRINT #RptHandle, TAB(40); LEFT$(Vin$(Records), 20);
  PRINT #RptHandle, TAB(61); USING "####,#.##"; VAL(PPTaxPd$(Records));
  PRINT #RptHandle, TAB(71); USING "####.##"; VAL(PPTaxReimb$(Records))
  lc = lc + 1
  IF lc >= 56 THEN
   PRINT #RptHandle, CHR$(12);
   GOSUB RDReportHeading
  END IF

  NEXT Records
  RETURN

RDDMVHeaderWrite:
  TotalAmt$ = LTRIM$(STR$(TotalReimb# * 100))
  TLen = LEN(TotalAmt$)

  TotalCrCnt$ = ""
  TotalCrAmt$ = ""

  IF VAL(TotalAmt$) = 0 THEN
    TotalAmt$ = "0.00"
   ELSE
    TotalAmt$ = LEFT$(TotalAmt$, TLen - 2) + "." + RIGHT$(TotalAmt$, 2)
  END IF
  Vehicles$ = LTRIM$(STR$(veh))
  PRINT #10, "H@";
  PRINT #10, "1@";
  PRINT #10, Jury$ + "@";
  PRINT #10, RIGHT$(DATE$, 4) + "@";
  'PRINT #10, WhatYear$ + "@";
  PRINT #10, RTRIM$(Today$) + "@";
  PRINT #10, RTRIM$(Vehicles$) + "@";
  PRINT #10, TotalAmt$ + "@";
  PRINT #10, TotalCrCnt$ + "@";
  PRINT #10, TotalCrAmt$
  RETURN


RDReportHeading:
 PRINT #RptHandle, "** DMV Processing : Data File Contents"
 PRINT #RptHandle, "Submission Date: "; Form$(2, 0); TAB(60); "Batch #"; Batch$
 PRINT #RptHandle, ""
 PRINT #RptHandle, "Name"; TAB(40); "VIN #"; TAB(60); "Tax Paid"; TAB(70); "Reimb Amt"
 PRINT #RptHandle, STRING$(79, "-")
 lc = 5
 RETURN
RDpptra:
 WhatYear = VAL(TaxYear$)
 IF WhatYear = 1998 THEN Perc! = .125: WhatYear$ = "1998"
 IF WhatYear = 1999 THEN Perc! = .275: WhatYear$ = "1999"
 IF WhatYear = 2000 THEN Perc! = .475: WhatYear$ = "2000"
 IF WhatYear = 2001 THEN Perc! = .7: WhatYear$ = "2001"
 IF WhatYear = 2002 THEN Perc! = 1!: WhatYear$ = "2002"
 RETURN

END SUB

SUB DMVTest

END SUB

