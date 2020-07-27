DEFINT A-Z
DECLARE SUB ConditionReport ()
DECLARE SUB MasterReadReport ()
  
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Chk4BookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (FileHand%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayHHScrn (ScrnName$)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE SUB EstMeterReading ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB InactiveConsReport ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintHighLowReport ()
DECLARE SUB PrintMeterList ()
DECLARE SUB PrintMeterNotes ()
DECLARE SUB PrintMeterReport ()
DECLARE SUB PrintMeterSheets ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB ProcessHusky ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (TempScrn())
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubhusky.bi'
  '$INCLUDE: 'Newcust.bi'
  
  CONST False = 0, True = NOT False
  
  REDIM MChoice$(1 TO 10)
  
  MChoice$(1) = "Manual Meter Reading Entry "
  MChoice$(2) = "Hand-Held Meter Reading Entry"
  MChoice$(3) = "Estimated Meter Reading Entry"
  MChoice$(4) = "Print High/Low Report"
  MChoice$(5) = "Print Conditions Report"
  MChoice$(6) = "Print Master Reading Report"
  MChoice$(7) = "Inactive Consumption Report"
  MChoice$(8) = "Print Meter Reading Notes"
  MChoice$(9) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit box to go no lower than line 17
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt! = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt!))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) + 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    
    LOCATE Row, Col, 0
    
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Meter Reading Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      'ManualMeterReading             'unrem
    CASE 2
      ProcessHusky
    CASE 3
      EstMeterReading           'unrem
    CASE 4
      PrintHighLowReport        'unrem
    CASE 5
      ConditionReport
    CASE 6
      MasterReadReport          'unrem
    CASE 7
      InactiveConsReport        'unrem
    CASE 8
      PrintMeterNotes           'unrem
    CASE 9
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  RUN "ubbillin"

SUB DisplayHHScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB PrintHelp (Help$)
  
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
  
END SUB

SUB ProcessHusky
  REDIM TempScrn(0)
  
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen
  
  HighVar# = UBSetupRec(1).HighRead
  LowVar# = UBSetupRec(1).LowRead
  
  IF Exist("UBHUSKY.CFG") THEN
    CFGFile = FREEFILE
    OPEN "UBHUSKY.CFG" FOR INPUT AS #CFGFile
    LINE INPUT #CFGFile, Port$
    CLOSE CFGFile
  ELSE
    CLS
    PRINT "CAN NOT FIND THE HUSKY CONFIGURATION FILE:  UBHUSKY.CFG"
    PRINT
    PRINT "Please call Southern Software for support. . ."
    PRINT
    PRINT "Press any key to EXIT.";
    WaitForAction
    EXIT SUB
  END IF
  
  q$ = CHR$(34)
  cb$ = SPACE$(45)
  
  SHARED Choice$()
  REDIM Route(100)
  
  LibName$ = "UBHH"
  ScrnName$ = "UBHUSKHH"
  
  'CursorOff
  'Define Fields
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
  
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Send Info to HUSKY"
  Choice$(2, 0) = "Get Info From HUSKY"
  
  Action = 1
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      Operation$ = QPTrim$(Form$(1, 0))
      IF LEN(Operation$) = 0 THEN
        SaveScrn TempScrn()
        DisplayUBScrn "ERRSCRN1"
        QPrintRC "Invalid HandHeld Operation.", 10, 27, -1
        QPrintRC "Press any key to Continue.", 13, 28, -1
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 1
        Done = False
      ELSE
        Done = True
      END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  SELECT CASE LEFT$(Operation$, 1)
  CASE "S"
    GOSUB SendHusky
  CASE "G"
    GOSUB GetHusky
  CASE ELSE
  END SELECT
  
  EXIT SUB
  
SendHusky:
  LibName$ = "UBHH"
  ScrnName$ = "UBHUSK1S"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  'LOOK HERE
  
  REDIM Choice$(5, 0)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Water/Sewer"
  Choice$(2, 0) = "Electric"
  Choice$(3, 0) = "Gas Meters"
  Choice$(4, 0) = "ALL METERS"
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  NumberofRoutes = 0
  
  DisplayHHScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 13 AND Frm(1).FldNo > 2 THEN
      IF VAL(Form$(2, 0)) <> 0 THEN
        NumberofRoutes = NumberofRoutes + 1
        IF NumberofRoutes <= 10 THEN
          Form$(NumberofRoutes + 2, 0) = Form$(2, 0)
          Frm(1).FldNo = 2
          Form$(2, 0) = ""
          Action = 1
        END IF
      END IF
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      WhatTypes$ = LEFT$(Form$(1, 0), 1)
      ' Next Check Routes
      IF NumberofRoutes > 0 THEN
        QPrintRC "      Scanning Account:        ", 11, 21, -1
        ' Set Route Number Values
        FOR Cnt = 1 TO NumberofRoutes
          Route(Cnt) = VAL(Form$(Cnt + 2, 0))
        NEXT
        'Open Customer Data File
        REDIM UBCustRec(1) AS NewUBCustRecType
        UBCustRecLen = LEN(UBCustRec(1))
        UBFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
        
        'Open and Initialize the Husky Read Information File
        REDIM UBHuskyHHRec(1) AS UBHuskyHHRecType
        UBHuskyRdRecLen = LEN(UBHuskyHHRec(1))
        UBHuskyRdFile = FREEFILE
        KillFile "UBCUSTTR.DAT"
        UBHuskyRdFile = FREEFILE
        OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS UBHuskyRdFile LEN = UBHuskyRdRecLen
        NumHuskyRdRecs = LOF(UBHuskyRdFile) / UBHuskyRdRecLen
        
        'Open the Correct Order for Reading
        IF UBSetupRec(1).UseSeq = "Y" THEN
          IdxRecLen = 4         'we are using a integer
          MakeSequenceIndex "Sequence Number"
          IdxNumOfRecs = FileSize&("UBTEMP.IDX") \ 4
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
        ELSE
          IdxRecLen = 4         'we are using a integer
          IdxFileSize& = FileSize&("UBCUSTBK.IDX")
          IdxNumOfRecs = IdxFileSize& \ IdxRecLen
          REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
          FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs            'load it
        END IF
        
        Cnt! = 1
        DO
          PRec& = IdxBuff(Cnt!).RecNum
          IF NOT (PRec&) = 0 THEN
            GET UBFile, PRec&, UBCustRec(1)
            IF INSTR(UBCustRec(1).HHMsg1, "NOREAD") > 0 THEN
              GOTO HWriteSkip
            END IF
            DisplayAcct$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
            QPrintRC DisplayAcct$, 11, 45, -1
            FOR BookCnt = 1 TO NumberofRoutes
              IF VAL(UBCustRec(1).Book) = Route(BookCnt) AND (UBCustRec(1).Status <> "F") THEN
                GOSUB HuskyWriteRec
              END IF
            NEXT
          END IF
HWriteSkip:
          Cnt! = Cnt! + 1
        LOOP UNTIL Cnt! > IdxNumOfRecs
        CLOSE
        Ok = MsgBox%("UBSETUP", "UBHUSK1C")
        
        IF Ok = 1 THEN
          DisplayHHScrn "UBHUSK3C"
          QPrintRC "Transfering file to the HUSKY.", 10, 25, -1
          LOCATE 14, 40
          IF WEnvTest = 0 THEN
            SHELL "cmdhcom " + Port$ + " /tx=ubcusttr.dat /noscreen /abort"
          END IF
          'WaitForAction
          QPrintRC "    File Transfer Complete.   ", 10, 25, 15
          QPrintRC "  Press any key to continue.  ", 14, 25, 15
          BEEP
          WaitForAction
          Done = True
        ELSE
          DisplayHHScrn "UBHUSK3C"
          QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
          QPrintRC "  Press any key to continue.  ", 14, 25, -1
          BEEP
          WaitForAction
          Done = True
        END IF
      END IF
      
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
  RETURN
  
GetHusky:
  
  Ok = MsgBox%("UBSETUP", "UBHUSK1C")
  IF Ok = 1 THEN
    DisplayHHScrn "UBHUSK3C"
    QPrintRC "Transfering file from the HUSKY.", 10, 24, -1
    LOCATE 14, 40
    IF WEnvTest = 0 THEN
      SHELL "cmdhcom " + Port$ + " /RX=ubcusttr.dat /noscreen /abort"
    END IF
    'WaitForAction
    QPrintRC "     File Transfer Complete.    ", 10, 24, 15
    QPrintRC "   Press any key to continue.   ", 14, 24, 15
    BEEP
    WaitForAction
    Done = True
  ELSE
    DisplayHHScrn "UBHUSK3C"
    QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
    QPrintRC "  Press any key to continue.  ", 14, 25, -1
    BEEP
    WaitForAction
    Done = True
    GOTO HuskyGetExit:
  END IF
  
  '    'Open Customer Data File
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBFile LEN = UBCustRecLen
  
  'Open and Initialize the Husky Read Information File
  REDIM UBHuskyGetRDRec(1) AS UBHuskyHHRecType
  UBHuskyGetRdRecLen = LEN(UBHuskyGetRDRec(1))
  UBHuskyGetRdFile = FREEFILE
  FileNam$ = "UBCUSTTR.DAT"
  OPEN FileNam$ FOR RANDOM ACCESS READ WRITE SHARED AS UBHuskyGetRdFile LEN = UBHuskyGetRdRecLen
  NumHuskyGetRdRecs = LOF(UBHuskyGetRdFile) / UBHuskyGetRdRecLen
  
  IF NumHuskyGetRdRecs = 0 THEN
    CLOSE
    QPrintRC "                                ", 10, 24, 15
    QPrintRC "                                ", 14, 24, 15
    QPrintRC "   ERROR: NO READINGS FOUND   ", 9, 25, -1
    QPrintRC " Check the handheld connection", 11, 25, -1
    QPrintRC " and try the transfer again!!!", 13, 25, -1
    QPrintRC "   Press any key to continue. ", 15, 25, 14
    BEEP: BEEP: BEEP
    WaitForAction
    GOTO GetHusky
  END IF
  
  QPrintRC "  Processing Meter Readings.  ", 9, 25, 15
  QPrintRC "                              ", 10, 25, 11
  QPrintRC "    Updated Count:            ", 11, 25, 11
  QPrintRC "     Unread Count:  0         ", 12, 25, 11
  QPrintRC "                              ", 14, 25, 11
  
  Cnt! = 1      ' Initialize File Counter to 1
  
  DO
    GET UBHuskyGetRdFile, Cnt!, UBHuskyGetRDRec(1)
    ' New Way Undelete When Our System Sends and Retrieves
    PRec& = VAL(QPTrim$(UBHuskyGetRDRec(1).Account))
    
    IF NOT (PRec&) = 0 THEN
      GET UBFile, PRec&, UBCustRec(1)
      GOSUB HuskyExtractRecord
    END IF
    Cnt! = Cnt! + 1
    
  LOOP UNTIL Cnt! > NumHuskyGetRdRecs
  
  CLOSE
  Done = True
  
  QPrintRC " Readings Updated Successfully. ", 13, 24, 15
  QPrintRC "  Press any key to continue.  ", 15, 25, 14
  BEEP
  WaitForAction
  Done = True
  
HuskyGetExit:
  RETURN
  
HuskyExtractRecord:
  UpdCnt = UpdCnt + 1
  QPrintRC STR$(UpdCnt), 11, 44, 15
  'New Way To Activate
  MeterRecord = VAL(RIGHT$((QPTrim$(UBHuskyGetRDRec(1).Account)), 1))
  
  ' Check Meter Updated Flag
  ' Update Meter W/Reading

  IF UBHuskyGetRDRec(1).ReadFlag = "Y" THEN
    CurReading# = UBHuskyGetRDRec(1).CurRead
    IF UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y" THEN
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
      UBCustRec(1).LocMeters(MeterRecord).CurDate = UBHuskyGetRDRec(1).ReadDate
    ELSE
      UBCustRec(1).LocMeters(MeterRecord).PrevRead = UBCustRec(1).LocMeters(MeterRecord).CurRead
      UBCustRec(1).LocMeters(MeterRecord).PastDate = UBCustRec(1).LocMeters(MeterRecord).CurDate
      UBCustRec(1).LocMeters(MeterRecord).ReadFlag = "Y"
      UBCustRec(1).LocMeters(MeterRecord).CurDate = UBHuskyGetRDRec(1).ReadDate
      UBCustRec(1).LocMeters(MeterRecord).CurRead = CurReading#
    END IF
  'Modifed 04-28-97
    IF UBHuskyGetRDRec(1).NoteStatus = "P" THEN
      UBCustRec(1).NewNotes = True
      UBCustRec(1).HHMsg1 = UBHuskyGetRDRec(1).Notes(1).Note
      UBCustRec(1).HHMsg2 = UBHuskyGetRDRec(1).Notes(2).Note
      UBCustRec(1).HHMsg3 = UBHuskyGetRDRec(1).Notes(3).Note
    END IF
    '   ^^^
    PUT UBFile, PRec&, UBCustRec(1)
  ELSE
    UnRead = UnRead + 1
    
  END IF

  QPrintRC STR$(UnRead), 12, 44, 15

RETURN
  
  
HuskyWriteRec:
  'May Have Up to 7 Meters to Read
  FOR MtrCount = 1 TO 7
    MeterOK = False
    Account$ = STR$(PRec&)
    Account$ = LEFT$(Account$, 6) + "-" + RIGHT$(STR$(MtrCount), 1)
    
    IF ASC(UBCustRec(1).LocMeters(MtrCount).MtrType) > 32 THEN
      MtrType$ = UBCustRec(1).LocMeters(MtrCount).MtrType
      IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" OR MtrType$ = "E" OR MtrType$ = "D" OR MtrType$ = "G" THEN
        SELECT CASE WhatTypes$
        CASE "W"                'water/sewer
          IF MtrType$ = "W" OR MtrType$ = "S" OR MtrType$ = "C" THEN
            MeterOK = True
          END IF
        CASE "E"                'electric & demand elec.
          IF MtrType$ = "E" OR MtrType$ = "D" THEN
            MeterOK = True
          END IF
        CASE "G"                'gas
          IF MtrType$ = "G" THEN
            MeterOK = True
          END IF
        CASE "A", " "           'all meters
          MeterOK = True
        END SELECT
        
        IF MeterOK = True THEN
          ' Determine High and Low Reading
          Average# = UBCustRec(1).LocMeters(MtrCount).AvgUse

          MinUsed# = UBCustRec(1).MeterAux(MtrCount).MinUsed
          MinUsed# = MinUsed# + 1

          MaxUsed# = UBCustRec(1).MeterAux(MtrCount).MaxUsed
          IF MaxUsed# < MinUsed# THEN
            MaxUsed# = MinUsed#
          END IF

          LowRead# = FIX(MinUsed# * (LowVar# / 100))
          HighRead# = FIX(MaxUsed# * (HighVar# / 100))
          
          MeterID$ = QPTrim$(UBCustRec(1).LocMeters(MtrCount).MTRNUM)
          IF LEN(MeterID$) = 0 THEN
            MeterID$ = UBCustRec(1).Book + UBCustRec(1).SeqNumb
          END IF
          'IF LEN(MeterID$) < 8 THEN
          '  MeterID$ = STRING$(8 - LEN(MeterID$), "0") + MeterID$
          'END IF
          'MeterID$ = LEFT$(MeterID$, 8)
          'Set Record Fields and Put On Disk
          UBHuskyHHRec(1).Custname = UBCustRec(1).Custname
          UBHuskyHHRec(1).ServAddress = UBCustRec(1).ServAddr
          C$ = QPTrim$(UBCustRec(1).USERCODE1)
          IF LEN(C$) > 0 THEN
            UBHuskyHHRec(1).UCode1 = LEFT$(C$, 2)
          END IF
          C$ = QPTrim$(UBCustRec(1).UserCode2)
          IF LEN(C$) > 0 THEN
            UBHuskyHHRec(1).UCode2 = LEFT$(C$, 2)
          END IF
          
          UBHuskyHHRec(1).MeterID = MeterID$
          UBHuskyHHRec(1).LowRead = LowRead#
          UBHuskyHHRec(1).HighRead = HighRead#
          UBHuskyHHRec(1).Account = Account$
          UBHuskyHHRec(1).MeterType$ = UBCustRec(1).LocMeters(MtrCount).MtrType
          UBHuskyHHRec(1).Book = VAL(UBCustRec(1).Book)
          UBHuskyHHRec(1).CurRead = UBCustRec(1).LocMeters(MtrCount).CurRead
          UBHuskyHHRec(1).PastRead = 0
          UBHuskyHHRec(1).ReadFlag = "N"
          'Modifed 04-28-97
          
          UBHuskyHHRec(1).Notes(1).Note = UBCustRec(1).HHMsg1
          UBHuskyHHRec(1).Notes(2).Note = UBCustRec(1).HHMsg2
          UBHuskyHHRec(1).Notes(3).Note = UBCustRec(1).HHMsg3
          '         ^^^
          UBHuskyHHRec(1).NoteStatus = ""
          PUT UBHuskyRdFile, (LOF(UBHuskyRdFile) / UBHuskyRdRecLen) + 1, UBHuskyHHRec(1)
        END IF
      END IF
    END IF
  NEXT MtrCount
  RETURN
  
END SUB

