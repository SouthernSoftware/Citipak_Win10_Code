DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB PrintMeterSheets
  BlockClear
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetupRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetupRec(), UBSetUpLen

  IF INSTR(UBSetupRec(1).UtilName, "WHITE LAKE") THEN
     WhiteLakeFlag = 1
  END IF
  
  IF UBSetupRec(1).UseSeq = "Y" THEN
    MakeSequenceIndex "Sequence Number"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&("UBTEMP.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBTEMP.IDX", IdxBuff(1), 4, IdxNumOfRecs
    UseSeq = True
  ELSE
    ShowProcessingScrn "Scanning Accounts"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&("UBCUSTBK.IDX")
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    UBCustFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
    GET UBCustFile, IdxBuff(IdxNumOfRecs).RecNum, UBCustRec(1)
    Book2$ = UBCustRec(1).BOOK
    SeqNumb2$ = UBCustRec(1).SEQNUMB
    FOR Cnt& = 1 TO IdxNumOfRecs
      GET UBCustFile, IdxBuff(Cnt&).RecNum, UBCustRec(1)
    
      IF LEN(QPTrim$(UBCustRec(1).BOOK)) > 0 THEN
        Book1$ = UBCustRec(1).BOOK
        SeqNumb1$ = UBCustRec(1).SEQNUMB
        EXIT FOR
      END IF
      ShowPctComp Cnt&, IdxNumOfRecs
    NEXT
    CLOSE UBCustFile
    UseSeq = False
  END IF
  
  ReportFile$ = "UBMTRSHT.RPT"
  REDIM TempScrn(0)
  
  IF UseSeq THEN
    GOTO SeqJump
  END IF
  
  
  LibName$ = "UB"
  ScrnName$ = "UBMTRSHT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  
  FirstTime = True
  BlockClear
  DisplayUBScrn ScrnName$
  
  'Frm(1).StayOnField = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = Book1$
      LSET Form$(2, 0) = SeqNumb1$
      LSET Form$(3, 0) = Book2$
      LSET Form$(4, 0) = SeqNumb2$
      Action = 1
    END IF
    
    IF Frm(1).FldNo = 3 AND Frm(1).PrevFld < 3 THEN             '>2
      In1 = True
      GOSUB CheckBookSequence
    ELSEIF Frm(1).FldNo > 4 AND Frm(1).PrevFld = 4 THEN         '>2
      In1 = False
      GOSUB CheckBookSequence
    END IF
    
    SELECT CASE Frm(1).KeyCode  '--Check for Key presses
      
    CASE F10Key
      'Check for valid Order of Route Questions
      In1 = True
      GOSUB CheckBookSequence
      IF OkFlag THEN
        In1 = False
        GOSUB CheckBookSequence
      END IF
      IF OkFlag THEN
        Book1& = QPValL(Form$(1, 0))
        Sequ1& = QPValL(Form$(2, 0))
        Book2& = QPValL(Form$(3, 0))
        Sequ2& = QPValL(Form$(4, 0))
        Done = True
      END IF
      
    CASE F5KEY
      GOSUB DoSheetMask
      
    CASE EscKey
      GOTO ExitPrintSheets
      
    END SELECT
    
  LOOP UNTIL Done
  
  'Free Up Some Memory
  ERASE Form$, Fld, Frm
  
SeqJump:
  
  ' Location Order ********************************************************
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle     'Open Report File
  
  ShowProcessingScrn "Reading Meter Information"
  
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    AcctNumber& = IdxBuff(LCnt&).RecNum
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipSheet
    END IF
    
    IF UseSeq = False THEN
      BOOK = QPValL(UBCustRec(1).BOOK)
      Sequ& = QPValL(UBCustRec(1).SEQNUMB)
      IF BOOK < Book1& OR BOOK > Book2& THEN
        GOTO SkipSheet
      END IF
      IF Sequ& < Sequ1& OR Sequ& > Sequ2& THEN
        GOTO SkipSheet
      END IF
      GOSUB PrintEm
    ELSE
      GOSUB PrintEm
    END IF
    
SkipSheet:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
   NEXT
  
  CLOSE
  
  IF AbortFlag GOTO ExitPrintSheets
  
  Header$ = "Meter Reading Sheets"
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  GOTO ExitPrintSheets
  
PrintEm:
  GOSUB GetMeterFlags
  GOSUB LookForSecLights
  GOSUB LookForTrash
  IF WhiteLakeFlag <> 1 THEN
   GOSUB PrintSkipHeader
  END IF

  ZONE$ = QPTrim$(UBCustRec(1).ZONE)
  SELECT CASE LEFT$(QPTrim$(UBCustRec(1).CUSTTYPE), 1)
  CASE "B"
    CustT$ = "Commerical"
  CASE "R"
    CustT$ = "Residential"
  CASE ELSE
    CustT$ = "??????????"
  END SELECT


  IF WhiteLakeFlag = 1 THEN
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "Acct #"; AcctNumber&
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); TAB(15); UBCustRec(1).CUSTNAME
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); UBCustRec(1).ADDR1
   PRINT #RptHandle, TAB(15); UBCustRec(1).ADDR2
   PRINT #RptHandle, TAB(15); QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "SERVICE AT "; UBCustRec(1).SERVADDR
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(14); RIGHT$(DATE$, 2); TAB(43); UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB
   FOR zz = 16 TO 43
   PRINT #RptHandle,
   NEXT zz
   PRINT #RptHandle, TAB(15); WatRead&
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); WatSer$
   FOR zz = 51 TO 59
    PRINT #RptHandle,
   NEXT zz
    PRINT #RptHandle,
    RETURN
  END IF
  



  IF EleFlag THEN
    PRINT #RptHandle, " Electric"; "  "; ECode$
  ELSE
    PRINT #RptHandle, ""
  END IF
  IF WatFlag THEN
    PRINT #RptHandle, " Water"; "  "; ZONE$; "  Min ="; WatMin
  ELSE
    PRINT #RptHandle, ""
  END IF
  IF SewFlag THEN
    PRINT #RptHandle, " Sewer"; "  "; ZONE$; "  Min ="; SewMin
  ELSE
    PRINT #RptHandle, ""
  END IF
  
  IF SecFlag THEN
    PRINT #RptHandle, " Security Lights   #"; SecCnt
  ELSE
    PRINT #RptHandle, ""
  END IF
  
  IF TrashFlag THEN
    PRINT #RptHandle, " Trash   "; TCode$
  ELSE
    PRINT #RptHandle, ""
  END IF
  
  PRINT #RptHandle, TAB(15); EleMin
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, EleSer$; TAB(24); WatSer$
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(6); USING "#########"; EleRead&; TAB(24); WatRead&
  FOR zz = 21 TO 43
    PRINT #RptHandle,
  NEXT
  PRINT #RptHandle, UBCustRec(1).CUSTNAME
  PRINT #RptHandle, UBCustRec(1).ADDR1
  PRINT #RptHandle, UBCustRec(1).SERVADDR
  PRINT #RptHandle, QPTrim$(UBCustRec(1).CITY); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(11); UBCustRec(1).BOOK; "-"; UBCustRec(1).SEQNUMB
  PRINT #RptHandle, "~"
  RETURN
  
  
LookForSecLights:
  SecFlag = False
  FOR FRCnt = 1 TO 4
    IF INSTR(UBCustRec(1).FlatRates(FRCnt).FRDESC, "SECUR") THEN
      SecFlag = True
      SecCnt = UBCustRec(1).FlatRates(FRCnt).NumMin
      EXIT FOR
    END IF
  NEXT
  RETURN
  
LookForTrash:
  TrashFlag = False
  IF LEN(QPTrim$(UBCustRec(1).Serv(9).RateCode)) > 0 THEN
    TrashFlag = True
    TCode$ = UBCustRec(1).Serv(9).RateCode
  END IF
  RETURN
  
PrintSkipHeader:
  FOR zz = 1 TO 9
    PRINT #RptHandle,
  NEXT
  RETURN
  
GetMeterFlags:
  WatFlag = False: WatMin = 0: WatSer$ = "": WatRead& = 0
  SewFlag = False: SewMin = 0:
  EleFlag = False: EleMin = 0: EleSer$ = "": EleRead& = 0
  
  FOR MtrCnt = 1 TO 7
    SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MTRType
    CASE "C"
      WatFlag = True
      SewFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "W"
      WatFlag = True
      WatMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "S"
      SewFlag = True
      SewMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      WatSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      WatRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      IF LEN(WatSer$) = 0 THEN
        WatSer$ = "?????"
      END IF
    CASE "E"
      EleFlag = True
      EleMin = UBCustRec(1).LocMeters(MtrCnt).NumUser
      EleSer$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
      EleRead& = UBCustRec(1).LocMeters(MtrCnt).CurRead
      ECode$ = UBCustRec(1).Serv(3).RateCode
      IF LEN(EleSer$) = 0 THEN
        EleSer$ = "?????"
      END IF
    END SELECT
  NEXT
  IF WatRead& < 0 THEN
    WatRead& = 0
  END IF
  IF EleRead& < 0 THEN
    EleRead& = 0
  END IF
  RETURN
  
DoSheetMask:
  SaveScrn TempScrn()
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle     'Open Report File
  IF WhiteLakeFlag = 1 THEN
   PRINT #RptHandle, "TOP"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "Acct # XXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(14); RIGHT$(DATE$, 2); TAB(43); "XX-XXXXXXX"
   FOR zz = 16 TO 43
   PRINT #RptHandle,
   NEXT zz
   PRINT #RptHandle, TAB(15); "XXXXXXXXXX"
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(15); "XXXXXXXXXXX"
   FOR zz = 51 TO 59
   PRINT #RptHandle,
   NEXT zz
   PRINT #RptHandle, "BOTTOM"
   ELSE
  PRINT #RptHandle, "TOP"
  FOR zz = 1 TO 8
    PRINT #RptHandle, ""
  NEXT
  PRINT #RptHandle, " Electric  XXXXXXXXXX"
  PRINT #RptHandle, " Water  X  Min = X"
  PRINT #RptHandle, " Sewer  X  Min = X"
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  PRINT #RptHandle, "               X"
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  PRINT #RptHandle, "XXXXXXXXX              XXXXXXXXX"
  PRINT #RptHandle, ""
  PRINT #RptHandle, ""
  PRINT #RptHandle, "       XXXXXXX           XXXXXXX"
  FOR zz = 1 TO 23
    PRINT #RptHandle, ""
  NEXT
  PRINT #RptHandle, "XXXXXXX XXXXXXXXXXX"
  PRINT #RptHandle, "XX XXX XXX"
  PRINT #RptHandle, "XXXXXXXXXXX"
  PRINT #RptHandle, "XXXXXXXXX XX XXXXX"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "          XX-XXXXXX"
  PRINT #RptHandle, "BOTTOM"
  END IF
  CLOSE RptHandle
  
  CALL CursorOff
  
  PrintRptFile Header$, ReportFile$, 1, RetCode%, 4
  RestScrn TempScrn()
  Action = 1
  
  RETURN
  
CheckBookSequence:
  OkFlag = False
  IF In1 THEN
    Fld1 = 1
    Fld2 = 2
  ELSE
    Fld1 = 3
    Fld2 = 4
  END IF
  Form$(Fld1, 0) = FmtBook$(Form$(Fld1, 0))
  Form$(Fld2, 0) = FmtSeqN$(Form$(Fld2, 0))
  IF Chk4BookSeqNum(Form$(Fld1, 0), Form$(Fld2, 0)) = 0 THEN
    EditedFlag = False
    Frm(1).KeyCode = 0
    SaveScrn TempScrn()
    Ok = MsgBox%("UB", "BADACCTN")
    RestScrn TempScrn()
    IF In1 THEN
      Frm(1).FldNo = 1
    ELSE
      Frm(1).FldNo = 3
    END IF
  ELSE
    Frm(1).PageDisable = False
    OkFlag = True
  END IF
  Action = 2
  
  RETURN
  
ExitPrintSheets:
END SUB

