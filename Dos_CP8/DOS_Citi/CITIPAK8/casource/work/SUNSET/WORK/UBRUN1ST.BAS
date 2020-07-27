DEFINT A-Z
DECLARE SUB CheckDepTrans ()
DECLARE SUB CheckCustFile ()
DECLARE SUB LoadInfo ()
DECLARE SUB CheckOwnerFile ()
DECLARE SUB ClrBlock ()
DECLARE SUB ShowNewRec (NRate AS ANY)
DECLARE SUB ShowOldRec (ORate AS ANY)
DECLARE SUB BackUpRateTbl ()
DECLARE SUB CheckSetupFile ()
DECLARE SUB CheckRateFile ()
DECLARE SUB CheckDraftFile ()
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE FUNCTION BiosKey% ()
DECLARE FUNCTION FUsing$ (Number$, Fmt$)
  
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubowner.BI'
  '$INCLUDE: 'ubtrans.BI'

  CONST False = 0, True = NOT False
  
  DIM SHARED AbortFlag, Fmt$
  
  Fmt$ = "##########.##"
  
  STACK 8000

'  NAME "UBLOG.DAT" AS "UBLG.DAT"

  CLS
  PRINT
  PRINT "    UB File Checking/Conversion Utility v3.10"
  PRINT "      (C)1997-1999 Southern Software Inc."
  
  'GOTO MakeRateFile

  CheckDraftFile
  CheckRateFile
  CheckSetupFile
  CheckOwnerFile
  CheckCustFile
  CheckDepTrans

  PRINT
  PRINT

  'LOCATE 15, 1
  
  IF AbortFlag THEN
    PRINT "  PROCESSING ABORTED!!!"
    PRINT "  PLEASE CALL CUSTOMER SUPPORT AT SOUTHERN SOFTWARE"
    PRINT "  BEFORE CONTINUING."
  ELSE
    PRINT " Processing Completed."
  END IF
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT " Press any key to continue.";
'END
  a = BiosKey
  
  IF AbortFlag = False THEN
    IF INSTR(COMMAND$, "TEST") = 0 THEN
      KillFile "UBRUN1ST.EXE"
      RUN "UBMENU.EXE"
    END IF
  END IF
  
  PRINT
  
  END
  
MakeRateFile:
  KillFile "UBRATE.DAT"
  REDIM ORate(1) AS oUBRateTblRecType
  OLen = LEN(ORate(1))
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = OLen
  FOR Cnt = 1 TO 29
    REDIM ORate(1) AS oUBRateTblRecType
    RCCnt$ = QPTrim$(STR$(Cnt))
    SELECT CASE Cnt
    CASE 1 TO 9
      RC$ = "WA0"
    CASE ELSE
      RC$ = "WA"
    END SELECT
    ORate(1).RateCode = RC$ + RCCnt$
    IF Cnt MOD 2 = 0 THEN
      ORate(1).RATEDESC = "Water Code " + ORate(1).RateCode
    ELSE
      ORate(1).RATEDESC = "Sewer Service " + ORate(1).RateCode
    END IF
    ORate(1).MINAMT = Cnt
    ORate(1).MINUNITS = 1000
    IF Cnt = 2 THEN
      ORate(1).DiscPct = 12.5
    ELSE
      ORate(1).DiscPct = 0
    END IF
    FOR TblCnt = 1 TO 10
      ORate(1).TblBreaks(TblCnt).UNITS = 0
      ORate(1).TblBreaks(TblCnt).UNITAMT = 0
    NEXT
    PUT RateFile, Cnt, ORate(1)
  NEXT
  CLOSE

SUB BackUpRateTbl
  
  FOR Cnt = 1 TO 99
    IF Cnt < 10 THEN
      ORateFile$ = "UBRATE0" + QPTrim$(STR$(Cnt)) + ".DAT"
    ELSE
      ORateFile$ = "UBRATE" + QPTrim$(STR$(Cnt)) + ".DAT"
    END IF
    IF Exist(ORateFile$) = False THEN
      RateFile = FREEFILE
      OPEN "UBRATE.DAT" FOR BINARY SHARED AS RateFile
      ORateFile = FREEFILE
      OPEN ORateFile$ FOR BINARY SHARED AS ORateFile
      RateLen = LOF(RateFile)
      RateRec$ = SPACE$(RateLen)
      GET RateFile, 1, RateRec$
      PUT ORateFile, 1, RateRec$
      CLOSE ORateFile, RateFile
      EXIT FOR
    END IF
  NEXT
  
END SUB

SUB CheckCustFile

  ChkByte$ = CHR$(1)
  PRINT
  PRINT "  UB System Customer file: ";
  DIM CustRec AS NewUBCustRecType
  CustLen = LEN(CustRec)

  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = CustLen
  NumOfCust& = LOF(CustFile) / CustLen
  IF NumOfCust& > 0 THEN
    GET CustFile, 1, CustRec
    IF CustRec.ChkByte = ChkByte$ THEN
      CustOK = True
    END IF
  END IF
  IF CustOK THEN
    PRINT "OK."
  ELSE
    FOR Cnt& = 1 TO NumOfCust&
      GET CustFile, Cnt&, CustRec

      SWAP CustRec.USEDRAFT, CustRec.AcctType
      IF CustRec.USEDRAFT = "Y" THEN
        CustRec.AcctType = "C"
        BANKACCT$ = QPTrim$(CustRec.BANKACCT)
        DashPos = INSTR(BANKACCT$, "-")
        DO WHILE DashPos > 0
          BANKACCT$ = LEFT$(BANKACCT$, DashPos - 1) + MID$(BANKACCT$, DashPos + 1)
          DashPos = INSTR(BANKACCT$, "-")
        LOOP
        CustRec.BANKACCT = BANKACCT$
      END IF
      CustRec.ChkByte = ChkByte$
      PUT CustFile, Cnt&, CustRec
    NEXT
    PRINT "Converted."
  END IF
  CLOSE CustFile


END SUB

SUB CheckDepTrans

  ChkByte$ = CHR$(1)
  PRINT
  PRINT "      UB Transaction file: ";

  Row = CSRLIN
  Col = POS(0)
  NCol = Col + 2

  DIM TranRec AS UBTransRecType
  TranLen = LEN(TranRec)

  TranFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TranLen
  NumOfTran& = LOF(TranFile) / TranLen
  IF NumOfTran& > 0 THEN
    GET TranFile, 1, TranRec
    IF TranRec.ChkByte = ChkByte$ THEN
      TranOK = True
    END IF
  END IF
  IF TranOK THEN
    PRINT "OK."
  ELSE
    FOR Cnt& = 1 TO NumOfTran&
      GET TranFile, Cnt&, TranRec
      LOCATE Row, NCol
      PRINT "Checking:"; Cnt&; " of"; NumOfTran&;
      IF TranRec.TransType = 107 THEN
        TranRec.TransType = 7
      END IF
      TranRec.ChkByte = ChkByte$
      PUT TranFile, Cnt&, TranRec
    NEXT
    LOCATE Row, NCol: PRINT SPACE$(40):
    LOCATE Row, Col
    PRINT "OK."
  END IF
  CLOSE TranFile


END SUB

SUB CheckDraftFile
  
  PRINT
  PRINT
  PRINT "     UB System Draft file: ";
  
  DIM ODraft AS oUBDraftRecType
  DIM ODraft2 AS oUBDraftRecType2
  DIM NDraft AS UBDraftRecType
  
  OLen = LEN(ODraft)
  OLen2 = LEN(ODraft2)
  NLen = LEN(NDraft)
  
  DraftFile = FREEFILE
  OPEN "UBSDRAFT.DAT" FOR BINARY AS DraftFile
  DraftSize = LOF(DraftFile)
  CLOSE
  
  IF DraftSize > 0 THEN
    SELECT CASE DraftSize
    CASE OLen
      DraftFile = FREEFILE
      OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS DraftFile LEN = OLen
      GET DraftFile, 1, ODraft
      CLOSE
      NDraft.BANKDEST = ODraft.BANKDEST
      NDraft.BANKORIG = ODraft.BANKORIG
      NDraft.BANKNAME = ODraft.BANKNAME
      NDraft.BANKLOC = ODraft.BANKLOC
      LSET NDraft.COMPACCT = ""
      NDraft.FEDID = ODraft.FEDID
      NDraft.FEDPREFX = ODraft.FEDPREFX

    CASE OLen2
      DraftFile = FREEFILE
      OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS DraftFile LEN = OLen2
      GET DraftFile, 1, ODraft2
      CLOSE
      NDraft.BANKDEST = ODraft2.BANKDEST
      NDraft.BANKORIG = ODraft2.BANKORIG
      NDraft.BANKNAME = ODraft2.BANKNAME
      NDraft.BANKLOC = ODraft2.BANKLOC
      NDraft.COMPACCT = ODraft2.COMPACCT
      NDraft.FEDID = ODraft2.FEDID
      NDraft.FEDPREFX = ODraft2.FEDPREFX
    CASE NLen
      PRINT "OK."
      GOTO DraftExit
    END SELECT
    DraftFile = FREEFILE
    OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS DraftFile LEN = NLen
    PUT DraftFile, 1, NDraft
    CLOSE
    PRINT "Converted."
  ELSE
    PRINT "FILE SIZE 0 bytes"
  END IF

DraftExit:
END SUB

SUB CheckOwnerFile

  ChkByte$ = CHR$(1)
  PRINT
  PRINT "     UB System Owner file: ";

  DIM CustRec AS NewUBCustRecType
  DIM Owner AS UBOwnerRecType

  CustLen = LEN(CustRec)
  NLen = LEN(Owner)

  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = CustLen
  NumOfCust& = LOF(CustFile) / CustLen
  CLOSE CustFile

  OwnerFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS OwnerFile LEN = NLen
  NumOfOwner& = LOF(OwnerFile) / NLen

  IF NumOfOwner& > 0 THEN
    GET #OwnerFile, 1, Owner
    IF Owner.ChkByte = ChkByte$ THEN
      OwnerOK = True
    END IF
  END IF
  IF OwnerOK THEN
    PRINT "OK."
  ELSE
    Owner.OwnLName = ""
    Owner.OwnFName = ""
    Owner.ADDR1 = ""
    Owner.ADDR2 = ""
    Owner.CITY = ""
    Owner.STATE = ""
    Owner.ZIPCODE = ""
    Owner.HPHONE = ""
    Owner.WPHONE = ""
    Owner.ChkByte = ChkByte$
    FOR Cnt& = 1 TO NumOfCust&
      PUT #OwnerFile, Cnt&, Owner
    NEXT
    PRINT "Reinitialized."
  END IF
  CLOSE OwnerFile


END SUB

SUB CheckRateFile
  
  ChkByte$ = CHR$(1)
  Esc = -27
  a$ = " "
  PRINT

  'LOCATE 8, 4
  PRINT "      UB System Rate file: ";
  
  DIM ORate AS oUBRateTblRecType
  DIM NRate AS UBRateTblRecType
  
  OLen = LEN(ORate)
  NLen = LEN(NRate)
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = OLen
  RateFileSize& = LOF(RateFile)
  CLOSE
  
  IF RateFileSize& > 0 THEN
    'BackUpRateTbl  'Backup the current rate table if there is one
    IF (RateFileSize& MOD NLen) = 0 AND (RateFileSize& MOD OLen) = 0 THEN
      GOTO ProcessSameMOD
    END IF
  ELSE
    PRINT "FILE SIZE 0 bytes"
    GOTO ExitChkRate
  END IF
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = OLen
  GET RateFile, 1, ORate
  CLOSE
  
  'if chkbyte isn't there and RateCnt is evenly Divisible by Old rec len
  IF (RIGHT$(ORate.RATEDESC, 1) <> ChkByte$) THEN
    IF (RateFileSize& MOD OLen) = 0 THEN
ConvertOld:
      BackUpRateTbl             'Backup the current rate table if there is one
      KillFile "oUBRate.DAT"
      NAME "UBRate.DAT" AS "oUBRate.DAT"
      ORateFile = FREEFILE
      OPEN "oUBRate.DAT" FOR RANDOM SHARED AS ORateFile LEN = OLen
      NumOfRate = LOF(ORateFile) / OLen
      NRateFile = FREEFILE
      OPEN "UBRate.DAT" FOR RANDOM SHARED AS NRateFile LEN = NLen
      FOR Cnt = 1 TO NumOfRate
        GET ORateFile, Cnt, ORate
        NRate.RateCode = ORate.RateCode
        NRate.RATEDESC = ORate.RATEDESC
        NRate.ChkByte = ChkByte$
        NRate.MINAMT = ORate.MINAMT
        NRate.MINUNITS = ORate.MINUNITS
        NRate.MaxAmt = 0
        FOR TblCnt = 1 TO 10
          LSET NRate.TblBreaks(TblCnt) = ORate.TblBreaks(TblCnt)
        NEXT
        PUT NRateFile, Cnt, NRate
      NEXT
      PRINT "Converted."
    ELSE
      GOTO AddChkByte
    END IF
  ELSE
    PRINT "OK."
  END IF
  CLOSE
  
  GOTO ExitChkRate
  
ProcessSameMOD:
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = OLen
  NRateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS NRateFile LEN = NLen
  
  GET NRateFile, 1, NRate
  
  'If chkbyte is there, then the file has already been converted
  IF NRate.ChkByte = ChkByte$ THEN
    PRINT "OK."
    CLOSE
    GOTO ExitChkRate
  END IF
  
  LOCATE 8, 1
  PRINT " CAN NOT DETERMINE THE CORRECT RATE STRUCTURE!!"
  PRINT " PLEASE SELECT THE GROUP THAT LOOKS CORRECT."
  
  NumRates = LOF(NRateFile) \ NLen
  IF NumRates > 3 THEN
    Start = NumRates - 2
  ELSE
    Start = 1
  END IF
  Last = NumRates
  
  FOR Cnt = Start TO Last
    DidCnt = DidCnt + 1
    GET RateFile, Cnt, ORate
    ShowOldRec ORate
    GET NRateFile, Cnt, NRate
    ShowNewRec NRate
  NEXT
  
  CLOSE
  
GetInput:
  LOCATE 23, 1
  PRINT "Select Group: (1-2 or ESC)?   "
  WInput a$, "12", 23, 29, ExitCode
  
  IF ExitCode = Esc THEN
    ClrBlock
    LOCATE 8, 1
    PRINT "      UB System Rate file: ERROR! PROCESSING WAS ABORTED!"
    AbortFlag = True
    GOTO ExitChkRate
  END IF
  
  
  SELECT CASE a$
  CASE "1"      'Rates are in old format
    ClrBlock
    LOCATE 8, 1
    PRINT "      UB System Rate file: ";
    GOTO ConvertOld
  CASE "2"
AddChkByte:
    ClrBlock
    LOCATE 8, 1
    PRINT "      UB System Rate file: ";
    BackUpRateTbl               'Backup the current rate table if there is one
    NRateFile = FREEFILE
    OPEN "UBRATE.DAT" FOR RANDOM SHARED AS NRateFile LEN = NLen
    RateCnt = LOF(NRateFile) \ NLen
    FOR Cnt = 1 TO RateCnt
      GET NRateFile, Cnt, NRate
      NRate.ChkByte = ChkByte$
      PUT NRateFile, Cnt, NRate
    NEXT
    CLOSE
    PRINT "ADDED CHECK BYTE."
  CASE ELSE
    GOTO GetInput
  END SELECT
  
ExitChkRate:
  
END SUB

SUB CheckSetupFile

  DoCust = False
  PRINT

  PRINT "     UB System Setup file: ";
  DIM OSetup AS oUBSetupRecType
  DIM NSetup AS UBSetupRecType
  
  OLen = LEN(OSetup)
  NLen = LEN(NSetup)
  
  SetupFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR BINARY AS SetupFile
  SetupSize = LOF(SetupFile)
  CLOSE
  
  IF SetupSize > 0 THEN
    IF SetupSize < NLen THEN
      SetupFile = FREEFILE
      OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS SetupFile LEN = OLen
      GET SetupFile, 1, OSetup
      CLOSE
      NSetup.UTILNAME = OSetup.UTILNAME
      NSetup.DEFCITY = OSetup.DEFCITY
      NSetup.DEFSTATE = OSetup.DEFSTATE
      NSetup.ZIPCODE = OSetup.ZIPCODE
      NSetup.PreByBook = OSetup.PreByBook
      NSetup.RecpPort = OSetup.RecpPort
      NSetup.RECPDEFT = OSetup.RECPDEFT
      NSetup.EstRead = OSetup.EstRead
      NSetup.BANKDFT = OSetup.BANKDFT
      NSetup.UseSeq = OSetup.UseSeq
      NSetup.BILLCYCL = OSetup.BILLCYCL
      NSetup.DefLook = OSetup.DefLook
      NSetup.MethAcct = OSetup.MethAcct
      NSetup.SkipInactive = OSetup.SkipInactive
      NSetup.SkipSeparator = OSetup.SkipSeparator
      NSetup.Make99File = OSetup.Make99File
      NSetup.LowRead = OSetup.LowRead
      NSetup.HighRead = OSetup.HighRead
      NSetup.HHDEVICE = OSetup.HHDEVICE
      FOR ZZ = 1 TO 15
        LSET NSetup.Revenues(ZZ) = OSetup.Revenues(ZZ)
        NSetup.Revenues(ZZ).ProRate = "N"
        LSET NSetup.BillAcct(ZZ) = OSetup.BillAcct(ZZ)
        LSET NSetup.PayAcct(ZZ) = OSetup.PayAcct(ZZ)
        LSET NSetup.DepAcct(ZZ) = OSetup.DepAcct(ZZ)
      NEXT
      OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS SetupFile LEN = NLen
      PUT SetupFile, 1, NSetup
      CLOSE
      PRINT "Converted."
      DoCust = True
    ELSE
      PRINT "OK."
    END IF
  ELSE
    PRINT "FILE SIZE 0 bytes"
  END IF

  IF DoCust THEN
    Row = CSRLIN
    REDIM Cust(1) AS NewUBCustRecType
    CustLen = LEN(Cust(1))
    CustFile = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM AS CustFile LEN = CustLen
    NumOfCust = LOF(CustFile) \ CustLen
    FOR Cnt = 1 TO NumOfCust
      LOCATE Row, 1
      PRINT "    Prorate Percentage:"; Cnt; " of"; NumOfCust;
      GET #CustFile, Cnt, Cust(1)
      Cust(1).ProratePCT = 100
      PUT #CustFile, Cnt, Cust(1)
    NEXT
    CLOSE
    LOCATE Row, 1
    PRINT SPACE$(50);
  END IF

END SUB

SUB ClrBlock
  
  Block$ = SPACE$(80)
  FOR ZZ = 8 TO 23
    LOCATE ZZ, 1
    PRINT Block$
  NEXT
  
END SUB

SUB ShowNewRec (NRate AS UBRateTblRecType)
  
  STATIC WhatLine
  SRow = 17
  WhatLine = WhatLine + 1

  IF WhatLine = 1 THEN
    LOCATE SRow, 1
    PRINT "Group: 2"
  END IF

  LOCATE SRow + WhatLine, 4
  PRINT STR$(WhatLine); ")   "; NRate.RateCode; "        "; LEFT$(NRate.RATEDESC, 20); "  "; FUsing$(STR$(NRate.MINAMT), Fmt$);
  PRINT FUsing$(STR$(NRate.MINUNITS), Fmt$)
END SUB

SUB ShowOldRec (ORate AS oUBRateTblRecType)
  
  STATIC WhatLine
  SRow = 11
  WhatLine = WhatLine + 1

  IF WhatLine = 1 THEN
    LOCATE SRow, 1
    PRINT "Group: 1"
  END IF

  LOCATE SRow + WhatLine, 4
  PRINT STR$(WhatLine); ")   "; ORate.RateCode; "        "; LEFT$(ORate.RATEDESC, 20); "  "; FUsing$(STR$(ORate.MINAMT), Fmt$);
  PRINT FUsing$(STR$(ORate.MINUNITS), Fmt$)
END SUB

SUB WInput (Edit$, GoodKey$, Row, Col, ExitCode)
  'Return Codes:
  ' Enter Key Exit
  '     0
  ' Escape Key Exit
  '   -27 Escape Key was Pressed to exit sub
  '       Edit string reset to original string
  '
  STATIC InsertMode
  
  Original$ = Edit$
  ExitCode = 0
  
  MaxLength = LEN(Edit$)        'max length of input string
  Ptr = LEN(RTRIM$(Edit$))
  IF Ptr = MaxLength THEN Ptr = MaxLength - 1
  
  DO            'main processing loop
    
    LOCATE Row, Col             'update cursor position.
    PRINT Edit$;                'and display the line
    LOCATE Row, Col + Ptr       'update cursor position.
    
    PressedKey = BiosKey        'wait for a key press
    
    SELECT CASE PressedKey      'process the key
      
      '****  Exit keys
    CASE 13     'ENTER - Accept line and, exit
      'Edit$ = QPTrim$(Edit$)
      IF LEN(Edit$) = 0 THEN
        ExitCode = -1
      ELSE
        ExitCode = 13
      END IF
      
      EXIT DO
      
    CASE 27     'ESCAPE - Abort operation exit
      ExitCode = -27
      Edit$ = Original$
      EXIT DO
      
      
    CASE -59, -60, -62, -72, -80                'F4key
      'UpArrow DnArrow F1 F2 F4
      ExitCode = ABS(PressedKey)
      EXIT DO
      
      'add alt exits here
      ' i.e. downarrow
      'End of Exit keys
      
      '****  Cursor keys
    CASE -82    'INSERT - Toggle insert or, overtype mode
      InsertMode = NOT InsertMode
      IF InsertMode THEN
        LOCATE , , , 0, 6
      ELSE      'show appropriate cursor type
        LOCATE , , , 5, 6
      END IF
    CASE -75    'LEFTARROW - Moves cursor left
      IF Ptr THEN Ptr = Ptr - 1
    CASE -77    'RIGHTARROW - Moves cursor right
      IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1
    CASE -71    'HOME - Moves cursor to first position
      Ptr = 0
    CASE -79    'END - Moves cursor to last position
      Ptr = LEN(RTRIM$(Edit$))
      IF Ptr = MaxLength THEN Ptr = MaxLength - 1
      'End of Cursor keys
      
      '****   Editing Keys & Normal Input
    CASE -117   'CONTROLEND - Erase from cursor to end of Edit$
      IF Ptr THEN
        LSET Edit$ = LEFT$(Edit$, Ptr)
      ELSE
        Edit$ = SPACE$(LEN(Edit$))
      END IF
    CASE -22    'ALT-U     - Restore to orginial string
      Edit$ = Original$
      Ptr = 0
      
    CASE 8      'BACKSPACE - Moves cursor left and erase characters
      IF Ptr THEN
        Edit$ = LEFT$(Edit$, Ptr - 1) + MID$(Edit$, Ptr + 1) + " "
        Ptr = Ptr - 1
      END IF
      
    CASE -83    'DELETE - Erases characters at cursor
      Edit$ = LEFT$(Edit$, Ptr) + MID$(Edit$, Ptr + 2) + " "
      
    CASE ELSE   'ALL OTHER NORMAL KEYS
      IF PressedKey > 1 AND PressedKey < 256 THEN
        K$ = CHR$(PressedKey)
        K$ = UCASE$(K$)
        IF INSTR(GoodKey$, K$) THEN             'if key is in the keylist
          IF InsertMode THEN    'Insert or Overtype the key
            Edit$ = LEFT$(Edit$, Ptr) + K$ + MID$(Edit$, Ptr + 1)
            Edit$ = LEFT$(Edit$, MaxLength)
          ELSE
            IF Ptr < MaxLength THEN MID$(Edit$, Ptr + 1, 1) = K$
          END IF
          IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1             'Are we at max length
        END IF
      END IF
      'End of Editing & Input keys
    END SELECT
  LOOP
  
END SUB

