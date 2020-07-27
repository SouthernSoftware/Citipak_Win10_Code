DEFINT A-Z
DECLARE SUB CheckTransFiles (OkFlag%)
DECLARE SUB ArchiveTransHistory (ExitFlag%, RelinkYear%)
DECLARE SUB CheckEmp2LeaveTblFile ()
DECLARE SUB CheckRetireFile ()
DECLARE SUB CheckUnitFile ()
DECLARE SUB CheckLeaveFile ()
DECLARE SUB CheckStateTaxFile ()
DECLARE SUB RebuildEmpIndexs ()
DECLARE FUNCTION UtilRound# (DblNum#)
DECLARE SUB SumEmpYTD ()
DECLARE SUB RelinkTransHistory (RelinkYear)
DECLARE SUB ParseOld2New ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION QPValI% (Num$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB ReadTransPins ()
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION UtilRound# (DblNum#)
DECLARE SUB NameFile (OldName$, NewName$)
DECLARE SUB KillFile (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'Retire.Bi'
  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'StateTax.Bi'
  'INCLUDE: 'FedTax.Bi'
  'INCLUDE: 'PREIC.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'oPRTRANS.Bi'
  '$INCLUDE: 'PROLDTYP.Bi'
  '$INCLUDE: 'prif.bi'
  
  TYPE EmpNumSortType

    EmpNum AS STRING * 10
    RecNo  AS INTEGER
  END TYPE
  
  
  CONST False = 0, True = NOT False
  CONST OldHistFileName = "prdata\oldtranH.dat"
  
  DIM SHARED Blank3 AS TransRecType
  
  DIM SHARED TRecType1 AS TransRecType1
  DIM SHARED TRecType2 AS TransRecType2
  DIM SHARED TRecType3 AS TransRecType
  
  REDIM SHARED TPins(1 TO 1)
  
  DIM SHARED Emp1Rec   AS EmpData1Type
  DIM SHARED Emp2Rec   AS EmpData2Type
  DIM SHARED Emp3Rec   AS EmpData3Type
  DIM SHARED Emp3RecB  AS EmpData3Type
  DIM SHARED TranRec   AS TransRecType
  
  DIM SHARED TNumOfRec AS INTEGER, TranRecLen AS INTEGER
  
  CLS
  
  IF Exist(PRActiveFile) THEN
    PRINT "*";
    KillFile PRActiveFile
  END IF
  
  PRINT "   PRUTIL1 v2.0  1994-2000 Southern Software."
  PRINT
  
  'ON ERROR GOTO ErrorExit
  
'GOTO Reindex
  CheckTransFiles OkFlag
  
  IF NOT OkFlag THEN
    GOTO ProgramExit
  END IF
  
  'NOTE: Make sure the last routine in this group has the
  '      Smart Drive Flush routine at the end
  CheckRetireFile
  CheckUnitFile
  CheckLeaveFile
  CheckStateTaxFile
  CheckEmp2LeaveTblFile
  '*** end of group
  
  PRINT
  
  Cmd$ = COMMAND$
  
  IF INSTR(Cmd$, "ARCHIVE") THEN
    ArchiveTransHistory ExitFlag, RelinkYear
  END IF
  
  IF INSTR(Cmd$, "RELINK") AND NOT ExitFlag THEN
    RelinkTransHistory RelinkYear
  END IF
  
  IF INSTR(Cmd$, "REINDEX") THEN
Reindex:
    RebuildEmpIndexs
  END IF
  
  'PRINT
  
ProgramExit:
  PRINT
  IF NOT OkFlag THEN
    PRINT "PROCESSING ABORTED!"
  ELSE
    PRINT "Processing complete."
  END IF
  
  END
  
ErrorExit:
  ErrCode = ERR
  OkFlag = False
  PRINT
  PRINT "A SERIOUS ERROR HAS OCCURRED!"
  PRINT
  
  SELECT CASE ErrCode
  CASE 53
    PRINT "PRUTIL1:  Could NOT find the payroll data files."
  CASE ELSE
    PRINT "ERROR CODE:"; ErrCode
  END SELECT
  PRINT
  PRINT "Please call Southern Software for assistance."
  PRINT
  GOTO ProgramExit

SUB ArchiveTransHistory (ExitFlag, RelinkYear)
  
  Cmd$ = COMMAND$
  
  ArchivePos = INSTR(Cmd$, "ARCHIVE")
  ArchivePos = ArchivePos + 7
  
  ReLinkPos = INSTR(Cmd$, "RELINK")
  IF ReLinkPos = 0 THEN
    GOTO ArchiveErrExit
  END IF
  ReLinkPos = ReLinkPos + 6
  
  ArchiveYear$ = MID$(Cmd$, ArchivePos, 3)
  RelinkYear$ = MID$(Cmd$, ReLinkPos, 6)
  
  StartPos = INSTR(ArchiveYear$, "9")
  IF StartPos = 0 THEN
    GOTO ArchiveErrExit
  END IF
  YearArchive = QPValI(ArchiveYear$)
  
  'StartPos = INSTR(RelinkYear$, "9")
  'IF StartPos = 0 THEN
  '  GOTO ArchiveErrExit
  'END IF
  YearReLink = QPValI(RelinkYear$)
  
  IF YearArchive >= YearReLink THEN
    GOTO ArchiveErrExit
  END IF
  
  SELECT CASE YearArchive
  CASE 94 TO 99
    ArchiveOk = True
  CASE ELSE
    ArchiveOk = False
  END SELECT
  
  SELECT CASE YearReLink
  CASE IS >= 2000
    ReLinkOk = True
  CASE ELSE
    ReLinkOk = False
  END SELECT
  
  IF NOT ReLinkOk OR NOT ArchiveOk THEN
    GOTO ArchiveErrExit
  END IF
  
  RelinkYear = YearReLink
  
  ArchiveYear$ = QPStrI$(YearArchive)
  
  ArchiveFile$ = "PRDATA\PRTRANSH.D" + ArchiveYear$
  PRINT
  PRINT "Archive Payroll History for year " + ArchiveYear$
  PRINT "ARE YOU SURE YOU WANT TO DO THIS (YES)? ";
  LINE INPUT ""; YesNo$
  YesNo$ = UCASE$(YesNo$)
  
  IF YesNo$ <> "YES" THEN
    PRINT "Archive Aborted!!!"
    ExitFlag = True
    GOTO ArchiveExit
  END IF
  
  LowDate = Date2Num("01-01-19" + ArchiveYear$)
  HiDate = Date2Num("12-31-19" + ArchiveYear$)
  
  REDIM PTransRec(1)   AS TransRecType
  
  TranRecLen = LEN(PTransRec(1))
  
  TNumOfRec = FileSize(TransHistFileName) \ TranRecLen
  
  NameFile TransHistFileName, OldHistFileName
  
  OPEN OldHistFileName FOR RANDOM AS #1 LEN = TranRecLen
  
  OPEN TransHistFileName FOR RANDOM AS #2 LEN = TranRecLen
  
  OPEN ArchiveFile$ FOR RANDOM AS #3 LEN = TranRecLen
  
  PRINT "Number of transactions archived: 0";
  Row = CSRLIN
  Col = POS(0) - 1
  ArchiveCnt = 0
  
  FOR Cnt = 1 TO TNumOfRec
    GET #1, Cnt, PTransRec(1)
    IF PTransRec(1).CheckDate < LowDate OR PTransRec(1).CheckDate > HiDate THEN
      PUT #2, , PTransRec(1)
    ELSE
      ArchiveCnt = ArchiveCnt + 1
      LOCATE Row, Col
      PRINT ArchiveCnt;
      PUT #3, , PTransRec(1)
    END IF
  NEXT
  CLOSE
  KILL OldHistFileName
  
  GOTO ArchiveExit
  
ArchiveErrExit:
  BEEP
  PRINT
  PRINT
  PRINT "Invalid year spec for Archive or relink!"
  
ArchiveExit:
  
END SUB

SUB CheckEmp2LeaveTblFile
  
  'Emp3Rec   AS EmpData3Type
  REDIM NEmp2Rec(1)  AS EmpData2Type
  
  Emp2Len = LEN(NEmp2Rec(1))
  
  NumOfRec = FileSize(EmpData2Name) \ Emp2Len
  
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2Len
  FOR Cnt = 1 TO NumOfRec
    GET #1, Cnt, NEmp2Rec(1)
    IF NEmp2Rec(1).LeaveTbl < 1 THEN
      NEmp2Rec(1).LeaveTbl = 1
      DidItFlag = True
      PUT #1, Cnt, NEmp2Rec(1)
    END IF
  NEXT
  CLOSE #1
  IF DidItFlag THEN
    PRINT "Converted Employee leave table entries.";
  ELSE
    PRINT "Employee leave table entries are valid.";
  END IF
  ERASE NEmp2Rec
  
END SUB

SUB CheckLeaveFile
  
  
  REDIM OldLeaveRec(1) AS oLeaveRecType
  REDIM NewLeaveRec(1) AS LeaveRecType
  
  OLen = LEN(OldLeaveRec(1))
  NLen = LEN(NewLeaveRec(1))
  
  IF FileSize(LeaveFileName) = OLen THEN
    PRINT "Converting LEAVE file structure."
    NAME LeaveFileName AS "prdata\oleave.dat"
    OPEN "prdata\oleave.dat" FOR RANDOM AS #1 LEN = OLen
    OPEN LeaveFileName FOR RANDOM AS #2 LEN = NLen
    GET #1, , OldLeaveRec(1)
    NewLeaveRec(1).VACMAX = OldLeaveRec(1).VACMAX
    NewLeaveRec(1).SICKMAX = OldLeaveRec(1).SICKMAX
    FOR Cnt = 1 TO 6
      NewLeaveRec(1).VEntry(Cnt).YEARS = OldLeaveRec(1).VEntry(Cnt).YEARS1
      NewLeaveRec(1).VEntry(Cnt).EARN = OldLeaveRec(1).VEntry(Cnt).EARN
      NewLeaveRec(1).SEntry(Cnt).YEARS = OldLeaveRec(1).SEntry(Cnt).YEARS1
      NewLeaveRec(1).SEntry(Cnt).EARN = OldLeaveRec(1).SEntry(Cnt).EARN
    NEXT
    PUT #2, , NewLeaveRec(1)
    CLOSE
    KILL "prdata\oLeave.dat"
  ELSE
    PRINT "Current LEAVE file structure is correct."
  END IF
  ERASE OldLeaveRec, NewLeaveRec
  
END SUB

SUB CheckRetireFile
  
  REDIM OldRetRec(1) AS oRetireRecType
  REDIM NewRetRec(1) AS RetireRecType
  
  OLen = LEN(OldRetRec(1))
  NLen = LEN(NewRetRec(1))
  
  IF FileSize(RetireFileName) = OLen * 6 THEN
    PRINT "Converting RETIREMENT file structure."
    NAME RetireFileName AS "prdata\oretire.dat"
    
    OPEN "prdata\oretire.dat" FOR RANDOM AS #1 LEN = OLen
    OPEN RetireFileName FOR RANDOM AS #2 LEN = NLen
    FOR Cnt = 1 TO 6
      GET #1, , OldRetRec(1)
      LSET NewRetRec(1) = OldRetRec(1)
      NewRetRec(1).TYPETD1 = "Y"
      PUT #2, , NewRetRec(1)
    NEXT
    CLOSE
    KILL "prdata\oretire.dat"
  ELSE
    PRINT "Current retirement file structure is correct."
  END IF
  ERASE OldRetRec, NewRetRec
  
END SUB

SUB CheckStateTaxFile
  
'  REDIM OldStateRec(1) AS oStateTaxRecType
'  REDIM NewStateRec(1) AS StateTaxRecType
'
'  OLen = LEN(OldStateRec(1))
'  NLen = LEN(NewStateRec(1))
'
'  IF FileSize(StateTaxFileName) = OLen THEN
'    PRINT "Converting State Tax file structure."
'    NAME StateTaxFileName AS "prdata\oSta.dat"
'    OPEN "prdata\oSta.dat" FOR RANDOM AS #1 LEN = OLen
'    OPEN StateTaxFileName FOR RANDOM AS #2 LEN = NLen
'    GET #1, , OldStateRec(1)
'    '-----------------
'    NewStateRec(1).TAX101 = OldStateRec(1).STSGWCA
'    NewStateRec(1).TAX102 = OldStateRec(1).STSMINSA
'    NewStateRec(1).TAX103 = OldStateRec(1).STSMAXSA
'    NewStateRec(1).TAX104 = OldStateRec(1).STSSUI1
'    NewStateRec(1).TAX105 = OldStateRec(1).STSSUI2
'    NewStateRec(1).TAX106 = OldStateRec(1).STSSDAA
'
'    NewStateRec(1).TAX201 = OldStateRec(1).STMGWCA
'    NewStateRec(1).TAX202 = OldStateRec(1).STMMINSA
'    NewStateRec(1).TAX203 = OldStateRec(1).STMMAXSA
'    NewStateRec(1).TAX204 = OldStateRec(1).STMSUI1
'    NewStateRec(1).TAX205 = OldStateRec(1).STMSUI2
'    NewStateRec(1).TAX206 = OldStateRec(1).STMSDAA
'
'    NewStateRec(1).TAX301 = OldStateRec(1).STHGWCA
'    NewStateRec(1).TAX302 = OldStateRec(1).STHMINSA
'    NewStateRec(1).TAX303 = OldStateRec(1).STHMAXSA
'    NewStateRec(1).TAX304 = OldStateRec(1).STHSUI1
'    NewStateRec(1).TAX305 = OldStateRec(1).STHSUI2
'    NewStateRec(1).TAX306 = OldStateRec(1).STHSDAA
'
'    FOR Cnt1 = 1 TO 3
'      FOR Cnt2 = 1 TO 10
'        NewStateRec(1).STS(Cnt1, Cnt2) = OldStateRec(1).STS(Cnt1, Cnt2)
'        NewStateRec(1).STM(Cnt1, Cnt2) = OldStateRec(1).STM(Cnt1, Cnt2)
'        NewStateRec(1).STH(Cnt1, Cnt2) = OldStateRec(1).STH(Cnt1, Cnt2)
'      NEXT
'    NEXT
'
'    '-----------------
'    PUT #2, , NewStateRec(1)
'    CLOSE
'    KILL "prdata\oSta.dat"
'  ELSE
'    PRINT "Current State Tax file structure is correct."
'  END IF
'
'  ERASE OldStateRec, NewStateRec
  
  
END SUB

SUB CheckTransFiles (OkFlag%)
  
  TLenType1 = LEN(TRecType1)
  TLenType2 = LEN(TRecType2)
  TLenType3 = LEN(TRecType3)
  
  NumOfEmp = FileSize(EmpIdxLName) \ 2
  
  IF NumOfEmp = 0 THEN
    ERROR 53
  END IF
  
  CurRecLen# = FileSize(TransWorkFileName) / NumOfEmp
  'this is double precision to help detect damaged file errors.
  'i.e. if there is ANY REMAINDER then the file is broken.
  
  SELECT CASE CurRecLen#
  CASE TLenType1
    OldRecLen = TLenType1
    Type1Flag = True
    OkFlag = False
    PRINT "Converting from TYPE1 record structure."
  CASE TLenType2
    OldRecLen = TLenType2
    OkFlag = False
    PRINT "Converting from TYPE2 record structure."
  CASE TLenType3
    OkFlag = True
    PRINT "Current transaction structure is correct."
  CASE ELSE
    DamagedFlag = True
    OkFlag = False
    PRINT "INVALID TRANSACTION RECORD LENGTH."
  END SELECT
  
  IF OkFlag THEN                'if the trans structure is ok then skip
    'OkFlag = True
    EXIT SUB
  END IF
  
  IF DamagedFlag THEN
    PRINT "YOUR DATA FILES ARE PROBABLY DAMAGED!!"
    PRINT "UNABLE TO CONTINUE"
    OkFlag = False
    EXIT SUB
  END IF
  
  NumOfRecs = FileSize(TransHistFileName) \ OldRecLen
  
  IF NumOfRecs = 0 THEN GOTO SkipHist
  
  NAME TransHistFileName AS "TEMPHIST.DAT"
  
  OPEN "TEMPHIST.DAT" FOR RANDOM AS #1 LEN = OldRecLen
  OPEN TransHistFileName FOR RANDOM AS #2 LEN = TLenType3
  
  FOR Cnt = 1 TO NumOfRecs
    IF Type1Flag THEN
      GET #1, Cnt, TRecType1
      ParseOld2New
    ELSE
      GET #1, Cnt, TRecType2
    END IF
    TRecType3 = Blank3
    LSET TRecType3 = TRecType2
    TRecType3.RetGrossPay = TRecType3.GrossPay
    PUT #2, Cnt, TRecType3
  NEXT
  CLOSE
  
  KILL "TEMPHIST.DAT"
  
SkipHist:
  
  NumOfRecs = FileSize(TransWorkFileName) \ OldRecLen
  
  NAME TransWorkFileName AS "TEMPWORK.DAT"
  
  OPEN "TEMPWORK.DAT" FOR RANDOM AS #1 LEN = OldRecLen
  OPEN TransWorkFileName FOR RANDOM AS #2 LEN = TLenType3
  
  FOR Cnt = 1 TO NumOfRecs
    IF Type1Flag THEN
      GET #1, Cnt, TRecType1
      ParseOld2New
    ELSE
      GET #1, Cnt, TRecType2
    END IF
    LSET TRecType3 = TRecType2
    TRecType3.RetGrossPay = TRecType3.GrossPay
    PUT #2, Cnt, TRecType3
  NEXT
  
  
  CLOSE
  
  KILL "TEMPWORK.DAT"
  
END SUB

SUB CheckUnitFile
  
  REDIM OldUnitRec(1) AS oUnitFileRecType
  REDIM NewUnitRec(1) AS UnitFileRecType
  
  OLen = LEN(OldUnitRec(1))
  NLen = LEN(NewUnitRec(1))
  
  IF FileSize(UnitFileName) = OLen THEN
    PRINT "Converting UNIT file structure."
    NAME UnitFileName AS "prdata\ounit.dat"
    OPEN "prdata\ounit.dat" FOR RANDOM AS #1 LEN = OLen
    OPEN UnitFileName FOR RANDOM AS #2 LEN = NLen
    GET #1, , OldUnitRec(1)
    LSET NewUnitRec(1) = OldUnitRec(1)
    NewUnitRec(1).ESCRTYPE = 1
    NewUnitRec(1).TAXWBASE = 0
    PUT #2, , NewUnitRec(1)
    CLOSE
    KILL "prdata\ounit.dat"
  ELSE
    PRINT "Current UNIT file structure is correct."
  END IF
  ERASE OldUnitRec, NewUnitRec
  
END SUB

SUB HoldThisStuff
  
  'StartHere:
  '    RebuildEmpIndexs
  
  '   '  NDate = Date2Num("01-20-1995")
  '   '  get date
  '     CDate = Date2Num("05-10-1995")
  '     DIM TR AS TransRecType
  '     RecLen = LEN(TR)
  '
  
  
  '     OPEN TransHistFileName FOR RANDOM AS #1 LEN = RecLen
  '     OPEN "prdata\newtranh.dat" FOR RANDOM AS #2 LEN = RecLen
  '
  '   ''  OPEN "trandate.lst" FOR OUTPUT AS #2
  '   ''
  '     NumOfRec = FileSize(TransHistFileName) \ RecLen
  '     FOR Cnt = 1 TO NumOfRec
  '       GET #1, Cnt, TR
  '       IF TR.CheckDate = CDate THEN  'AND TR.GrossPay = 0 THEN
  '         'TR.StaGrossPay = TR.EAmt(1)
  '         'TR.FedGrossPay = TR.EAmt(1)
  '
  '         RCnt = RCnt + 1
  '       ELSE
  '         PUT #2, , TR
  '
  '   '      TR.RegHrsWork = 1
  '   '      '? TR.
  '       END IF 'ELSE
  '   '    PUT #2, , TR
  '   '    'END IF
  '   ''      TDate = TR.PostDate   'PayPdEnd
  '   ''      RecNum = Cnt
  '   ''      Pin = TR.EmpPin
  '   ''      TR.TActive = True
  '   ''      TR.CheckDate = Date2Num("07-07-1994")
  '   ''      PUT #1, 428, TR
  '   ''    END IF
  '   ''     PRINT #2, Num2Date(TR.PostDate)
  '   '
  '     NEXT
  '     CLOSE
  '     PRINT "Match Count:"; RCnt
  '   ''  PRINT "Check Date:"; Num2Date(TR.CheckDate), "Rec Num:"; RecNum, Pin
  '   END
  
  
  
  'StartHere2:
  '
  '  CDate = Date2Num("05-12-1995")
  '
  '  IdxRecLen = 2
  '  IdxFileSize& = FileSize(EmpIdxLName)
  '  NumOfRecs = IdxFileSize& \ IdxRecLen
  '
  '  REDIM IdxBuff(1 TO NumOfRecs)
  '  REDIM TransHRec(1)     AS TransRecType
  '  REDIM Emp2DRec(1)  AS EmpData2Type
  '
  '  EmpRecSize = LEN(Emp2DRec(1))
  '  TRecSize = LEN(TransHRec(1))
  '
  '  FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, NumOfRecs
  '
  '  OPEN EmpData2Name FOR RANDOM AS #1 LEN = EmpRecSize 'open employee data file
  '
  '  OPEN TransHistFileName FOR RANDOM AS #2 LEN = TRecSize
  '
  '  OPEN "prdata\newtranh.dat" FOR RANDOM AS #3 LEN = TRecSize
  '
  '  FOR RecNo = 1 TO NumOfRecs
  '    GotIt = False
  '
  '    GET #1, IdxBuff(RecNo), Emp2DRec(1)
  '
  '    IF Emp2DRec(1).LastTransRec <= 0 THEN
  '      GOTO SkipEm2
  '    END IF
  '
  '    TransRecNum& = CLNG(Emp2DRec(1).LastTransRec)
  '
  '    DO
  '      GET #2, TransRecNum&, TransHRec(1)
  '
  '      IF (TransHRec(1).CheckDate = CDate) AND NOT GotIt THEN
  '        GotIt = True
  '      ELSE
  '
  '        PUT #3, , TransHRec(1)
  '      END IF
  '
  '      IF TransHRec(1).PrevTransRec > 0 THEN
  '        TransRecNum& = CLNG(TransHRec(1).PrevTransRec)
  '      ELSE
  '        EXIT DO
  '      END IF
  '    LOOP
  'SkipEm2:
  '   ' ShowPctComp RecNo, NumOfRecs
  '  NEXT
  '
  '  CLOSE
  '  END
  
END SUB

SUB ParseOld2New
  TRecType2.TActive = TRecType1.TActive
  TRecType2.PrevTransRec = TRecType1.PrevTransRec
  TRecType2.EmpPin = TRecType1.EmpPin
  
  TRecType2.PaySFlag = TRecType1.PaySFlag
  
  TRecType2.CheckNum = TRecType1.CheckNum
  TRecType2.PayPdStart = TRecType1.PayPdStart
  TRecType2.PayPdEnd = TRecType1.PayPdEnd
  TRecType2.CheckDate = TRecType1.CheckDate
  TRecType2.PostDate = TRecType1.PostDate
  
  TRecType2.PayType = TRecType1.PayType
  TRecType2.BaseRate = TRecType1.BaseRate
  TRecType2.OTRate = TRecType1.OTRate
  
  TRecType2.RegHrsWork = TRecType1.RegHrsWork
  TRecType2.VacUsed = TRecType1.VacUsed
  TRecType2.SickUsed = TRecType1.SickUsed
  TRecType2.CompUsed = TRecType1.CompUsed
  TRecType2.HOLHOURS = TRecType1.HOLHOURS
  TRecType2.RegHrsPaid = TRecType1.RegHrsPaid
  
  TRecType2.OTHours = TRecType1.OTHours
  TRecType2.OTHrsPaid = TRecType1.OTHrsPaid
  TRecType2.OT2Comp = TRecType1.OT2Comp
  
  FOR Cnt = 1 TO 8
    TRecType2.TDist(Cnt) = TRecType1.TDist(Cnt)
  NEXT
  
  TRecType2.TotRegWage = TRecType1.TotRegWage
  TRecType2.TotOTWage = TRecType1.TotOTWage
  
  TRecType2.GrossWage = TRecType1.GrossWage
  
  FOR Cnt = 1 TO 3
    TRecType2.EAmt(Cnt) = TRecType1.EAmt(Cnt)
  NEXT
  FOR Cnt = 1 TO 6
    TRecType2.EDist(Cnt) = TRecType1.EDist(Cnt)
  NEXT
  TRecType2.TotAdditEarn = TRecType1.TotAdditEarn
  
  TRecType2.GrossPay = TRecType1.GrossPay
  
  TRecType2.SocGrossPay = TRecType1.SocGrossPay
  TRecType2.MedGrossPay = TRecType1.MedGrossPay
  TRecType2.FedGrossPay = TRecType1.FedGrossPay
  TRecType2.StaGrossPay = TRecType1.StaGrossPay
  
  TRecType2.SocTaxAmt = TRecType1.SocTaxAmt
  TRecType2.MedTaxAmt = TRecType1.MedTaxAmt
  TRecType2.FedTaxAmt = TRecType1.FedTaxAmt
  TRecType2.StaTaxAmt = TRecType1.StaTaxAmt
  TRecType2.TotTaxAmt = TRecType1.TotTaxAmt
  TRecType2.RetireAmt = TRecType1.RetireAmt
  
  FOR Cnt = 1 TO 12
    TRecType2.DAmt(Cnt) = TRecType1.DAmt(Cnt)
  NEXT
  
  TRecType2.TotDedAmt = TRecType1.TotDedAmt
  
  TRecType2.EICAmt = 0
  
  TRecType2.NetPay = TRecType1.NetPay
  
  TRecType2.PeriodHistRec = TRecType1.PeriodHistRec
  
  TRecType2.MatchRetAmt = TRecType1.MatchRetAmt
  TRecType2.MatchSocAmt = TRecType1.MatchSocAmt
  TRecType2.MatchMedAmt = TRecType1.MatchMedAmt
  
  
END SUB

SUB ReadTransPins
  
  '  TNumOfRec = FileSize(TransHistFileName) \ TranRecLen
  PRINT
  PRINT "Reading Transaction history PIN numbers: ";
  Col = POS(0)
  REDIM TPins(1 TO TNumOfRec)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TranRecLen
  FOR Cnt = 1 TO TNumOfRec
    LOCATE , Col: PRINT Cnt;
    GET #1, , TranRec
    TPins(Cnt) = TranRec.EmpPin
  NEXT
  CLOSE
END SUB

SUB RebuildEmpIndexs
  
  PRINT "Reindexing employee data.";
  
  Emp1Len = LEN(Emp1Rec)
  NumOfRec = FileSize(EmpData1Name) \ Emp1Len
  
  REDIM Emp1Data(1 TO NumOfRec) AS EmpData1Type
  
  FGetAH EmpData1Name, SEG Emp1Data(1), Emp1Len, NumOfRec
  
  SortT Emp1Data(1), NumOfRec, 0, Emp1Len, 0, 10                'emp number
  
  OPEN EmpIdxNName FOR RANDOM AS #1 LEN = 2
  FOR Cnt = 1 TO NumOfRec
    PUT #1, Cnt, Emp1Data(Cnt).Data1RecNum
  NEXT
  CLOSE
  
  SortT Emp1Data(1), NumOfRec, 0, Emp1Len, 10, 24               'last name
  
  'this sorts the first names for each last name
  First = 1
  DO
    FOR Cnt = First TO NumOfRec
      IF Cnt = NumOfRec THEN EXIT FOR
      IF INSTR(Emp1Data(Cnt).EMPLNAME, Emp1Data(Cnt + 1).EMPLNAME) THEN
        Last = Cnt
        DO
          Last = Last + 1
        LOOP WHILE INSTR(Emp1Data(Cnt).EMPLNAME, Emp1Data(Last).EMPLNAME)
        SortT Emp1Data(Cnt), Last - Cnt, 0, Emp1Len, 34, 24     'first name
        First = Last
        EXIT FOR
      END IF
    NEXT
  LOOP UNTIL Cnt >= NumOfRec
  
  
  OPEN EmpIdxLName FOR RANDOM AS #1 LEN = 2
  FOR Cnt = 1 TO NumOfRec
    PUT #1, Cnt, Emp1Data(Cnt).Data1RecNum
  NEXT
  CLOSE
  LOCATE , 1
  PRINT "Reindexing Complete.     "
  
  'OPEN EmpData3Name FOR RANDOM AS #4 LEN = Emp3RecLen
  
END SUB

SUB RelinkTransHistory (YearReLink)
  
'  IF YearReLink > 0 THEN
'    RelinkYear$ = QPStrI(YearReLink)
'  ELSE
'    RelinkYear$ = RIGHT$(COMMAND$, 2)
'  END IF
'
'  IF NOT INSTR("949596979899", RelinkYear$) > 0 THEN
'    BEEP
'    PRINT
'    PRINT
'    PRINT "Invalid year spec for relink!"
'    EXIT SUB
'  END IF
'
'  LowDate = Date2Num("01-01-19" + RelinkYear$)
'  HiDate = Date2Num("12-31-19" + RelinkYear$)

  IF YearReLink > 0 THEN
    RelinkYear$ = QPStrI(YearReLink)
  ELSE
    RelinkYear$ = RIGHT$(COMMAND$, 2)
  END IF

  IF NOT INSTR("949596979899000102030405", RelinkYear$) > 0 THEN
    BEEP
    PRINT
    PRINT
    PRINT "Invalid year spec for relink!"
    EXIT SUB
  END IF

  Year = QPValI(RIGHT$(DATE$, 4))

  SELECT CASE Year
  CASE IS < 2000
    LowDate = Date2Num("01-01-19" + RelinkYear$)
    HiDate = Date2Num("12-31-19" + RelinkYear$)
  CASE ELSE 'it greater or equal
    LowDate = Date2Num("01-01-20" + RelinkYear$)
    HiDate = Date2Num("12-31-20" + RelinkYear$)
  END SELECT


  


  REDIM TPntr(0 TO 200)
  
  REDIM TPins(1 TO 1)
  
  DIM ToScrn AS STRING * 60
  
  Emp2RecLen = LEN(Emp2Rec)
  Emp3RecLen = LEN(Emp3Rec)
  
  TranRecLen = LEN(TranRec)
  
  ENumOfRec = FileSize(EmpData2Name) \ Emp2RecLen
  TNumOfRec = FileSize(TransHistFileName) \ TranRecLen
  
  '  REDIM IdxBuff(1 TO ENumOfRec)
  '  FGetAH EmpIdxLName, IdxBuff(1), 2, ENumOfRec
  
  NewRecCnt = 1
  
  '  COLOR 15, 1
  '  CLS
  ReadTransPins
  
  NAME TransHistFileName AS OldHistFileName
  
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2RecLen
  OPEN EmpData3Name FOR RANDOM AS #4 LEN = Emp3RecLen
  OPEN OldHistFileName FOR RANDOM AS #2 LEN = TranRecLen
  OPEN TransHistFileName FOR RANDOM AS #3 LEN = TranRecLen
  
  '  OPEN EmpData3Name FOR RANDOM AS #3 LEN = Emp3RecLen
  Row = CSRLIN
  FOR ECnt = 1 TO ENumOfRec
    GET #1, ECnt, Emp2Rec
    LOCATE Row, 1
    LSET ToScrn = "Relinking Transactions: " + FUsing(STR$((ECnt / ENumOfRec) * 100), "###") + "% Completed."
    PRINT ToScrn;
    GOSUB GetTransRecNums
    IF TPntr(0) THEN
      GOSUB RebuildTransHistory
    ELSE
      Emp2Rec.LastTransRec = 0
    END IF
    PUT #1, ECnt, Emp2Rec
    PUT #4, ECnt, Emp3Rec
    
    Emp3Rec = Emp3RecB
    
  NEXT
  CLOSE
  
  LOCATE Row, 1
  LSET ToScrn = "Transaction Relink Complete."
  PRINT ToScrn;
  KILL OldHistFileName
  
  EXIT SUB
  
  
GetTransRecNums:
  REDIM TPntr(0 TO 600)
  TotalTransRecs = 0
  FOR TCnt = 1 TO TNumOfRec
    IF TPins(TCnt) = Emp2Rec.EmpPin THEN
      TotalTransRecs = TotalTransRecs + 1
      TPntr(TotalTransRecs) = TCnt
    END IF
    TPntr(0) = TotalTransRecs
  NEXT
  RETURN
  
RebuildTransHistory:
  FirstEmpHRec = NewRecCnt
  '   PRINT TPntr(0): STOP
  
  
  FOR Cnt = 1 TO TPntr(0)
    GET #2, TPntr(Cnt), TranRec
    IF Cnt = 1 THEN
      TranRec.PrevTransRec = 0
    ELSE
      TranRec.PrevTransRec = NewRecCnt - 1
    END IF
    TranRec.RetGrossPay = TranRec.GrossPay
    PUT #3, NewRecCnt, TranRec
    NewRecCnt = NewRecCnt + 1
    Emp2Rec.LastTransRec = NewRecCnt - 1
    SELECT CASE TranRec.CheckDate
    CASE LowDate TO HiDate
      SumEmpYTD
    END SELECT
  NEXT
  RETURN
  
  
END SUB

SUB SumEmpYTD
  ''** Update employee 3 file
  ''-=-=man
  Emp3Rec.YTDGrossPay = UtilRound(Emp3Rec.YTDGrossPay + TranRec.GrossPay)
  Emp3Rec.YTDFedGrossPay = UtilRound(Emp3Rec.YTDFedGrossPay + TranRec.FedGrossPay)
  Emp3Rec.YTDStaGrossPay = UtilRound(Emp3Rec.YTDStaGrossPay + TranRec.StaGrossPay)
  Emp3Rec.YTDSocGrossPay = UtilRound(Emp3Rec.YTDSocGrossPay + TranRec.SocGrossPay)
  Emp3Rec.YTDMedGrossPay = UtilRound(Emp3Rec.YTDMedGrossPay + TranRec.MedGrossPay)
  
  Emp3Rec.YTDRegPay = UtilRound(Emp3Rec.YTDRegPay + TranRec.TotRegWage)
  Emp3Rec.YTDOTPay = UtilRound(Emp3Rec.YTDOTPay + TranRec.TotOTWage)
  
  Emp3Rec.YTDNet = UtilRound(Emp3Rec.YTDNet + TranRec.NetPay)
  
  Emp3Rec.YTDFederal = UtilRound(Emp3Rec.YTDFederal + TranRec.FedTaxAmt)
  Emp3Rec.YTDState = UtilRound(Emp3Rec.YTDState + TranRec.StaTaxAmt)
  Emp3Rec.YTDSocial = UtilRound(Emp3Rec.YTDSocial + TranRec.SocTaxAmt)
  Emp3Rec.YTDMedicare = UtilRound(Emp3Rec.YTDMedicare + TranRec.MedTaxAmt)
  Emp3Rec.YTDRetire = UtilRound(Emp3Rec.YTDRetire + TranRec.RetireAmt)
  
  'year to date totals on deductions
  FOR Cnt = 1 TO 12
    Emp3Rec.YTDDAmt(Cnt) = UtilRound(Emp3Rec.YTDDAmt(Cnt) + TranRec.DAmt(Cnt))
    Emp3Rec.YTDDAmtT = UtilRound(Emp3Rec.YTDDAmtT + TranRec.DAmt(Cnt))
  NEXT
  
  'year to date totals on alt earnings
  Emp3Rec.YTDEarn1 = UtilRound(Emp3Rec.YTDEarn1 + TranRec.EAmt(1))
  Emp3Rec.YTDEarn2 = UtilRound(Emp3Rec.YTDEarn2 + TranRec.EAmt(2))
  Emp3Rec.YTDEarn3 = UtilRound(Emp3Rec.YTDEarn3 + TranRec.EAmt(3))
  Emp3Rec.YTDEarnT = UtilRound(Emp3Rec.YTDEarn1 + Emp3Rec.YTDEarn2 + Emp3Rec.YTDEarn3)
  
  '    EmpRec2(1).EMPVACE = UtilRound(EmpRec2(1).EMPVBAL + EmpRec2(1).EMPVUSED)
  '    EmpRec2(1).EMPSLE = UtilRound(EmpRec2(1).EMPSLBAL + EmpRec2(1).EMPSLUSE)
  '    EmpRec2(1).EMPCTE = UtilRound(EmpRec2(1).EMPCTBAL + EmpRec2(1).EMPCTUSE)
END SUB

FUNCTION UtilRound# (DblNum#)
  UtilRound# = (INT((DblNum# * 100) + .5) / 100)
END FUNCTION

