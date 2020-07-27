DEFINT A-Z
DECLARE SUB CheckTransFile ()
DECLARE SUB ReadTransPins ()
DECLARE SUB SumEmpYTD ()
DECLARE SUB CheckEmp2File ()
DECLARE SUB CheckEmpFiles (OKFlag%)
DECLARE SUB CheckTransFiles (OKFlag%)
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
DECLARE FUNCTION Exist (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FOpenAll (FileName$, AccessMode%, SharedMode%, Handle%)

DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB ReadTransPins ()
DECLARE FUNCTION Date2Num (DateIn$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION UtilRound# (DblNum#)
DECLARE SUB NameFile (OldName$, NewName$)
DECLARE SUB KillFile (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'Retire.Bi'
  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRPPDEF.Bi'
  '$INCLUDE: 'PRTRANS.Bi'
  '$INCLUDE: 'OPRTYPes.Bi'
  ''$INCLUDE: 'PROLDTYP.Bi'
  '$INCLUDE: 'prif.bi'

TYPE UnitFileRecType1
    UFEMPR   AS STRING * 35
    UFATTN   AS STRING * 35
    UFADDR1  AS STRING * 35
    UFADDR2  AS STRING * 35
    UFCITY   AS STRING * 25
    UFSTATE  AS STRING * 2
    UFZIP    AS STRING * 10
    UFFEDID  AS STRING * 14
    UFSTAID  AS STRING * 14
    UFRETID  AS STRING * 14
    UFRETIDL AS STRING * 14
    ESCRTYPE AS INTEGER
    TAXWBASE AS DOUBLE
    BBTCNTNO AS STRING * 12
    BBTBATCH AS STRING * 12
END TYPE

  
  TYPE EmpNumSortType
    EmpNum AS STRING * 10
    RecNo  AS INTEGER
  END TYPE
  
  CONST False = 0, True = NOT False
  CONST OldHistFileName = "prdata\oldtranH.dat"
  
  DIM SHARED Emp1Rec   AS EmpData1Type
  DIM SHARED Emp2Rec    AS EmpData2Type
  'DIM SHARED Emp2RecO   AS oEmpData2Type
  DIM SHARED Emp3Rec   AS EmpData3Type
  DIM SHARED Emp3RecB  AS EmpData3Type
  DIM SHARED TranRec   AS TransRecType
  
  DIM SHARED TNumOfRec AS INTEGER, TranRecLen AS INTEGER, OKFlag AS INTEGER
  
  CLS
  IF Exist("PRUTIL1.EXE") THEN
    PRINT "!";
    KILL "PRUTIL1.EXE"
  END IF

  IF Exist(PRActiveFile) THEN
    PRINT "*";
    KillFile PRActiveFile
  END IF
  
  LOCATE , 4
  PRINT "PRUTIL v9.01  1994-2000 Southern Software."
  PRINT

  Cmd$ = COMMAND$
  
  IF INSTR(Cmd$, "HELP") THEN
    PRINT "   No help Available."
    PRINT
    PRINT "   Please contact Southern Software if you need Assistance. "
    GOTO HelpExit
  END IF

'  ON ERROR GOTO ErrorExit

'  OkFlag = True
  
  CheckEmp2File
  CheckUnitFile
  CheckTransFile

  '*** end of group
  
  PRINT
  
  IF INSTR(Cmd$, "ARCHIVE") THEN
    ArchiveTransHistory ExitFlag, RelinkYear
  END IF
  
  IF INSTR(Cmd$, "RELINK") AND NOT ExitFlag THEN
    RelinkTransHistory RelinkYear
  END IF
  
  IF INSTR(Cmd$, "REINDEX") THEN
    RebuildEmpIndexs
  END IF
  
ProgramExit:
  PRINT
  IF NOT OKFlag THEN
    PRINT "PROCESSING ABORTED!"
  ELSE
    PRINT "Processing complete."
  END IF

HelpExit:
  END
  
ErrorExit:
  ErrCode = ERR
  OKFlag = False
  PRINT
  PRINT "A SERIOUS ERROR HAS OCCURRED!"
  PRINT
  
  SELECT CASE ErrCode
  CASE 53
    PRINT "PRUTIL:  Could NOT find the payroll data files."
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
  RelinkYear$ = MID$(Cmd$, ReLinkPos, 3)
  
  StartPos = INSTR(ArchiveYear$, "9")
  IF StartPos = 0 THEN
    GOTO ArchiveErrExit
  END IF
  YearArchive = QPValI(MID$(ArchiveYear$, StartPos))
  
  StartPos = INSTR(RelinkYear$, "9")
  IF StartPos = 0 THEN
    GOTO ArchiveErrExit
  END IF
  YearReLink = QPValI(MID$(RelinkYear$, StartPos))
  
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
  CASE 94 TO 99, 1996 TO 2000
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

SUB CheckEmp2File
  
  REDIM OldEmp2Rec(1) AS oEmpData2Type1
  REDIM NewEmp2Rec(1) AS EmpData2Type

  OLen = LEN(OldEmp2Rec(1))
  NLen = LEN(NewEmp2Rec(1))

  IF FileSize(EmpData2Name) MOD NLen <> 0 THEN
    PRINT "Converting Employee file."

    IF Exist%("prdata\oemp2.dat") THEN
      KillFile "prdata\oemp2.dat"
    END IF

    NAME EmpData2Name AS "prdata\oemp2.dat"

    OPEN "prdata\oemp2.dat" FOR RANDOM AS #1 LEN = OLen
    OPEN EmpData2Name FOR RANDOM AS #2 LEN = NLen
    NumOfRec = LOF(1) \ OLen

    FOR Cnt = 1 TO NumOfRec
      REDIM NewEmp2Rec(1) AS EmpData2Type
      GET #1, , OldEmp2Rec(1)

      NewEmp2Rec(1).EMPNO = OldEmp2Rec(1).EMPNO
      NewEmp2Rec(1).EMPSSN = OldEmp2Rec(1).EMPSSN
      NewEmp2Rec(1).EMPLNAME = OldEmp2Rec(1).EMPLNAME
      NewEmp2Rec(1).EMPFNAME = OldEmp2Rec(1).EMPFNAME
      NewEmp2Rec(1).EMPADDR1 = OldEmp2Rec(1).EMPADDR1
      NewEmp2Rec(1).EMPADDR2 = OldEmp2Rec(1).EMPADDR2
      NewEmp2Rec(1).EMPCITY = OldEmp2Rec(1).EMPCITY
      NewEmp2Rec(1).EMPSTATE = OldEmp2Rec(1).EMPSTATE
      NewEmp2Rec(1).EMPZIP = OldEmp2Rec(1).EMPZIP
      NewEmp2Rec(1).EMPBDAY = OldEmp2Rec(1).EMPBDAY
      NewEmp2Rec(1).EMPGENDR = OldEmp2Rec(1).EMPGENDR
      NewEmp2Rec(1).EMPRACE = OldEmp2Rec(1).EMPRACE
      NewEmp2Rec(1).EMPRETNO = OldEmp2Rec(1).EMPRETNO
      NewEmp2Rec(1).EMPRETTP = OldEmp2Rec(1).EMPRETTP
      NewEmp2Rec(1).DRAFTCOD = ""
      NewEmp2Rec(1).EMPDDACC = ""
      NewEmp2Rec(1).PRENOTED = "N"
      NewEmp2Rec(1).BANKNAME = ""
      NewEmp2Rec(1).BANKLOC = ""
      NewEmp2Rec(1).TRANSIT = ""
      NewEmp2Rec(1).EMPJOB = OldEmp2Rec(1).EMPJOB
      NewEmp2Rec(1).EMPWCCLS = OldEmp2Rec(1).EMPWCCLS
      NewEmp2Rec(1).EMPSTATS = OldEmp2Rec(1).EMPSTATS
      NewEmp2Rec(1).EMPBCODE = OldEmp2Rec(1).EMPBCODE
      NewEmp2Rec(1).EMPPTYPE = OldEmp2Rec(1).EMPPTYPE
      NewEmp2Rec(1).EMPPFREQ = OldEmp2Rec(1).EMPPFREQ
      NewEmp2Rec(1).EMPPRATE = OldEmp2Rec(1).EMPPRATE
      NewEmp2Rec(1).EMPORATE = OldEmp2Rec(1).EMPORATE
      NewEmp2Rec(1).EMPHDATE = OldEmp2Rec(1).EMPHDATE
      NewEmp2Rec(1).EMPRDATE = OldEmp2Rec(1).EMPRDATE
      NewEmp2Rec(1).EMPTDATE = OldEmp2Rec(1).EMPTDATE
      NewEmp2Rec(1).EMPFEDX = OldEmp2Rec(1).EMPFEDX
      NewEmp2Rec(1).EMPFEDO2 = OldEmp2Rec(1).EMPFEDO2
      NewEmp2Rec(1).EMPFEDO1 = OldEmp2Rec(1).EMPFEDO1
      NewEmp2Rec(1).EMPFEDS = OldEmp2Rec(1).EMPFEDS
      NewEmp2Rec(1).EMPFEDA = OldEmp2Rec(1).EMPFEDA
      NewEmp2Rec(1).EMPFEDAA = OldEmp2Rec(1).EMPFEDAA
      NewEmp2Rec(1).EMPSTAX = OldEmp2Rec(1).EMPSTAX
      NewEmp2Rec(1).EMPSTAO2 = OldEmp2Rec(1).EMPSTAO2
      NewEmp2Rec(1).EMPSTAO1 = OldEmp2Rec(1).EMPSTAO1
      NewEmp2Rec(1).EMPSTAS = OldEmp2Rec(1).EMPSTAS
      NewEmp2Rec(1).EMPSTAA = OldEmp2Rec(1).EMPSTAA
      NewEmp2Rec(1).EMPSTAAA = OldEmp2Rec(1).EMPSTAAA
      NewEmp2Rec(1).EMPSOCX = OldEmp2Rec(1).EMPSOCX
      NewEmp2Rec(1).EMPMEDX = OldEmp2Rec(1).EMPMEDX
      NewEmp2Rec(1).EMPEIC = OldEmp2Rec(1).EMPEIC
      FOR zz = 1 TO 12
        NewEmp2Rec(1).EmpDed(zz) = OldEmp2Rec(1).EmpDed(zz)
      NEXT
      NewEmp2Rec(1).EMPEACT1 = OldEmp2Rec(1).EMPEACT1
      NewEmp2Rec(1).EMPEAMT1 = OldEmp2Rec(1).EMPEAMT1
      NewEmp2Rec(1).EMPEACT2 = OldEmp2Rec(1).EMPEACT2
      NewEmp2Rec(1).EMPEAMT2 = OldEmp2Rec(1).EMPEAMT2
      NewEmp2Rec(1).EMPEACT3 = OldEmp2Rec(1).EMPEACT3
      NewEmp2Rec(1).EMPEAMT3 = OldEmp2Rec(1).EMPEAMT3
      NewEmp2Rec(1).EMPHP = OldEmp2Rec(1).EMPHP
      FOR zz = 1 TO 8
        NewEmp2Rec(1).EDist(zz) = OldEmp2Rec(1).EDist(zz)
      NEXT
      NewEmp2Rec(1).EMPVACE = OldEmp2Rec(1).EMPVACE
      NewEmp2Rec(1).EMPVUSED = OldEmp2Rec(1).EMPVUSED
      NewEmp2Rec(1).EMPVBAL = OldEmp2Rec(1).EMPVBAL
      NewEmp2Rec(1).EMPSLE = OldEmp2Rec(1).EMPSLE
      NewEmp2Rec(1).EMPSLUSE = OldEmp2Rec(1).EMPSLUSE
      NewEmp2Rec(1).EMPSLBAL = OldEmp2Rec(1).EMPSLBAL
      NewEmp2Rec(1).EMPCTE = OldEmp2Rec(1).EMPCTE
      NewEmp2Rec(1).EMPCTUSE = OldEmp2Rec(1).EMPCTUSE
      NewEmp2Rec(1).EMPCTBAL = OldEmp2Rec(1).EMPCTBAL
      NewEmp2Rec(1).PERERN = 0
      NewEmp2Rec(1).PERUSED = 0
      NewEmp2Rec(1).PERBAL = 0
      NewEmp2Rec(1).HOLERN = 0
      NewEmp2Rec(1).HOLUSED = 0
      NewEmp2Rec(1).HOLBAL = 0
      NewEmp2Rec(1).LeaveTbl = OldEmp2Rec(1).LeaveTbl
      NewEmp2Rec(1).ExcludeESC = "N"
      NewEmp2Rec(1).UseLife = ""
      '    UseHosp  AS STRING * 1
      '    UseLife  AS STRING * 1
      NewEmp2Rec(1).LastTransRec = OldEmp2Rec(1).LastTransRec
      NewEmp2Rec(1).EmpPin = OldEmp2Rec(1).EmpPin
      NewEmp2Rec(1).Deleted = OldEmp2Rec(1).Deleted
      NewEmp2Rec(1).Unused = ""
      PUT #2, , NewEmp2Rec(1)
   NEXT
   CLOSE
   KILL "prdata\oemp2.dat"
   PRINT "Employee file Conversion Completed OK."
  ELSE
    PRINT "Current Employee file structure is OK."
  END IF
  ERASE OldEmp2Rec, NewEmp2Rec
  OKFlag = -1
END SUB

SUB CheckTransFile

  REDIM OldTrans(1) AS oTransRecType1
  REDIM NewTrans(1) AS TransRecType
  OLen = LEN(OldTrans(1))
  NLen = LEN(NewTrans(1))

  IF FileSize(TransHistFileName) MOD NLen <> 0 THEN
    PRINT "Converting Transaction History file."

    IF Exist%("prdata\Transh.old") THEN
      KillFile "prdata\Transh.old"
    END IF

    NAME TransHistFileName AS "prdata\Transh.old"

    OPEN "prdata\Transh.old" FOR RANDOM AS #1 LEN = OLen
    OPEN TransHistFileName FOR RANDOM AS #2 LEN = NLen

    NumOfRec = LOF(1) \ OLen

    FOR Cnt = 1 TO NumOfRec
      GET #1, , OldTrans(1)

      NewTrans(1).TActive = OldTrans(1).TActive
      NewTrans(1).PrevTransRec = OldTrans(1).PrevTransRec
      NewTrans(1).EmpPin = OldTrans(1).EmpPin
      NewTrans(1).PaySFlag = OldTrans(1).PaySFlag
      NewTrans(1).CHECKNUM = OldTrans(1).CHECKNUM
      NewTrans(1).PayPdStart = OldTrans(1).PayPdStart
      NewTrans(1).PayPdEnd = OldTrans(1).PayPdEnd
      NewTrans(1).CheckDate = OldTrans(1).CheckDate
      NewTrans(1).PostDate = OldTrans(1).PostDate
      NewTrans(1).PayType = OldTrans(1).PayType
      NewTrans(1).BaseRate = OldTrans(1).BaseRate
      NewTrans(1).OTRate = OldTrans(1).OTRate
      NewTrans(1).RegHrsWork = OldTrans(1).RegHrsWork
      NewTrans(1).VacUsed = OldTrans(1).VacUsed
      NewTrans(1).SickUsed = OldTrans(1).SickUsed
      NewTrans(1).CompUsed = OldTrans(1).CompUsed
      NewTrans(1).HOLHOURS = OldTrans(1).HOLHOURS

      NewTrans(1).PerHours = 0

      NewTrans(1).RegHrsPaid = OldTrans(1).RegHrsPaid
      NewTrans(1).OTHours = OldTrans(1).OTHours
      NewTrans(1).OTHrsPaid = OldTrans(1).OTHrsPaid
      NewTrans(1).OT2Comp = OldTrans(1).OT2Comp
      FOR zz = 1 TO 8
        LSET NewTrans(1).TDist(zz) = OldTrans(1).TDist(zz)
      NEXT

      NewTrans(1).TotRegWage = OldTrans(1).TotRegWage
      NewTrans(1).TotOTWage = OldTrans(1).TotOTWage
      NewTrans(1).GrossWage = OldTrans(1).GrossWage

      FOR zz = 1 TO 3
        NewTrans(1).EAmt(zz) = OldTrans(1).EAmt(zz)
      NEXT

      FOR zz = 1 TO 6
        LSET NewTrans(1).EDist(zz) = OldTrans(1).EDist(zz)
      NEXT

      NewTrans(1).TotAdditEarn = OldTrans(1).TotAdditEarn
      NewTrans(1).GROSSPAY = OldTrans(1).GROSSPAY
      NewTrans(1).SocGrossPay = OldTrans(1).SocGrossPay
      NewTrans(1).MedGrossPay = OldTrans(1).MedGrossPay
      NewTrans(1).FedGrossPay = OldTrans(1).FedGrossPay
      NewTrans(1).StaGrossPay = OldTrans(1).StaGrossPay
      NewTrans(1).SocTaxAmt = OldTrans(1).SocTaxAmt
      NewTrans(1).MedTaxAmt = OldTrans(1).MedTaxAmt
      NewTrans(1).FedTaxAmt = OldTrans(1).FedTaxAmt
      NewTrans(1).StaTaxAmt = OldTrans(1).StaTaxAmt
      NewTrans(1).TotTaxAmt = OldTrans(1).TotTaxAmt
      NewTrans(1).RetireAmt = OldTrans(1).RetireAmt

      FOR zz = 1 TO 12
        NewTrans(1).DAMT(zz) = OldTrans(1).DAMT(zz)
      NEXT

      NewTrans(1).TotDedAmt = OldTrans(1).TotDedAmt
      NewTrans(1).EICAmt = OldTrans(1).EICAmt
      NewTrans(1).NETPAY = OldTrans(1).NETPAY
      NewTrans(1).PeriodHistRec = OldTrans(1).PeriodHistRec
      NewTrans(1).MatchRetAmt = OldTrans(1).MatchRetAmt
      NewTrans(1).MatchSocAmt = OldTrans(1).MatchSocAmt
      NewTrans(1).MatchMedAmt = OldTrans(1).MatchMedAmt
      NewTrans(1).RetGrossPay = OldTrans(1).RetGrossPay
      NewTrans(1).TaxFring = OldTrans(1).TaxFring

      NewTrans(1).Pad1 = ""
    
      PUT #2, , NewTrans(1)
    NEXT
    CLOSE
    KILL "prdata\Transh.old"
    PRINT "Transaction History Conversion Completed OK."
  ELSE
    PRINT "Current Transaction History file structure is OK."
  END IF

  ERASE OldTrans, NewTrans

  OKFlag = -1


END SUB

SUB CheckUnitFile
  
  REDIM OldUnitRec(1) AS UnitFileRecType1
  REDIM NewUnitRec(1) AS UnitFileRecType
  
  OLen = LEN(OldUnitRec(1))
  NLen = LEN(NewUnitRec(1))
  
  IF FileSize(UnitFileName) <> NLen THEN
    PRINT "Converting UNIT file structure."
    OPEN UnitFileName FOR RANDOM SHARED AS #1 LEN = OLen
    OPEN UnitFileName FOR RANDOM SHARED AS #2 LEN = NLen
    GET #1, , OldUnitRec(1)
    LSET NewUnitRec(1) = OldUnitRec(1)

    NewUnitRec(1).USEACH = ""
    NewUnitRec(1).IMMDNUM = ""
    NewUnitRec(1).IMMONUM = ""
    NewUnitRec(1).DBANK = ""
    NewUnitRec(1).OBANK = ""
    NewUnitRec(1).FileVer = 9.1
    PUT #2, , NewUnitRec(1)
    CLOSE
    PRINT "Unit File Conversion Completed OK."
  ELSE
    PRINT "Current UNIT file structure is correct."
  END IF
  ERASE OldUnitRec, NewUnitRec
  
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


'  OPEN EmpIdxNName FOR RANDOM AS #1 LEN = 2
'  FOR Cnt = 1 TO 21  'NumOfRec
'    PUT #1, , Emp1Data(Cnt).Data1RecNum
'  NEXT
'  FOR Cnt = 23 TO NumOfRec
'    PUT #1, , Emp1Data(Cnt).Data1RecNum
'  NEXT

  CLOSE
  
  SortT Emp1Data(1), NumOfRec, 0, Emp1Len, 10, 24               'last name
  
  'this sorts the first names for each last name
'  First = 1
'  DO
'    FOR Cnt = First TO NumOfRec
'      IF Cnt = NumOfRec THEN EXIT FOR
'      IF INSTR(Emp1Data(Cnt).EMPLNAME, Emp1Data(Cnt + 1).EMPLNAME) THEN
'        Last = Cnt
'        DO
'          Last = Last + 1
'        LOOP WHILE INSTR(Emp1Data(Cnt).EMPLNAME, Emp1Data(Last).EMPLNAME)
'        SortT Emp1Data(Cnt), Last - Cnt, 0, Emp1Len, 34, 24     'first name
'        First = Last
'        EXIT FOR
'      END IF
'    NEXT
'  LOOP UNTIL Cnt >= NumOfRec
  
  
  OPEN EmpIdxLName FOR RANDOM AS #1 LEN = 2
  FOR Cnt = 1 TO NumOfRec
    PUT #1, Cnt, Emp1Data(Cnt).Data1RecNum
  NEXT
  CLOSE

  KILL EMPNumFileName

  REDIM EmpNum(1) AS EmpNumType
  NumLen = LEN(EmpNum(1))

  OPEN EMPNumFileName FOR RANDOM AS #1 LEN = NumLen
  FOR Cnt = 1 TO NumOfRec
    EmpNum(1).EmpNum = Emp1Data(Cnt).EMPNO
    PUT #1, , EmpNum(1)
  NEXT
  CLOSE

  'EmpNum        AS STRING * 10

'  FOR Cnt = 1 TO 21  'NumOfRec
'    PUT #1, , Emp1Data(Cnt).Data1RecNum
'  NEXT
'  FOR Cnt = 23 TO NumOfRec
'    PUT #1, , Emp1Data(Cnt).Data1RecNum
'  NEXT

  LOCATE , 1
  PRINT "Reindexing Complete.     "
  
  'OPEN EmpData3Name FOR RANDOM AS #4 LEN = Emp3RecLen
  
END SUB

SUB RelinkTransHistory (YearReLink)

  KillFile OldHistFileName

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
  
  REDIM TPntr(0 TO 400)
  
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
  TNumOfRec = FileSize&(TransHistFileName) \ TranRecLen

  PRINT
  PRINT "Reading Transaction history PIN numbers: ";
  Col = POS(0)
  OPEN TransHistFileName FOR RANDOM AS #1 LEN = TranRecLen
  TNumOfRec = LOF(1) \ TranRecLen

  REDIM TPins(1 TO TNumOfRec)

  FOR Cnt = 1 TO TNumOfRec
    LOCATE , Col: PRINT Cnt;
    GET #1, , TranRec
    TPins(Cnt) = TranRec.EmpPin
  NEXT
  CLOSE

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
  KillFile OldHistFileName
  
  EXIT SUB
  
  
GetTransRecNums:
  REDIM TPntr(0 TO 400)
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
    TranRec.RetGrossPay = TranRec.GROSSPAY
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
  Emp3Rec.YTDGrossPay = UtilRound(Emp3Rec.YTDGrossPay + TranRec.GROSSPAY)
  Emp3Rec.YTDFedGrossPay = UtilRound(Emp3Rec.YTDFedGrossPay + TranRec.FedGrossPay)
  Emp3Rec.YTDStaGrossPay = UtilRound(Emp3Rec.YTDStaGrossPay + TranRec.StaGrossPay)
  Emp3Rec.YTDSocGrossPay = UtilRound(Emp3Rec.YTDSocGrossPay + TranRec.SocGrossPay)
  Emp3Rec.YTDMedGrossPay = UtilRound(Emp3Rec.YTDMedGrossPay + TranRec.MedGrossPay)
  
  Emp3Rec.YTDRegPay = UtilRound(Emp3Rec.YTDRegPay + TranRec.TotRegWage)
  Emp3Rec.YTDOTPay = UtilRound(Emp3Rec.YTDOTPay + TranRec.TotOTWage)
  
  Emp3Rec.YTDNet = UtilRound(Emp3Rec.YTDNet + TranRec.NETPAY)
  
  Emp3Rec.YTDFederal = UtilRound(Emp3Rec.YTDFederal + TranRec.FedTaxAmt)
  Emp3Rec.YTDState = UtilRound(Emp3Rec.YTDState + TranRec.StaTaxAmt)
  Emp3Rec.YTDSocial = UtilRound(Emp3Rec.YTDSocial + TranRec.SocTaxAmt)
  Emp3Rec.YTDMedicare = UtilRound(Emp3Rec.YTDMedicare + TranRec.MedTaxAmt)
  Emp3Rec.YTDRetire = UtilRound(Emp3Rec.YTDRetire + TranRec.RetireAmt)
  
  'year to date totals on deductions
  FOR Cnt = 1 TO 12
    Emp3Rec.YTDDAmt(Cnt) = UtilRound(Emp3Rec.YTDDAmt(Cnt) + TranRec.DAMT(Cnt))
    Emp3Rec.YTDDAmtT = UtilRound(Emp3Rec.YTDDAmtT + TranRec.DAMT(Cnt))
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

