DEFINT A-Z

DECLARE FUNCTION FUsing$ (number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetEmployerName$ ()
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB PRBuildTransmitFile ()
DECLARE SUB PRDraftEmpList ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB PREmp2Draft ()
DECLARE SUB PRPrenote ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RPTSetupPRN (RPTNum%, FHandle%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
  
  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
  '$INCLUDE: 'FieldInf.bi'    '        "          field information
  '$INCLUDE: 'FormEdit.bi'    '        "          form editing information
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.biw'
  '$INCLUDE: 'PRUNIT.Bi'
  
  '$INCLUDE: 'PRdraft.BI'
  
  CONST False = 0
  CONST True = NOT False

FUNCTION GetEmployerName$
  REDIM UnitFileRec(1)   AS UnitFileRecType
  UnitFileRecLen = LEN(UnitFileRec(1))
  UnitFile = FREEFILE
  OPEN UnitFileName FOR RANDOM SHARED AS UnitFile LEN = UnitFileRecLen
  GET #UnitFile, , UnitFileRec(1)
  CLOSE UnitFile
  GetEmployerName$ = UnitFileRec(1).UFEMPR
  ERASE UnitFileRec
END FUNCTION

SUB PRBuildTransmitFile

  REDIM PPDFInfo(1)     AS PRPPDraftInfoType
  PPDFLen = LEN(PPDFInfo(1))

  PPDFFile = FREEFILE
  OPEN PPDraftInfoFileName FOR RANDOM AS PPDFFile LEN = PPDFLen
  NumOfRecs& = LOF(PPDFFile) \ PPDFLen
  GET PPDFFile, 1, PPDFInfo(1)
  CLOSE

  TDraftDate$ = Num2Date$(PPDFInfo(1).DraftDate)

  ToDay$ = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  
  LibName$ = MiscQLib
  ScrnName$ = "PRDFTFIN"
  
  Employer$ = GetEmployerName$
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  FirstTime = True

  BlockClear
  DisplayMiscScrn ScrnName$
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = TDraftDate$
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Done = True
    CASE ESC
      Done = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL Done
  IF ExitFlag THEN
    EXIT SUB
  END IF
  
ProcessDraft:
  
  DraftDate$ = Form$(1, 0)
  DraftDate$ = LEFT$(DraftDate$, 2) + MID$(DraftDate$, 4, 2) + RIGHT$(DraftDate$, 2)
  
  IF VAL(DraftDate$) = 0 THEN EXIT SUB
  
  ScrnName$ = "PRDFTDF1"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  BlockClear
  DisplayMiscScrn ScrnName$
  
  QPrintRC "Building Record Type 1", 9, 26, -1
  
  PPDFFile = FREEFILE
  OPEN PPDraftInfoFileName FOR RANDOM AS PPDFFile LEN = PPDFLen
  NumOfRecs& = LOF(PPDFFile) \ PPDFLen
  
  ' Process Record Type 1
  GOSUB DraftProcessStep1
  QPrintRC "Building Record Type 1 ..Done!", 9, 26, -1
  QPrintRC "Building Record Type 5", 10, 26, -1
  
  GOSUB DraftProcessStep2
  QPrintRC "Building Record Type 5 ..Done!", 10, 26, -1
  QPrintRC "Building Record Type 6", 11, 26, -1
  
  GOSUB DraftProcessStep3
  QPrintRC "Building Record Type 6 ..Done!", 11, 26, -1
  QPrintRC "Building Record Type 8", 12, 26, -1
  
  GOSUB DraftProcessStep4
  QPrintRC "Building Record Type 8 ..Done!", 12, 26, -1
  QPrintRC "Building Record Type 9", 13, 26, -1
  
  GOSUB DraftProcessStep5
  QPrintRC "Building Record Type 9 ..Done!", 13, 26, -1
  QPrintRC "File Name Is: " + "DD" + DraftDate$, 15, 30, -1
  

  CLOSE

  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
  LOOP UNTIL Frm(1).KeyCode <> 0
  
  EXIT SUB
  
OpenMainDraftInformation:
  REDIM PRDraftRec(1) AS PRDraftRecType
  DraftFile = FREEFILE
  OPEN DraftInfoFileName FOR RANDOM SHARED AS #DraftFile LEN = LEN(PRDraftRec(1))
  GET DraftFile, 1, PRDraftRec(1)
  CLOSE DraftFile
RETURN
  
DraftProcessStep1:
  GOSUB OpenMainDraftInformation
  
  REDIM PRDraftRecord1(1) AS PRDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN Draft1FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord1(1))
  PRDraftRecord1(1).Field1 = "1"
  PRDraftRecord1(1).Field2 = "01"
  PRDraftRecord1(1).Field3 = " " + PRDraftRec(1).BANKDEST
  PRDraftRecord1(1).Field4 = " " + PRDraftRec(1).BANKORIG
  PRDraftRecord1(1).Field5 = ToDay$
  PRDraftRecord1(1).Field6 = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  PRDraftRecord1(1).Field7 = "A"
  PRDraftRecord1(1).Field8 = "094"
  PRDraftRecord1(1).Field9 = "10"
  PRDraftRecord1(1).Field10 = "1"
  PRDraftRecord1(1).Field11 = UCASE$(PRDraftRec(1).BANKNAME)
  PRDraftRecord1(1).Field12 = UCASE$(PRDraftRec(1).BANKLOC)
  PRDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  PUT DraftFileNum, 1, PRDraftRecord1(1)
  CLOSE DraftFileNum
RETURN
  
DraftProcessStep2:
  REDIM PRDraftRecord5(1) AS PRDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN Draft5FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord5(1))
  PRDraftRecord5(1).Field1 = "5"
  PRDraftRecord5(1).Field2 = "200"
  PRDraftRecord5(1).Field3 = LEFT$(Employer$, 16)
  PRDraftRecord5(1).Field4 = "                    "
  PRDraftRecord5(1).Field5 = "1" + PRDraftRec(1).FEDID
  PRDraftRecord5(1).Field6 = "PPD"
  PRDraftRecord5(1).Field7 = "PAYROLL  "
  PRDraftRecord5(1).Field8 = RIGHT$(DraftDate$, 2) + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 3, 2)
  PRDraftRecord5(1).Field9 = RIGHT$(DraftDate$, 2) + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 3, 2)
  PRDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  PRDraftRecord5(1).Field11 = "1"
  PRDraftRecord5(1).Field12 = LEFT$(PRDraftRec(1).BANKORIG, 8)
  PRDraftRecord5(1).Field13 = "0000001"
  PUT DraftFileNum, 1, PRDraftRecord5(1)
  CLOSE DraftFileNum
RETURN
  
DraftProcessStep3:
  COUNTER = 0
  
  REDIM Emp2Rec(1) AS EmpData2Type
  Emp2Len = LEN(Emp2Rec(1))
  
  EmpFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM SHARED AS EmpFile LEN = Emp2Len
  
  REDIM PRDraftRecord6(1) AS PRDraftRecord6Type
  KillFile Draft6FileName
  DraftFileNum = FREEFILE
  OPEN Draft6FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord6(1))
  
  FOR Cnt& = 1 TO NumOfRecs&
    
    GET PPDFFile, Cnt&, PPDFInfo(1)
    GET EmpFile, PPDFInfo(1).EmpRec, Emp2Rec(1)
    
    NetPay# = Round#(PPDFInfo(1).NetPay)
    
    IF NetPay# > 0 THEN ' Process Employee's Here
      NetAmt$ = QPTrim$(STR$(NetPay# * 100))    'Remove the Decimals
      IF LEN(NetAmt$) < 10 THEN NetAmt$ = STRING$(10 - LEN(NetAmt$), "0") + NetAmt$
      'Get Running Net Amount
      NetTotal# = Round#(NetTotal# + (NetPay# * 100))

      COUNTER = COUNTER + 1

      AcctNumber$ = QPTrim$(STR$(Cnt&))
      IF LEN(AcctNumber$) < 15 THEN
        AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
      END IF
      
      nme$ = QPTrim$(Emp2Rec(1).EMPFNAME) + " " + QPTrim$(Emp2Rec(1).EMPLNAME)
      IF LEN(nme$) < 22 THEN
        nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
      ELSE
        nme$ = LEFT$(nme$, 22)
      END IF
      
      BANKACCT$ = QPTrim$(Emp2Rec(1).EMPDDACC)
      
      IF LEN(BANKACCT$) < 17 THEN BANKACCT$ = BANKACCT$ + STRING$(17 - LEN(BANKACCT$), 32)
      Trace = Trace + 1
      Trace$ = QPTrim$(STR$(Trace))
      IF LEN(Trace$) < 7 THEN
        Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
      END IF
      
      PRDraftRecord6(1).Field1 = "6"

      IF Emp2Rec(1).DraftCod = "C" THEN
        PRDraftRecord6(1).Field2 = "22"           'Designates Credit Checking
      ELSEIF Emp2Rec(1).DraftCod = "S" THEN
        PRDraftRecord6(1).Field2 = "32"           'Designates Credit Checking
      END IF

      PRDraftRecord6(1).Field3 = LEFT$(Emp2Rec(1).TRANSIT, 8)
      PRDraftRecord6(1).Field4 = RIGHT$(Emp2Rec(1).TRANSIT, 1)
      PRDraftRecord6(1).Field5 = LEFT$(BANKACCT$, 17)
      PRDraftRecord6(1).Field6 = NetAmt$        'The amt to credit
      PRDraftRecord6(1).Field7 = AcctNumber$
      PRDraftRecord6(1).Field8 = UCASE$(nme$)
      PRDraftRecord6(1).Field9 = "  "
      PRDraftRecord6(1).Field10 = "0"
      PRDraftRecord6(1).Field11 = LEFT$(PRDraftRec(1).BANKORIG, 8) + Trace$
      PUT DraftFileNum, COUNTER, PRDraftRecord6(1)
      hash# = hash# + VAL(LEFT$(Emp2Rec(1).TRANSIT, 8))
      number = number + 1
    END IF
  NEXT
  
  CLOSE DraftFileNum
RETURN
  
DraftProcessStep4:
  hash$ = QPTrim$(STR$(hash#))
  
  IF LEN(hash$) < 10 THEN
    hash$ = STRING$(10 - LEN(hash$), "0") + hash$
  END IF
  IF LEN(hash$) > 10 THEN
    hash$ = RIGHT$(hash$, 10)
  END IF
  
  IF LEN(Trace$) > 6 THEN Trace$ = RIGHT$(Trace$, 6)
  TotalAmount$ = QPTrim$(STR$(NetTotal#))
  IF LEN(TotalAmount$) < 12 THEN TotalAmount$ = STRING$(12 - LEN(TotalAmount$), "0") + TotalAmount$
  
  REDIM PRDraftRecord8(1) AS PRDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN Draft8FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord8(1))
  PRDraftRecord8(1).Field1 = "8"
  PRDraftRecord8(1).Field2 = "200"
  PRDraftRecord8(1).Field3 = Trace$
  PRDraftRecord8(1).Field4 = hash$
  PRDraftRecord8(1).Field5 = "000000000000"     ' zero for prenote
  PRDraftRecord8(1).Field6 = TotalAmount$       ' zero for prenote
  PRDraftRecord8(1).Field7 = "1" + PRDraftRec(1).FEDID
  PRDraftRecord8(1).Field8 = STRING$(19, 32)    ' Reserved
  PRDraftRecord8(1).Field9 = STRING$(6, 32)     ' Reserved for Federal Reserve Use
  PRDraftRecord8(1).Field10 = LEFT$(PRDraftRec(1).BANKORIG, 8)
  PRDraftRecord8(1).Field11 = "0000001"
  PUT DraftFileNum, 1, PRDraftRecord8(1)
  CLOSE DraftFileNum
RETURN
  
DraftProcessStep5:
  TotSize# = VAL(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records
  
  IF BlockSize! <> INT(BlockSize!) THEN
    BlockSize! = INT(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize! - 1)))
  ELSE
    FillSize! = 0
  END IF
  
  BlockSize$ = QPTrim$(STR$(BlockSize!))

  IF LEN(BlockSize$) < 6 THEN BlockSize$ = STRING$(6 - LEN(BlockSize$), "0") + BlockSize$
  IF LEN(Trace$) < 8 THEN Trace$ = STRING$(8 - LEN(Trace$), "0") + Trace$
  
  REDIM PRDraftRecord9(1) AS PRDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN Draft9FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord9(1))
  PRDraftRecord9(1).Field1 = "9"
  PRDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  PRDraftRecord9(1).Field3 = BlockSize$
  PRDraftRecord9(1).Field4 = Trace$
  PRDraftRecord9(1).Field5 = hash$
  PRDraftRecord9(1).Field6 = "000000000000"
  PRDraftRecord9(1).Field7 = TotalAmount$      ' zero for prenote
  PRDraftRecord9(1).Field8 = STRING$(39, 32)    ' Reserved
  PUT DraftFileNum, 1, PRDraftRecord9(1)
  CLOSE DraftFileNum
  
  ' Now Put Them Together In File Name UBDFNOTE
  OutFile = FREEFILE
  OPEN "O", OutFile, "DD" + DraftDate$: WIDTH #OutFile, 255
  
  REDIM PRDraftRecord1(1) AS PRDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN Draft1FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord1(1))
  GET DraftFileNum, 1, PRDraftRecord1(1)
  PRINT #OutFile, PRDraftRecord1(1).Field1;
  PRINT #OutFile, PRDraftRecord1(1).Field2;
  PRINT #OutFile, PRDraftRecord1(1).Field3;
  PRINT #OutFile, PRDraftRecord1(1).Field4;
  PRINT #OutFile, PRDraftRecord1(1).Field5;
  PRINT #OutFile, PRDraftRecord1(1).Field6;
  PRINT #OutFile, PRDraftRecord1(1).Field7;
  PRINT #OutFile, PRDraftRecord1(1).Field8;
  PRINT #OutFile, PRDraftRecord1(1).Field9;
  PRINT #OutFile, PRDraftRecord1(1).Field10;
  PRINT #OutFile, PRDraftRecord1(1).Field11;
  PRINT #OutFile, PRDraftRecord1(1).Field12;
  PRINT #OutFile, PRDraftRecord1(1).Field13
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord5(1) AS PRDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN Draft5FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord5(1))
  GET DraftFileNum, 1, PRDraftRecord5(1)
  PRINT #OutFile, PRDraftRecord5(1).Field1;
  PRINT #OutFile, PRDraftRecord5(1).Field2;
  PRINT #OutFile, PRDraftRecord5(1).Field3;
  PRINT #OutFile, PRDraftRecord5(1).Field4;
  PRINT #OutFile, PRDraftRecord5(1).Field5;
  PRINT #OutFile, PRDraftRecord5(1).Field6;
  PRINT #OutFile, PRDraftRecord5(1).Field7;
  PRINT #OutFile, PRDraftRecord5(1).Field8;
  PRINT #OutFile, PRDraftRecord5(1).Field9;
  PRINT #OutFile, PRDraftRecord5(1).Field10;
  PRINT #OutFile, PRDraftRecord5(1).Field11;
  PRINT #OutFile, PRDraftRecord5(1).Field12;
  PRINT #OutFile, PRDraftRecord5(1).Field13
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord6(1) AS PRDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN Draft6FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94
  
  FOR Cnt = 1 TO NumOfLines
    GET DraftFileNum, Cnt, PRDraftRecord6(1)
    PRINT #OutFile, PRDraftRecord6(1).Field1;
    PRINT #OutFile, PRDraftRecord6(1).Field2;
    PRINT #OutFile, PRDraftRecord6(1).Field3;
    PRINT #OutFile, PRDraftRecord6(1).Field4;
    PRINT #OutFile, PRDraftRecord6(1).Field5;
    PRINT #OutFile, PRDraftRecord6(1).Field6;
    PRINT #OutFile, PRDraftRecord6(1).Field7;
    PRINT #OutFile, PRDraftRecord6(1).Field8;
    PRINT #OutFile, PRDraftRecord6(1).Field9;
    PRINT #OutFile, PRDraftRecord6(1).Field10;
    PRINT #OutFile, PRDraftRecord6(1).Field11
  NEXT Cnt
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord8(1) AS PRDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN Draft8FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord8(1))
  GET DraftFileNum, 1, PRDraftRecord8(1)
  PRINT #OutFile, PRDraftRecord8(1).Field1;
  PRINT #OutFile, PRDraftRecord8(1).Field2;
  PRINT #OutFile, PRDraftRecord8(1).Field3;
  PRINT #OutFile, PRDraftRecord8(1).Field4;
  PRINT #OutFile, PRDraftRecord8(1).Field5;
  PRINT #OutFile, PRDraftRecord8(1).Field6;
  PRINT #OutFile, PRDraftRecord8(1).Field7;
  PRINT #OutFile, PRDraftRecord8(1).Field8;
  PRINT #OutFile, PRDraftRecord8(1).Field9;
  PRINT #OutFile, PRDraftRecord8(1).Field10;
  PRINT #OutFile, PRDraftRecord8(1).Field11
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord9(1) AS PRDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN Draft9FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord9(1))
  GET DraftFileNum, 1, PRDraftRecord9(1)
  PRINT #OutFile, PRDraftRecord9(1).Field1;
  PRINT #OutFile, PRDraftRecord9(1).Field2;
  PRINT #OutFile, PRDraftRecord9(1).Field3;
  PRINT #OutFile, PRDraftRecord9(1).Field4;
  PRINT #OutFile, PRDraftRecord9(1).Field5;
  PRINT #OutFile, PRDraftRecord9(1).Field6;
  PRINT #OutFile, PRDraftRecord9(1).Field7;
  PRINT #OutFile, PRDraftRecord9(1).Field8
  CLOSE DraftFileNum

  FOR Cnt = 1 TO FillSize! / 94
    PRINT #OutFile, STRING$(94, "9")
  NEXT Cnt
  CLOSE

RETURN
  
  
END SUB

SUB PRDraftEmpList
  
  ExitFlag = False
  OKFlag = False
  
  Image1$ = ",####.##"
  
  REDIM Dash(1) AS STRING * 78
  
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  'REDIM Pg(1)          AS STRING * 3
  
  REDIM ENumb(1) AS STRING * 11
  REDIM EName(1) AS STRING * 23
  REDIM ETitle(1) AS STRING * 12
  
  MaxLines = 55
  
  LineCnt = 0
  Dash(1) = STRING$(78, "-")
  EmpRecSize = LEN(Emp2Rec(1))
  
  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs)
  
  FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, NumOfRecs
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  RptTitle$ = "Employee Draft Listing"
  ShowProcessingScrn RptTitle$
  
  RptName$ = "PRRPTS\EMPDFLST.RPT"
  
  FCreate RptName$
  FOpenS RptName$, RptHandel
  RPTSetupPRN 10, RptHandel
  FOpenS EmpData2Name, EHandle  'open employee data file
  
  GOSUB PrintEmpListHeader
  
  FOR RecNo = 1 TO NumOfRecs
    FGetRTA EHandle, Emp2Rec(1), CLNG(IdxBuff(RecNo)), EmpRecSize
    IF NOT Emp2Rec(1).Deleted THEN
      IF Emp2Rec(1).DraftCod = "C" OR Emp2Rec(1).DraftCod = "S" THEN
        EmpCnt = EmpCnt + 1
        GOSUB PrintEmpListData
        IF LineCnt >= MaxLines THEN
          FPut RptHandel, FF$
          GOSUB PrintEmpListHeader
        END IF
      END IF
    END IF
    
SkipThisOne:
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  GOSUB PrintEmpListTotals
  RPTSetupPRN 0, RptHandel
  FClose EHandle
  FClose RptHandel
  
  ERASE Dash, Emp2Rec, Unit, ENumb, EName, ETitle, IdxBuff
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
  
PrintEmpListHeader:
  Page = Page + 1
  FPut RptHandel, Unit(1).UFEMPR + CrLf$
  FPut RptHandel, "Employee Draft Listing" + CrLf$
  FPut RptHandel, "Report Date: " + DATE$ + "                                            Page: " + QPStrI$(Page) + CrLf$
  FPut RptHandel, CrLf$
  FPut RptHandel, "    Number     Employee Name           Bank Acct No.        Transit No." + CrLf$
  FPut RptHandel, " Prenoted?     Bank Name                         Bank Location" + CrLf$
  FPut RptHandel, Dash$(1) + CrLf$
  LineCnt = 7
  RETURN
  
PrintEmpListData:
  ENumb(1) = Emp2Rec(1).EMPNO
  EName(1) = QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME)
  ETitle(1) = Emp2Rec(1).TRANSIT
  FPut RptHandel, ENumb(1) + "    " + EName(1) + " " + Emp2Rec(1).DraftCod + " " + Emp2Rec(1).EMPDDACC + " " + ETitle(1) + CrLf$
  IF Emp2Rec(1).PreNoted = "Y" THEN
    ENumb(1) = "Y"
  ELSE
    ENumb(1) = "N"
  END IF
  FPut RptHandel, ENumb(1) + "    " + Emp2Rec(1).BANKNAME + Emp2Rec(1).BANKLOC + CrLf$
  
  LineCnt = LineCnt + 2
  RETURN
  
PrintEmpListTotals:
  FPut RptHandel, Dash$(1) + CrLf$
  FPut RptHandel, "Total Employees: " + QPStrI(EmpCnt) + CrLf$
  FPut RptHandel, FF$
  RETURN
  
END SUB

SUB PRDraftMenu
  DO
    
    CursorOff
    
    Choice = 0
    
    PayMenu DraftMenu, Choice, 4
    
    SELECT CASE Choice
    CASE 1
      PREmp2Draft
    CASE 2
      PRBuildTransmitFile
    CASE 3
      PRPrenote
    CASE 4
      PRDraftEmpList
      
    END SELECT
    
  LOOP UNTIL Choice = EscKey
  
END SUB

SUB PREmp2Draft
  ExitFlag = False
  OKFlag = False
  
  Image1$ = ",####.##"
  
  REDIM Dash(1) AS STRING * 78
  
  REDIM Emp2Rec(1)     AS EmpData2Type
  REDIM Unit(1)        AS UnitFileRecType
  REDIM PPDFInfo(1)    AS PRPPDraftInfoType
  
  REDIM ENumb(1) AS STRING * 11
  REDIM EName(1) AS STRING * 30
  REDIM ETitle(1) AS STRING * 12
  
  MaxLines = 55
  
  LineCnt = 0
  Dash(1) = STRING$(78, "-")
  EmpRecSize = LEN(Emp2Rec(1))
  PPDFLen = LEN(PPDFInfo(1))
  
  FGetAH UnitFileName, Unit(1), LEN(Unit(1)), 1
  
  RptTitle$ = "Pay Period Draft Report"
  ShowProcessingScrn RptTitle$
  
  RptName$ = "PRRPTS\PPDF.RPT"
  
  FCreate RptName$
  FOpenS RptName$, RptHandel
  RPTSetupPRN 10, RptHandel
  FOpenS EmpData2Name, EHandle  'open employee data file
  
  PPDFFile = FREEFILE
  OPEN PPDraftInfoFileName FOR RANDOM AS PPDFFile LEN = PPDFLen
  NumOfRecs = LOF(PPDFFile) \ PPDFLen
  GET #PPDFFile, 1, PPDFInfo(1)
  PPDate$ = Num2Date$(PPDFInfo(1).DraftDate)
  
  GOSUB PrintEmp2DraftHeader
  
  FOR RecNo = 1 TO NumOfRecs
    GET #PPDFFile, RecNo, PPDFInfo(1)
    FGetRTA EHandle, Emp2Rec(1), CLNG(PPDFInfo(1).EmpRec), EmpRecSize
    EmpCnt = EmpCnt + 1
    GOSUB PrintEmp2DraftData
    IF LineCnt >= MaxLines THEN
      FPut RptHandel, FF$
      GOSUB PrintEmp2DraftHeader
    END IF
    
    ShowPctComp RecNo, NumOfRecs
  NEXT
  
  GOSUB PrintEmp2DraftTotals
  RPTSetupPRN 0, RptHandel
  
  FClose EHandle
  FClose RptHandel
  CLOSE PPDFFile
  
  ERASE Dash, Emp2Rec, Unit, ENumb, EName, ETitle
  
  PrintRptFile RptTitle$, RptName$, 1, RetCode, 0
  
  EXIT SUB
  
  
PrintEmp2DraftHeader:
  Page = Page + 1
  FPut RptHandel, Unit(1).UFEMPR + CrLf$
  FPut RptHandel, "Pay Period Draft Report" + CrLf$
  FPut RptHandel, "Draft Date: " + PPDate$ + "                                            Page: " + QPStrI$(Page) + CrLf$
  FPut RptHandel, CrLf$
  FPut RptHandel, "    Number     Employee Name                  Bank Acct No.       Draft Amt." + CrLf$
  FPut RptHandel, Dash$(1) + CrLf$
  LineCnt = 6
  RETURN
  
PrintEmp2DraftData:
  ENumb(1) = Emp2Rec(1).EMPNO
  EName(1) = QPTrim$(Emp2Rec(1).EMPLNAME) + ", " + QPTrim$(Emp2Rec(1).EMPFNAME)
  ETitle(1) = FUsing$(STR$(PPDFInfo(1).NetPay), ",#####.##")
  FPut RptHandel, ENumb(1) + "    " + EName(1) + " " + Emp2Rec(1).DraftCod + " " + Emp2Rec(1).EMPDDACC + " " + ETitle(1) + CrLf$
  'IF Emp2Rec(1).PreNoteFlag <> 0 THEN
  '  ENumb(1) = "Y"
  'ELSE
  '  ENumb(1) = "N"
  'END IF
  'FPut RptHandel, ENumb(1) + "    " + Emp2Rec(1).BANKNAME + Emp2Rec(1).BANKLOC + CrLf$
  '
  LineCnt = LineCnt + 1
  RETURN
  
PrintEmp2DraftTotals:
  FPut RptHandel, Dash$(1) + CrLf$
  FPut RptHandel, "Total Employees: " + QPStrI(EmpCnt) + CrLf$
  FPut RptHandel, FF$
  RETURN
  
  
END SUB

SUB PRPrenote
  
  ToDay$ = RIGHT$(DATE$, 2) + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2)
  ScrnName$ = "PRDFTPRE"
  Employer$ = GetEmployerName$
  
  NumFlds = LibNumberOfFields(MiscQLib, ScrnName$)
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef MiscQLib, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  BlockClear
  DisplayMiscScrn ScrnName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Done = True
    CASE ESC
      Done = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL Done

  IF ExitFlag THEN
    EXIT SUB
  END IF
  
ProcessPrenote:
  
  ScrnName$ = "PREDFT01"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(MiscQLib, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef MiscQLib, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  BlockClear
  DisplayMiscScrn ScrnName$
  
  QPrintRC "Building Record Type 1", 9, 26, -1
  
  ' Process Record Type 1
  GOSUB ProcessStep1
  QPrintRC "Building Record Type 1 ..Done!", 9, 26, -1
  QPrintRC "Building Record Type 5", 10, 26, -1
  
  GOSUB ProcessStep2
  QPrintRC "Building Record Type 5 ..Done!", 10, 26, -1
  QPrintRC "Building Record Type 6", 11, 26, -1
  
  GOSUB ProcessStep3
  QPrintRC "Building Record Type 6 ..Done!", 11, 26, -1
  QPrintRC "Building Record Type 8", 12, 26, -1
  
  GOSUB ProcessStep4
  QPrintRC "Building Record Type 8 ..Done!", 12, 26, -1
  QPrintRC "Building Record Type 9", 13, 26, -1
  
  GOSUB ProcessStep5
  QPrintRC "Building Record Type 9 ..Done!", 13, 26, -1
  QPrintRC "File Name Is: PRDFNOTE.DAT", 15, 28, -1
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
  LOOP UNTIL Frm(1).KeyCode <> 0
  
  EXIT SUB
  
OpenMainDraftInfo:
  REDIM PRDraftRec(1) AS PRDraftRecType
  DraftFile = FREEFILE
  OPEN DraftInfoFileName FOR RANDOM SHARED AS #DraftFile LEN = LEN(PRDraftRec(1))
  GET DraftFile, 1, PRDraftRec(1)
RETURN
  
ProcessStep1:
  GOSUB OpenMainDraftInfo
  
  REDIM PRDraftRecord1(1) AS PRDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN Draft1FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord1(1))
  PRDraftRecord1(1).Field1 = "1"
  PRDraftRecord1(1).Field2 = "01"
  PRDraftRecord1(1).Field3 = " " + PRDraftRec(1).BANKDEST
  PRDraftRecord1(1).Field4 = " " + PRDraftRec(1).BANKORIG
  PRDraftRecord1(1).Field5 = ToDay$
  PRDraftRecord1(1).Field6 = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
  PRDraftRecord1(1).Field7 = "A"
  PRDraftRecord1(1).Field8 = "094"
  PRDraftRecord1(1).Field9 = "10"
  PRDraftRecord1(1).Field10 = "1"
  PRDraftRecord1(1).Field11 = UCASE$(PRDraftRec(1).BANKNAME)
  PRDraftRecord1(1).Field12 = UCASE$(PRDraftRec(1).BANKLOC)
  PRDraftRecord1(1).Field13 = "        "        'Must = 8 Spaces
  PUT DraftFileNum, 1, PRDraftRecord1(1)
  CLOSE DraftFileNum
RETURN
  
ProcessStep2:
  REDIM PRDraftRecord5(1) AS PRDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN Draft5FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord5(1))
  PRDraftRecord5(1).Field1 = "5"
  PRDraftRecord5(1).Field2 = "200"
  PRDraftRecord5(1).Field3 = LEFT$(Employer$, 16)
  PRDraftRecord5(1).Field4 = "                    "
  PRDraftRecord5(1).Field5 = "1" + PRDraftRec(1).FEDID
  PRDraftRecord5(1).Field6 = "PPD"
  PRDraftRecord5(1).Field7 = "PAYROLL  "
  PRDraftRecord5(1).Field8 = ToDay$
  PRDraftRecord5(1).Field9 = ToDay$
  PRDraftRecord5(1).Field10 = "   "             'Reserved w/3 blanks
  PRDraftRecord5(1).Field11 = "1"
  PRDraftRecord5(1).Field12 = LEFT$(PRDraftRec(1).BANKORIG, 8)
  PRDraftRecord5(1).Field13 = "0000001"
  PUT DraftFileNum, 1, PRDraftRecord5(1)
  CLOSE DraftFileNum
RETURN
  
ProcessStep3:
  COUNTER = 0
  
  REDIM PRDraftRecord6(1) AS PRDraftRecord6Type
  KillFile Draft6FileName
  DraftFileNum = FREEFILE
  OPEN Draft6FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord6(1))
  
  'GO THRU DATA FILE HERE
  REDIM Emp2Rec(1) AS EmpData2Type
  
  'working here
  Emp2Len = LEN(Emp2Rec(1))
  NumOfRecs = FileSize(EmpData2Name) \ Emp2Len
  EmpFile = FREEFILE
  OPEN EmpData2Name FOR RANDOM SHARED AS EmpFile LEN = Emp2Len
  FOR Cnt = 1 TO NumOfRecs
    GET EmpFile, Cnt, Emp2Rec(1)
    EmpRecNum = Cnt
    '  Process Customer Here
    IF Emp2Rec(1).Deleted = 0 THEN
      IF (Emp2Rec(1).DraftCod = "C" OR Emp2Rec(1).DraftCod = "S") AND Emp2Rec(1).PreNoted <> "Y" THEN
        COUNTER = COUNTER + 1
        AcctNumber$ = QPTrim$(STR$(EmpRecNum))
        IF LEN(AcctNumber$) < 15 THEN
          AcctNumber$ = AcctNumber$ + STRING$(15 - LEN(AcctNumber$), 32)
        END IF
        nme$ = QPTrim$(Emp2Rec(1).EMPFNAME) + " " + QPTrim$(Emp2Rec(1).EMPLNAME)
        IF LEN(nme$) < 22 THEN
          nme$ = nme$ + STRING$(22 - LEN(nme$), 32)
        ELSE
          nme$ = LEFT$(nme$, 22)
        END IF
      
        BANKACCT$ = QPTrim$(Emp2Rec(1).EMPDDACC)
      
        IF LEN(BANKACCT$) < 17 THEN BANKACCT$ = BANKACCT$ + STRING$(17 - LEN(BANKACCT$), 32)
        Trace = Trace + 1
        Trace$ = QPTrim$(STR$(Trace))
        IF LEN(Trace$) < 7 THEN Trace$ = STRING$(7 - LEN(Trace$), "0") + Trace$
      
        PRDraftRecord6(1).Field1 = "6"

        IF Emp2Rec(1).DraftCod = "C" THEN
          PRDraftRecord6(1).Field2 = "23"           'Designates Credit Checking
        ELSEIF Emp2Rec(1).DraftCod = "S" THEN
          PRDraftRecord6(1).Field2 = "33"           'Designates Credit Checking
        END IF

        PRDraftRecord6(1).Field3 = LEFT$(Emp2Rec(1).TRANSIT, 8)
        PRDraftRecord6(1).Field4 = RIGHT$(Emp2Rec(1).TRANSIT, 1)
        PRDraftRecord6(1).Field5 = LEFT$(BANKACCT$, 17)
        PRDraftRecord6(1).Field6 = "0000000000"   ' All zero's for Prenote
        PRDraftRecord6(1).Field7 = AcctNumber$
        PRDraftRecord6(1).Field8 = UCASE$(nme$)
        PRDraftRecord6(1).Field9 = "  "
        PRDraftRecord6(1).Field10 = "0"
        PRDraftRecord6(1).Field11 = LEFT$(PRDraftRec(1).BANKORIG, 8) + Trace$
        PUT DraftFileNum, COUNTER, PRDraftRecord6(1)
        hash# = hash# + VAL(LEFT$(Emp2Rec(1).TRANSIT, 8))

        Emp2Rec(1).PreNoted = "Y"

        'unrem
        PUT EmpFile, Cnt, Emp2Rec(1)
        number = number + 1
        'ELSE
        '  Emp2Rec(1).PreNoteFlag = 0
        '  PUT EmpFile, Cnt, Emp2Rec(1)
      END IF
    END IF
  NEXT Cnt
  CLOSE DraftFileNum
RETURN
  
ProcessStep4:
  hash$ = QPTrim$(STR$(hash#))
  
  IF LEN(hash$) < 10 THEN
    hash$ = STRING$(10 - LEN(hash$), "0") + hash$
  END IF
  IF LEN(hash$) > 10 THEN
    hash$ = RIGHT$(hash$, 10)
  END IF
  
  IF LEN(Trace$) > 6 THEN Trace$ = RIGHT$(Trace$, 6)
  
  REDIM PRDraftRecord8(1) AS PRDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN Draft8FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord8(1))
  PRDraftRecord8(1).Field1 = "8"
  PRDraftRecord8(1).Field2 = "200"
  PRDraftRecord8(1).Field3 = Trace$
  PRDraftRecord8(1).Field4 = hash$
  PRDraftRecord8(1).Field5 = "000000000000"     ' zero for prenote
  PRDraftRecord8(1).Field6 = "000000000000"     ' zero for prenote
  PRDraftRecord8(1).Field7 = "1" + PRDraftRec(1).FEDID
  PRDraftRecord8(1).Field8 = STRING$(19, 32)    ' Reserved
  PRDraftRecord8(1).Field9 = STRING$(6, 32)     ' Reserved for Federal Reserve Use
  PRDraftRecord8(1).Field10 = LEFT$(PRDraftRec(1).BANKORIG, 8)
  PRDraftRecord8(1).Field11 = "0000001"
  PUT DraftFileNum, 1, PRDraftRecord8(1)
  CLOSE DraftFileNum
RETURN
  
ProcessStep5:
  TotSize# = VAL(Trace$) + 4    ' Total Records= Trace + 4 control records
  TotSize# = TotSize# * 94      ' Total Bytes = 94 per record
  BlockSize! = TotSize# / 940   ' Rem Blocks Consist of Batchs of 10 Records
  
  IF BlockSize! <> INT(BlockSize!) THEN
    BlockSize! = INT(BlockSize!) + 1
    FillSize! = 940 - (TotSize# - (940 * (BlockSize - 1)))
  ELSE
    FillSize! = 0
  END IF
  
  BlockSize$ = QPTrim$(STR$(BlockSize!))

  IF LEN(BlockSize$) < 6 THEN BlockSize$ = STRING$(6 - LEN(BlockSize$), "0") + BlockSize$
  IF LEN(Trace$) < 8 THEN Trace$ = STRING$(8 - LEN(Trace$), "0") + Trace$
  
  REDIM PRDraftRecord9(1) AS PRDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN Draft9FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord9(1))
  PRDraftRecord9(1).Field1 = "9"
  PRDraftRecord9(1).Field2 = "000001"           ' only 1 batch
  PRDraftRecord9(1).Field3 = BlockSize$
  PRDraftRecord9(1).Field4 = Trace$
  PRDraftRecord9(1).Field5 = hash$
  PRDraftRecord9(1).Field6 = "000000000000"     ' zero for prenote
  PRDraftRecord9(1).Field7 = "000000000000"
  PRDraftRecord9(1).Field8 = STRING$(39, 32)    ' Reserved
  PUT DraftFileNum, 1, PRDraftRecord9(1)
  CLOSE DraftFileNum
  
  ' Now Put Them Together In File Name UBDFNOTE
  OutFile = FREEFILE
  OPEN "O", OutFile, "PRDFNOTE.DAT": WIDTH #OutFile, 255
  
  REDIM PRDraftRecord1(1) AS PRDraftRecord1Type
  DraftFileNum = FREEFILE
  OPEN Draft1FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord1(1))
  GET DraftFileNum, 1, PRDraftRecord1(1)
  PRINT #OutFile, PRDraftRecord1(1).Field1;
  PRINT #OutFile, PRDraftRecord1(1).Field2;
  PRINT #OutFile, PRDraftRecord1(1).Field3;
  PRINT #OutFile, PRDraftRecord1(1).Field4;
  PRINT #OutFile, PRDraftRecord1(1).Field5;
  PRINT #OutFile, PRDraftRecord1(1).Field6;
  PRINT #OutFile, PRDraftRecord1(1).Field7;
  PRINT #OutFile, PRDraftRecord1(1).Field8;
  PRINT #OutFile, PRDraftRecord1(1).Field9;
  PRINT #OutFile, PRDraftRecord1(1).Field10;
  PRINT #OutFile, PRDraftRecord1(1).Field11;
  PRINT #OutFile, PRDraftRecord1(1).Field12;
  PRINT #OutFile, PRDraftRecord1(1).Field13
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord5(1) AS PRDraftRecord5Type
  DraftFileNum = FREEFILE
  OPEN Draft5FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord5(1))
  GET DraftFileNum, 1, PRDraftRecord5(1)
  PRINT #OutFile, PRDraftRecord5(1).Field1;
  PRINT #OutFile, PRDraftRecord5(1).Field2;
  PRINT #OutFile, PRDraftRecord5(1).Field3;
  PRINT #OutFile, PRDraftRecord5(1).Field4;
  PRINT #OutFile, PRDraftRecord5(1).Field5;
  PRINT #OutFile, PRDraftRecord5(1).Field6;
  PRINT #OutFile, PRDraftRecord5(1).Field7;
  PRINT #OutFile, PRDraftRecord5(1).Field8;
  PRINT #OutFile, PRDraftRecord5(1).Field9;
  PRINT #OutFile, PRDraftRecord5(1).Field10;
  PRINT #OutFile, PRDraftRecord5(1).Field11;
  PRINT #OutFile, PRDraftRecord5(1).Field12;
  PRINT #OutFile, PRDraftRecord5(1).Field13
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord6(1) AS PRDraftRecord6Type
  DraftFileNum = FREEFILE
  OPEN Draft6FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord6(1))
  NumOfLines = LOF(DraftFileNum) / 94
  
  FOR Cnt = 1 TO NumOfLines
    GET DraftFileNum, Cnt, PRDraftRecord6(1)
    PRINT #OutFile, PRDraftRecord6(1).Field1;
    PRINT #OutFile, PRDraftRecord6(1).Field2;
    PRINT #OutFile, PRDraftRecord6(1).Field3;
    PRINT #OutFile, PRDraftRecord6(1).Field4;
    PRINT #OutFile, PRDraftRecord6(1).Field5;
    PRINT #OutFile, PRDraftRecord6(1).Field6;
    PRINT #OutFile, PRDraftRecord6(1).Field7;
    PRINT #OutFile, PRDraftRecord6(1).Field8;
    PRINT #OutFile, PRDraftRecord6(1).Field9;
    PRINT #OutFile, PRDraftRecord6(1).Field10;
    PRINT #OutFile, PRDraftRecord6(1).Field11
  NEXT Cnt
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord8(1) AS PRDraftRecord8Type
  DraftFileNum = FREEFILE
  OPEN Draft8FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord8(1))
  GET DraftFileNum, 1, PRDraftRecord8(1)
  PRINT #OutFile, PRDraftRecord8(1).Field1;
  PRINT #OutFile, PRDraftRecord8(1).Field2;
  PRINT #OutFile, PRDraftRecord8(1).Field3;
  PRINT #OutFile, PRDraftRecord8(1).Field4;
  PRINT #OutFile, PRDraftRecord8(1).Field5;
  PRINT #OutFile, PRDraftRecord8(1).Field6;
  PRINT #OutFile, PRDraftRecord8(1).Field7;
  PRINT #OutFile, PRDraftRecord8(1).Field8;
  PRINT #OutFile, PRDraftRecord8(1).Field9;
  PRINT #OutFile, PRDraftRecord8(1).Field10;
  PRINT #OutFile, PRDraftRecord8(1).Field11
  CLOSE DraftFileNum
  
  REDIM PRDraftRecord9(1) AS PRDraftRecord9Type
  DraftFileNum = FREEFILE
  OPEN Draft9FileName FOR RANDOM SHARED AS #DraftFileNum LEN = LEN(PRDraftRecord9(1))
  GET DraftFileNum, 1, PRDraftRecord9(1)
  PRINT #OutFile, PRDraftRecord9(1).Field1;
  PRINT #OutFile, PRDraftRecord9(1).Field2;
  PRINT #OutFile, PRDraftRecord9(1).Field3;
  PRINT #OutFile, PRDraftRecord9(1).Field4;
  PRINT #OutFile, PRDraftRecord9(1).Field5;
  PRINT #OutFile, PRDraftRecord9(1).Field6;
  PRINT #OutFile, PRDraftRecord9(1).Field7;
  PRINT #OutFile, PRDraftRecord9(1).Field8
  CLOSE DraftFileNum
  FOR Cnt = 1 TO FillSize! / 94
    PRINT #OutFile, STRING$(94, "9")
  NEXT Cnt
  CLOSE
RETURN

END SUB

