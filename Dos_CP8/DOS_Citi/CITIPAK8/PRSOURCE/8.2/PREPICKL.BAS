DEFINT A-Z
DECLARE SUB PCEnterEdTrans (TransRecNo%, EmpRecNo%)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB WazzWind (BYVAL TopRow, BYVAL LeftCol, BYVAL BotRow, BYVAL RghtCol, BYVAL FrameColor, BYVAL FrameType, BYVAL Shadow)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION ASCII% (Strng$)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ReplaceString (Work$, Old$, New$)
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB ShowCursor ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode, Attribute, ErrorCode)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB CursorOff ()

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.bi'
''$INCLUDE: 'PageInfo.bi'    '        "          Page information
''$INCLUDE: 'Retire.Bi'
''$INCLUDE: 'DedCodes.Bi'
''$INCLUDE: 'ErnCodes.Bi'
''$INCLUDE: 'Leave.Bi'
''$INCLUDE: 'StateTax.Bi'
''$INCLUDE: 'FedTax.Bi'
''$INCLUDE: 'PREIC.Bi'
''$INCLUDE: 'PRUNIT.Bi'
''$INCLUDE: 'PRSYSCTR.Bi'
''$INCLUDE: 'PRPPDEF.Bi'
'$INCLUDE: 'PRTRANS.Bi'
''$INCLUDE: 'PRTIMEIN.Bi'

  CONST False = 0
  CONST True = NOT False

  TYPE FLen
     V AS STRING * 43
  END TYPE

SUB PCPickEmpList (TransRecNo%, EmpRecNo%)

  STATIC Choice

  CursorOff

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  REDIM Emp2DataRec(1)     AS EmpData2Type
  REDIM TempTransRec(1)    AS TransRecType
  REDIM TempData(1)        AS STRING * 43

  Data2RecLen = LEN(Emp2DataRec(1))
  TransRecLen = LEN(TempTransRec(1))
  NumOfRecs = FileSize(EmpIdxLName) \ 2

  REDIM Items(1 TO NumOfRecs) AS FLen

  MatchCnt = 0

  FOpenS EmpData2Name, D2Handle
  FOpenS TransWorkFileName, WHandle             'open transaction history file

  FOR Cnt = 1 TO NumOfRecs
    FGetRTA D2Handle, Emp2DataRec(1), CLNG(Cnt), Data2RecLen
    IF NOT Emp2DataRec(1).Deleted = True THEN        'put deleted test here
      IF Emp2DataRec(1).EMPTDATE = -32767 THEN
        FGetRTA WHandle, TempTransRec(1), CLNG(Cnt), TransRecLen
        MatchCnt = MatchCnt + 1
        LSET Items(MatchCnt).V = RIGHT$(Emp2DataRec(1).EmpNo, 6)
        MID$(Items(MatchCnt).V, 9, 15) = QPTrim$(Emp2DataRec(1).EMPLNAME)
        MID$(Items(MatchCnt).V, 25, 10) = QPTrim$(Emp2DataRec(1).EMPFNAME)
        IF TempTransRec(1).TActive THEN
          MID$(Items(MatchCnt).V, 39) = "Y"
        ELSE
          MID$(Items(MatchCnt).V, 39) = "N"
        END IF
        MID$(Items(MatchCnt).V, 42) = MKI$(Cnt)
      END IF
    END IF
  NEXT

  FClose D2Handle
  FClose WHandle             'open transaction history file

  IF MatchCnt > 1 THEN
    REDIM PRESERVE Items(1 TO MatchCnt) AS FLen
    SortT Items(1), MatchCnt, 0, LEN(Items(1).V), 0, 43
  END IF

  VAction = 1
  BoxBot = 17
  BoxTop = 7
  MaxLen = 40
  LeftCol = ((80 - MaxLen) \ 2) - 1

  DO
    HideCursor
    LOCATE BoxTop, LeftCol
    BlockClear
    LibFile2Scrn CalcQLib, "PICKPREV", MonoCode, -2, ErrCode
    GOSUB DisplayPrevName

    QPrintRC " Number   Last Name       First       Trans ", BoxTop - 1, LeftCol, Cnf.ActivCH

    MPaintBox BoxTop, LeftCol + MaxLen + 4, BoxTop, LeftCol + MaxLen + 5, 8
    WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.MenBox, 2, True
    QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.ActivCH

    ShowCursor
'   CursorOff

    DO

    VertMenuT Items(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf

    SELECT CASE ASCII(Ky$)
      CASE EscKey
        TransRecNo = 0
        Choice = EscKey
        ExitFlag = True
        EXIT DO
      CASE 13
        EmpRecNo = CVI(RIGHT$(Items(Choice).V, 2))
        TransRecNo = EmpRecNo
        PCEnterEdTrans TransRecNo, EmpRecNo
        FOpenS TransWorkFileName, WHandle    'open transaction  file
        FGetRTA WHandle, TempTransRec(1), CLNG(TransRecNo), TransRecLen
        FClose WHandle             'close transaction file
        IF TempTransRec(1).TActive THEN
          MID$(Items(Choice).V, 39) = "Y"
        ELSE
          MID$(Items(Choice).V, 39) = "N"
        END IF
        EXIT DO
      CASE ELSE
    END SELECT
    LOOP

  LOOP UNTIL ExitFlag

  CursorOff

  RestScrn TempScrn()

  ERASE Items, TempScrn, TempData, Emp2DataRec, TempTransRec
EXIT SUB

DisplayPrevName:
  HideCursor
  IF EntryType = Normal THEN
    QPrintRC "Payroll", 1, 3, 112
  ELSE
    QPrintRC "Manual", 1, 4, 112
  END IF
  IF LEN(EmpName$) > 0 THEN QPrintRC EmpName$, 1, 45, 112
  ShowCursor
  RETURN

END SUB

