DECLARE SUB DisplayARScrn (ScrnName$)
DEFINT A-Z
DECLARE SUB CursorOff ()
DECLARE SUB ClearBack ()
DECLARE SUB CustMailLbl ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB WaitForAction ()

'$INCLUDE: 'DefCnf.BI'
'DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)

  CONST False = 0, True = NOT False

  TYPE Struct
   who AS STRING * 14
   RecNum AS INTEGER
  END TYPE

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARv85.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'

   DIM SHARED ARCust(1) AS ARCustRecType
   DIM SHARED ARCustRec(1) AS ARCustRecType
   DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType

SUB CustMailLbl

  SHARED Choice$()

  REDIM Choice$(0 TO 15, 0 TO 2)
  REDIM ToPrint(1 TO 5) AS STRING * 132
  FOR Cnt = 1 TO 5
    LSET ToPrint(Cnt) = ""
  NEXT
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = " Customer Name Order   "
  Choice$(2, 0) = " Account Number Order  "
  
  Choice$(0, 2) = "2"
  Choice$(1, 2) = "1)  1 x 3" + CHR$(171) + "  1 Wide"
  Choice$(2, 2) = "2)  1 x 3" + CHR$(171) + "  3 Wide"
  Choice$(3, 2) = "3)  1 x 3" + CHR$(171) + "  4 Wide"

  REDIM OSet(1 TO 4) AS INTEGER

  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110

  UsingAcct = False
  UsingName = False

  AbortFlag = False
  PageNo = 0

  LibName$ = "AR"
  ScrnName$ = "ARPRNLAB"
  
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  DisplayARScrn ScrnName$

  PrintTitle ""
  PrintHelp "Print Customer Mailing Labels"

  ShowCursor
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      LType = VAL(LEFT$(Form$(2, 0), 1))
      IF LType > 0 THEN
        GOSUB DoAlignLabelMask
      ELSE
        ErrCode = 2
        GOSUB ShowErrScrn1
        GOTO BadParm
      END IF
    CASE F10Key
      Order$ = LEFT$(QPTrim$(Form$(1, 0)), 3)
      SELECT CASE Order$
       CASE "Cus"
        IndexName$ = "ARCUST.IDX"
        UsingName = True
        OKFlag = True
      CASE "Acc"
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      CASE ELSE
        ErrCode = 1
        GOSUB ShowErrScrn1
        GOTO BadParm
      END SELECT

      LType = VAL(LEFT$(Form$(2, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB ShowErrScrn1
        OKFlag = False
        GOTO BadParm
      END IF
    CASE ESC
      GOTO ExitMailLabListing
    CASE ELSE
      Done = False
    END SELECT

BadParm:
  LOOP UNTIL OKFlag

'***************
  
  ERASE Frm, Form$, Fld, Choice$
  ClearBack
  ShowProcessingScrn "Creating Mailing Labels"

  OpenARCustFile NumOfArRecs, ARFile
  ARCustRecLen = LEN(ARCustRec(1))
  CLOSE ARFile

  IF UsingName THEN
    OpenARCustIdxFile NumOfARIdxRecs, ARIdxFile
    CLOSE ARIdxFile
    IdxRecLen = LEN(ARCustIdxRec(1))
    REDIM IdxBuff(1 TO NumOfARIdxRecs) AS ARCustIDXRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, NumOfARIdxRecs      'load it
    NumOfRecs = NumOfARIdxRecs
  ELSE
    NumOfRecs = FileSize("ARCUST.DAT") \ ARCustRecLen
  END IF

  ARCust = FREEFILE
  OPEN "ARCust.DAT" FOR RANDOM SHARED AS ARCust LEN = ARCustRecLen

  UBRpt = FREEFILE
  OPEN "ARLABEL.RPT" FOR OUTPUT AS UBRpt
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName THEN
      AcctNumber& = IdxBuff(Cnt).IDXRECORD
    ELSE
      AcctNumber& = Cnt
    END IF

    GET ARCust, AcctNumber&, ARCustRec(1)

    IF UCASE$(ARCustRec(1).Deleted) = "Y" THEN
      GOTO NextLabel
    END IF
    PrintHelp ARCustRec(1).CustName
GoodCust:
    CustPCnt = CustPCnt + 1

    Zip$ = LEFT$(ARCustRec(1).ZIPCODE, 5) + "-" + MID$(ARCustRec(1).ZIPCODE, 6)
    Zip$ = QPTrim$(Zip$)

    Zip$ = ARCustRec(1).ZIPCODE
    SELECT CASE LType
    CASE 1
      PRINT #UBRpt, "Cust #" + STR$(AcctNumber&)
      PRINT #UBRpt, LEFT$(QPTrim$(ARCustRec(1).CustName), 23)
      PRINT #UBRpt, LEFT$(QPTrim$(ARCustRec(1).ADDRESS1), 23)
      IF LEN(QPTrim$(ARCustRec(1).ADDRESS2)) > 0 THEN
        PRINT #UBRpt, LEFT$(QPTrim$(ARCustRec(1).ADDRESS2), 23)
        PRINT #UBRpt, LEFT$(QPTrim$(ARCustRec(1).CITY), 18) + ", " + ARCustRec(1).STATE + " " + Zip$
      ELSE
        PRINT #UBRpt, LEFT$(QPTrim$(ARCustRec(1).CITY), 18) + ", " + ARCustRec(1).STATE + " " + Zip$
        PRINT #UBRpt,
      END IF
      PRINT #UBRpt,
      DidCnt = DidCnt + 1
    CASE 2
      LabelCnt = LabelCnt + 1
      MID$(ToPrint(1), OSet(LabelCnt)) = "Cust #" + STR$(AcctNumber&)
      MID$(ToPrint(2), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).CustName), 23)
      MID$(ToPrint(3), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).ADDRESS1), 23)
      IF LEN(QPTrim$(ARCustRec(1).ADDRESS2)) > 0 THEN
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).ADDRESS2), 23)
        MID$(ToPrint(5), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).CITY), 18) + ", " + ARCustRec(1).STATE + " " + Zip$
      ELSE
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).CITY), 18) + ", " + ARCustRec(1).STATE + " " + Zip$
      END IF
      IF LabelCnt = 3 THEN
        FOR PCnt = 1 TO 5
          PRINT #UBRpt, ToPrint(PCnt)
          LSET ToPrint(PCnt) = ""
        NEXT
        PRINT #UBRpt,
        LabelCnt = 0
      END IF

    CASE 3
      LabelCnt = LabelCnt + 1
      MID$(ToPrint(1), OSet(LabelCnt)) = "Cust #" + STR$(AcctNumber&)
      MID$(ToPrint(2), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).CustName), 23)
      MID$(ToPrint(3), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).ADDRESS1), 23)
      IF LEN(QPTrim$(ARCustRec(1).ADDRESS2)) > 0 THEN
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).ADDRESS2), 23)
        MID$(ToPrint(5), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).CITY), 18) + ", " + ARCustRec(1).STATE + " " + Zip$
      ELSE
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(ARCustRec(1).CITY), 18) + ", " + ARCustRec(1).STATE + " " + Zip$
      END IF
      IF LabelCnt = 4 THEN
        FOR PCnt = 1 TO 5
          PRINT #UBRpt, ToPrint(PCnt)
          LSET ToPrint(PCnt) = ""
        NEXT
        PRINT #UBRpt,
        LabelCnt = 0
      END IF
    END SELECT

NextLabel:
    ShowPctComp Cnt, NumOfRecs
  NEXT

  IF LType = 2 OR LType = 3 THEN
    IF LabelCnt > 0 THEN
      FOR PCnt = 1 TO 5
        PRINT #UBRpt, QPTrim$(ToPrint(PCnt))
      NEXT
      PRINT #UBRpt,
    END IF
  END IF
  PCnt = 0

  CLOSE

  ERASE ToPrint
  ERASE Frm, Form$, Fld, OSet

  IF NOT AbortFlag THEN
    PrintRptFile "Mailing Labels", "ARLABEL.RPT", 1, RetCode, EntryPoint
  END IF

ExitMailLabListing:

  EXIT SUB

DoAlignLabelMask:

  REDIM TempScrn(0)
  SaveScrn TempScrn()
  Align$ = STRING$(34, "X")
  UBRpt = FREEFILE
  OPEN "ARLABEL.RPT" FOR OUTPUT AS UBRpt
  SELECT CASE LType
  CASE 1
    FOR Cnt = 1 TO 5
      PRINT #UBRpt, Align$
    NEXT
    PRINT #UBRpt,
  CASE 2
    FOR Cnt = 1 TO 5
      PRINT #UBRpt, Align$; TAB(OSet(2)); Align$; TAB(OSet(3)); Align$
    NEXT
    PRINT #UBRpt,
  CASE 3
    FOR Cnt = 1 TO 5
      PRINT #UBRpt, Align$; TAB(OSet(2)); Align$; TAB(OSet(3)); Align$; TAB(OSet(4)); Align$
    NEXT
    PRINT #UBRpt,
  END SELECT

  CLOSE UBRpt

  PrintRptFile "Mailing Labels Alignment", "ARLABEL.RPT", 1, RetCode, EntryPoint
  RestScrn TempScrn()
  ERASE TempScrn
  Action = 1
RETURN

ShowErrScrn1:
  SELECT CASE ErrCode
  CASE 2
    EText$ = "  Invalid Label Type!"
    Frm(1).FldNo = 2
  CASE ELSE
    EText$ = "Invalid Printing Order!"
    Frm(1).FldNo = 1
  END SELECT
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  CursorOff
  WazzWind 13, 27, 19, 63, 78, 1, True
  QPrintRC EText$, 15, 34, 79
  QPrintRC "Press Any Key To Continue.", 17, 33, 79
  WaitForAction
  RestScrn TempScrn()
  ERASE TempScrn
  Action = 1
  ERASE TempScrn

  RETURN

END SUB

