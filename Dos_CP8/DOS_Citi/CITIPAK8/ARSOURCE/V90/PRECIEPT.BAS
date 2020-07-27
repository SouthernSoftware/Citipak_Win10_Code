DECLARE SUB PrintHelp (Help$)
DECLARE SUB PrintTitle (Title$)
DEFINT A-Z



Acct$ = Form$(3, 0)
CustName$ = Form$(4, 0)
Address$ = Form$(5, 0)
Desc$ = Form$(6, 0)
TotalOwed# = Value(Form$(7, 0), ecode%)
TotalCash# = Value(Form$(9, 0), ecode%)
TotalCheck# = Value(Form$(10, 0), ecode%)
TotalPaid# = Value(Form$(11, 0), ecode%)
TotalChange# = Value(Form$(12, 0), ecode%)
SetLicense$ = Form$(13, 0)
SetFee$ = Form$(14, 0)
GOSUB PrintReceipt



PrintReceipt:
  Help$ = "Citipak Cash Management System v8.0"
  ShowCursor
  ScrnName$ = "CMPOST"

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

  DisplayCMScrn "CMPOST"

  PrintTitle User$
  PrintHelp Help$

  Form$(1, 0) = "Y"
  Action = 1



PrtReceipt:
  REDIM TmpScrn(0)
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "Y" THEN
        SaveScrn TmpScrn()
        GOSUB PrintReceiptFunction
        RestScrn TmpScrn()
        Form$(1, 0) = "N"
        Frm(1).FldNo = 1
        Action = 1
        GOTO PrtReceipt
      END IF
      Posted = True
      Done = True
    CASE ESC
      Posted = False
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  RETURN

PrintReceiptFunction:
  REDIM CMTRRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTRRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  RecNumber = (LOF(CHandle) \ CMTrRecLen) + 1
  CLOSE CHandle

  'Set Printer Port for Receipts Printer if none set to 1

  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))

  OPEN "O", 15, "RECPT.PRN"
  WIDTH #15, 255

  PRINT #15, ""
  PRINT #15, TownName$
  PRINT #15, "LICENSE PAYMENT"
  PRINT #15, "Date: "; PostDate$
  PRINT #15,
  PRINT #15, "Account #"; Acct$
  PRINT #15, CustName$
  PRINT #15, Address$
  PRINT #15, Desc$
  PRINT #15,
  PRINT #15, "Total Owed: "; USING "$$####,#.##"; TotalOwed#
  PRINT #15, ""
  PRINT #15, "  Cash Amt: "; USING "$$####,#.##"; TotalCash#
  PRINT #15, " Check Amt: "; USING "$$####,#.##"; TotalCheck#
  PRINT #15, "             -----------"
  PRINT #15, "Total Paid: "; USING "$$####,#.##"; TotalPaid#
  PRINT #15, ""
  PRINT #15, "    Change: "; USING "$$####,#.##"; TotalChange#
  PRINT #15,
  PRINT #15,
  PRINT #15, "Operator: "; OperRecNum%
  PRINT #15, "Receipt#: "; USING "######"; RecNumber
  PRINT #15,
  PRINT #15, TAB(7); "T H A N K   Y O U !"
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15,
  PRINT #15, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  CLOSE
  PrintRptFile Header$, "RECPT.PRN", GetDEFPort%, RetCode%, 5
RETURN

