'GLBANK.BAS

DEFINT A-Z

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'PRO7.BI'                      'QuickPak Pro Routines
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

'Subs in this module
DECLARE SUB AddBank ()
DECLARE SUB OpenBankFile (BankRecLen%, BankFile%, NumBanks%)
DECLARE SUB EditBank ()
DECLARE SUB PrintBankList ()
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)

CONST False = 0, True = NOT False

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info

REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices
REDIM MChoice$(1 TO 4)

MChoice$(1) = "Add a New Bank Record"
MChoice$(2) = "Edit a Bank Record"
MChoice$(3) = "Print Bank List"
MChoice$(4) = "Quit"

MaxLen = 0                              'Set menu width to zero
BoxBot = 20                               'limit the box to line 20
Action = 0                                'Stay on menu til a choice is made
Choice = 1                                'Pre-load choice to highlight

FOR cnt = 1 TO UBOUND(MChoice$)           'Find max menu width
   TLen = LEN(MChoice$(cnt))
   IF TLen > MaxLen THEN
      MaxLen = TLen
   END IF
NEXT

Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1 'Center Menu within Screen
Col = ((80 - MaxLen) \ 2) - 2

TextCursor -2, -2                         'Define the Mouse cursor
ShowCursor                                'Show it

DO

   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode

   'TitleBox 3, Col, MaxLen + 3, "Bank Maintenance Menu ", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   Title$ = "Bank Maintenance Menu "
   MnuTitle Title$, Row, Col, MaxLen, 15
   
   PrintTitle User$
   PrintHelp "Bank Maintenance"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
         ClrBG
         AddBank
      CASE 2
         ClrBG
         EditBank
      CASE 3
         ClrBG
         PrintBankList
      CASE 4
         CLS : END
   END SELECT


LOOP


ClrBG
HideCursor                                'Be nice and turn off mouse cursor
DispLoad
RUN "glmaint"

SUB AddBank
   
  DIM BankRec AS BankRecType
  BankRecLen = LEN(BankRec)
'  OpenBankFile BankRecLen, BankFile, NumBanks
'  RecNum = NumBanks + 1
'  CLOSE
  
  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo              'Field info Array
  DIM Frm AS FormInfo                    'Form editing info
  LibName$ = "GL.QSL"                    'Screen library
  FormName$ = "ADDBANK"
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  QPrintRC "Add a New Bank", 25, 2, -1
  Action = 1
  Form$(1, 0) = STR$(RecNum)

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key
        OpenBankFile BankRecLen, BankFile, NumBanks
        RecNum = NumBanks + 1
        BankRec.Deleted = 0
        BankRec.BankNum = RecNum
        BankRec.BankName = Form$(2, 0)
        BankRec.BankAcct = Form$(3, 0)
        BankRec.GLAcct = Form$(4, 0)
        BankRec.Pad = ""
        PUT BankFile, RecNum, BankRec
        CLOSE BankFile
        EXIT DO
      CASE EscKey
        EXIT DO
    END SELECT
  LOOP UNTIL Frm.KeyCode = EscKey

END SUB

SUB EditBank

   DIM BankRec AS BankRecType
   BankRecLen = LEN(BankRec)
   REDIM Form$(0, 0)                      'Form array holds data while editing
   REDIM Fld(0) AS FieldInfo              'Field info Array
   DIM Frm AS FormInfo                    'Form editing info
   LibName$ = "GL.QSL"                    'Screen library
   FormName$ = "BANKSETE"
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   QPrintRC "Edit Bank Record", 25, 2, -1
   GoodRec = False
   Action = 1

   DO
      EditForm Form$(), Fld(), Frm, Cnf, Action

      IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field
        IF Frm.FldNo = 1 THEN            'Blank the form if back on acct fld
           Form$(1, 0) = ""
           Form$(2, 0) = ""
           Form$(3, 0) = ""
           Form$(4, 0) = ""
           GoodRec = False
        END IF

        SELECT CASE Frm.PrevFld          'user is moving around in form
          CASE 1
            IF Frm.FldEdited THEN      'If the field was edited
              RecNum = VAL(Form$(1, 0))
              IF RecNum > 0 THEN
                GOSUB GetBank
              END IF
            END IF
          CASE ELSE
        END SELECT
        Action = 1
      END IF

      SELECT CASE Frm.KeyCode             'Check the user's key presses
        CASE F10Key                      'SaveButton
          IF GoodRec THEN
            GOSUB SaveRec
          ELSE
            BEEP
            ok = MsgBox("GL.QSL", "BADREC")
            Frm.FldNo = 1
          END IF
        CASE F3Key                       'Delete Button
          IF GoodRec THEN         'make sure we have a record
             Killed = False
             GOSUB DeleteRec
             IF Killed = True THEN
                GOSUB ClearForm
             END IF
          ELSE
            BEEP
          END IF
        CASE ELSE
      END SELECT

   LOOP UNTIL Frm.KeyCode = EscKey


EXIT SUB

GetBank:
  OpenBankFile BankRecLen, BankFile, NumBanks
  IF RecNum > 0 AND RecNum <= NumBanks THEN
    GET BankFile, RecNum, BankRec
    IF BankRec.Deleted = 0 THEN
      Form$(2, 0) = BankRec.BankName
      Form$(3, 0) = BankRec.BankAcct
      Form$(4, 0) = BankRec.GLAcct
      CLOSE BankFile
      GoodRec = True
    ELSE
      ok = MsgBox("GL.QSL", "DELREC")
      Frm.FldNo = 1
      Form$(1, 0) = ""
      GoodRec = False
    END IF
  ELSE
    ok = MsgBox("GL.QSL", "BADREC")
    Frm.FldNo = 1
    Form$(1, 0) = ""
    GoodRec = False
    Action = 1
  END IF
RETURN


SaveRec:
  FOR FCnt = 1 TO 4
    IF LEN(QPTrim$(Form$(FCnt, 0))) = 0 THEN
      BadFld = FCnt
      EXIT FOR
    END IF
  NEXT
  IF BadFld = 0 THEN
    LOCATE 1, 1, 0
    PrintHelp "Updating files. Please wait."
    OpenBankFile BankRecLen, BankFileNum, NumBanks
    BankRec.BankName = Form$(2, 0)
    BankRec.BankAcct = Form$(3, 0)
    BankRec.GLAcct = Form$(4, 0)
    PUT BankFile, RecNum, BankRec
    CLOSE
    GOSUB ClearForm
  ELSE
    ok = MsgBox("GL.QSL", "NOBLANK")
    Frm.FldNo = BadFld
    Action = 1
    BadFld = 0
  END IF
RETURN



DeleteRec:                             'delete the current record
   LOCATE 1, 1, 0
   IF MsgBox("GL.QSL", "OK2DEL") = 2 THEN  'Make sure (2=yes)
      OpenBankFile BankRecLen, BankFile, NumBanks
      GET BankFile, RecNum, BankRec
      BankRec.Deleted = -1
      PUT BankFile, RecNum, BankRec
      CLOSE
      Killed = True
   END IF
RETURN


ClearForm:                               'reset form for more editing
   LOCATE 1, 1, 0
   FOR cnt = 1 TO 4
      Form$(cnt, 0) = ""
   NEXT
   PrintHelp "Bank File Maintenance"
   Frm.FldNo = 1
   Action = 1

RETURN


END SUB

SUB OpenBankFile (BankRecLen, BankFile, NumBanks)
   
   BankFile = FREEFILE
   OPEN "GLBANK.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS BankFile LEN = BankRecLen
   NumBanks = LOF(BankFile) \ BankRecLen
   
END SUB

SUB PrintBankList

   LOCATE , , 0
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo
   REDIM Choice$(0 TO 3, 1)
   DIM Frm AS FormInfo
   LibName$ = "GL.QSL"
   FormName$ = "PRNSPEC"

   'Choices for Vendor Status field
   Choice$(0, 0) = "1"
   Choice$(1, 0) = "Screen"
   Choice$(2, 0) = "Prn#1"
   Choice$(3, 0) = "Prn#2"
   
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   PrintHelp "Print Account Listing"
   Action = 1
   DO                                      'Allow user to edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode             'Check the user's key presses

         CASE F10Key                      'SaveButton
            Lookfor$ = RTRIM$(Form$(1, 0))
            IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN
               Dev$ = "S"
               LPTNo = 1
            ELSE
               Dev$ = LEFT$(Form$(2, 0), 1)
               LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
            END IF

         CASE EscKey
            EXIT SUB

      END SELECT

   LOOP UNTIL Frm.KeyCode = F10Key
   
   LOCATE , , 0

   '--Define vars used for printing
   Header$ = "Master Bank Listing"
   FF$ = CHR$(12)
   MaxLines = 55

   '--Background screen
   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
   PrintHelp "Processing report. Please wait."

   DIM BankRec AS BankRecType
   BankRecLen = LEN(BankRec)
   OpenBankFile BankRecLen, BankFile, NumBanks
   
   PRNFile = FREEFILE
   ReportFile$ = "BANKLIST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile
   
   GOSUB PrintPageHeader
   
   FOR cnt = 1 TO NumBanks
      GET BankFile, cnt, BankRec
      IF BankRec.Deleted = 0 THEN
        ToPrint$ = SPACE$(80)
        MID$(ToPrint$, 2) = STR$(BankRec.BankNum)
        MID$(ToPrint$, 6) = BankRec.BankName
        MID$(ToPrint$, 38) = BankRec.GLAcct
        PRINT #PRNFile, ToPrint$
        LineCnt = LineCnt + 1
        IF LineCnt > MaxLines THEN
          PRINT #PRNFile, FF$
          GOSUB PrintPageHeader
        END IF
      END IF
   NEXT

   PRINT #PRNFile, FF$
   CLOSE
   
   SELECT CASE Dev$
      CASE "S"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
      CASE "P"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   END SELECT

   'KILL ReportFile$
   ShowCursor

EXIT SUB

PrintPageHeader:
   PRINT #PRNFile, Header$
   PRINT #PRNFile,
   PRINT #PRNFile, "No.  Bank Name                    G/L Acct Number"
   PRINT #PRNFile, STRING$(78, "-")
   LineCnt = 4
RETURN


END SUB

