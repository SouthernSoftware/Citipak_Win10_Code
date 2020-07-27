'GLACCT.BAS

DEFINT A-Z

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'PRO7.BI'                      'QuickPak Pro Routines
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'UBIFDAT.BI'                    'Report printer routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

'Subs in this module
'DECLARE SUB TitleBox (Row, Col, MaxLen, Text$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB EditUBIF (WhichOne)

DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes

REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

REDIM MChoice$(1 TO 5)

MChoice$(1) = "Utility Payments"
MChoice$(2) = "Utility Adjustments"
MChoice$(3) = "Utility Deposits Rec'd"
MChoice$(4) = "Utility Deposits Applied"
MChoice$(5) = "Quit"

MaxLen = 0                              'Set menu width to zero
BoxBot = 20                               'limit the box to line 20
Action = 0                                'Stay on menu til a choice is made
Choice = 1                                'Pre-load choice to highlight

FOR Cnt = 1 TO UBOUND(MChoice$)           'Find max menu width
   TLen = LEN(MChoice$(Cnt))
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

   'TitleBox 3, Col, MaxLen + 3, "Utility Billing Interface", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   PrintTitle User$
   PrintHelp "Utility - G/L Interface Setup"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
         EditUBIF 1
      CASE 2
         EditUBIF 2
      CASE 3
         EditUBIF 3
      CASE 4
         EditUBIF 4
      CASE 5
         CLS : END
      CASE ELSE
   END SELECT

LOOP

HideCursor                                'Be nice and turn off mouse cursor
DispLoad
RUN "glmaint"

SUB EditUBIF (WhichOne)

   SELECT CASE WhichOne
     CASE 1
       IFFileName$ = "UBIFPMT.DAT"
       WindowTitle$ = " Utility Payments "
     CASE 2
       IFFileName$ = "UBIFADJ.DAT"
       WindowTitle$ = " Utility Adjustments "
     CASE 3
       IFFileName$ = "UBIFDPR.DAT"
       WindowTitle$ = " Utility Deposits Rec'd "
     CASE 4
       IFFileName$ = "UBIFDPA.DAT"
       WindowTitle$ = " Utility Deposits Applied "
   END SELECT
   
   DIM UBIFRec AS UBIFDATRecType
   RecLen = LEN(UBIFRec)
   IFDataFile = FREEFILE
   OPEN IFFileName$ FOR RANDOM AS IFDataFile LEN = RecLen
   
   REDIM Form$(0, 0)                      'Form array holds data while editing
   REDIM Fld(0) AS FieldInfo              'Field info Array
   DIM Frm AS FormInfo                    'Form editing info

   LibName$ = "GL.QSL"                    'Screen library
   FormName$ = "UBIFDAT"

   '--Display the screen
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

   GOSUB PrintWindowTitle
   GOSUB GetRec

   Action = 1

   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

      IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field

         'SELECT CASE Frm.FldNo
         '  CASE 2 TO 21
         '    A$ = StripAcct$(Form$(Frm.FldNo, 0))
         '    LSET Form$(Frm.FldNo, 0) = A$
         '    Action = 1
         'END SELECT
 
         SELECT CASE Frm.PrevFld          'user is moving around in form
            CASE 2 TO 21
               IF Frm.FldEdited THEN      'If the field was edited
                  AcctField = Frm.PrevFld
                  RecordNum = FindAcct(Form$(AcctField, 0))
                  IF RecordNum = 0 THEN
                    Form$(AcctField, 0) = QPTrim$(Form$(AcctField, 0))   'trim the field &
                    IF LEN(Form$(AcctField, 0)) <> TotLen THEN 'make sure len is ok
                       'BEEP
                       Form$(AcctField, 0) = ""
                       Frm.FldNo = AcctField
                       Action = 1
                    ELSE                    'format & look up acct
                       Form$(AcctField, 0) = FmtAcct$(Form$(AcctField, 0), FundLen%, AcctLen%, DetLen%)
                       RecordNum = FindAcct(Form$(AcctField, 0))
                       IF RecordNum > 0 THEN
                          AccountDesc$ = " Acct Description: " + GetAcctTitle$(RecordNum)
                          QPrintRC AccountDesc$, 25, 1, -1
                       ELSE
                          BEEP
                          Form$(AcctField, 0) = ""
                          Frm.FldNo = AcctField
                          QPrintRC "Invalid Account Number", 25, 1, -1
                       END IF
                    END IF
                  END IF
               END IF
            CASE ELSE
         END SELECT
         Action = 1
      END IF

      SELECT CASE Frm.KeyCode
         CASE F10Key
            GOSUB SaveRec
         CASE EscKey
            EXIT DO
      END SELECT

   LOOP UNTIL Frm.KeyCode = EscKey

EXIT SUB

PrintWindowTitle:
  QPrintRC "[", 4, 7, 10
  QPrintRC WindowTitle$, 4, 8, 14
  QPrintRC "]", 4, 8 + LEN(WindowTitle$), 10
RETURN

GetRec:
  GET IFDataFile, 1, UBIFRec
  Form$(1, 0) = UBIFRec.ACTIVE
  Form$(2, 0) = UBIFRec.GLDR1
  Form$(3, 0) = UBIFRec.GLCR1
  Form$(4, 0) = UBIFRec.GLDR2
  Form$(5, 0) = UBIFRec.GLCR2
  Form$(6, 0) = UBIFRec.GLDR3
  Form$(7, 0) = UBIFRec.GLCR3
  Form$(8, 0) = UBIFRec.GLDR4
  Form$(9, 0) = UBIFRec.GLCR4
  Form$(10, 0) = UBIFRec.GLDR5
  Form$(11, 0) = UBIFRec.GLCR5
  Form$(12, 0) = UBIFRec.GLDR6
  Form$(13, 0) = UBIFRec.GLCR6
  Form$(14, 0) = UBIFRec.GLDR7
  Form$(15, 0) = UBIFRec.GLCR7
  Form$(16, 0) = UBIFRec.GLDR8
  Form$(17, 0) = UBIFRec.GLCR8
  Form$(18, 0) = UBIFRec.GLDR9
  Form$(19, 0) = UBIFRec.GLCR9
  Form$(20, 0) = UBIFRec.GLDR10
  Form$(21, 0) = UBIFRec.GLCR10
RETURN

SaveRec:
  UBIFRec.ACTIVE = Form$(1, 0)
  UBIFRec.GLDR1 = Form$(2, 0)
  UBIFRec.GLCR1 = Form$(3, 0)
  UBIFRec.GLDR2 = Form$(4, 0)
  UBIFRec.GLCR2 = Form$(5, 0)
  UBIFRec.GLDR3 = Form$(6, 0)
  UBIFRec.GLCR3 = Form$(7, 0)
  UBIFRec.GLDR4 = Form$(8, 0)
  UBIFRec.GLCR4 = Form$(9, 0)
  UBIFRec.GLDR5 = Form$(10, 0)
  UBIFRec.GLCR5 = Form$(11, 0)
  UBIFRec.GLDR6 = Form$(12, 0)
  UBIFRec.GLCR6 = Form$(13, 0)
  UBIFRec.GLDR7 = Form$(14, 0)
  UBIFRec.GLCR7 = Form$(15, 0)
  UBIFRec.GLDR8 = Form$(16, 0)
  UBIFRec.GLCR8 = Form$(17, 0)
  UBIFRec.GLDR9 = Form$(18, 0)
  UBIFRec.GLCR9 = Form$(19, 0)
  UBIFRec.GLDR10 = Form$(20, 0)
  UBIFRec.GLCR10 = Form$(21, 0)
  PUT IFDataFile, 1, UBIFRec
  EXIT SUB
RETURN

END SUB

SUB PrintAcctList

   LOCATE , , 0
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo
   REDIM Choice$(0 TO 3, 1)
   DIM Frm AS FormInfo
   LibName$ = "GL.QSL"
   FormName$ = "PRNACCT"

   'Choices for Vendor Status field
   Choice$(0, 0) = "2"
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
   Header$ = "Master Account Listing"
   FF$ = CHR$(12)
   MaxLines = 55

   '--Background screen
   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
   PrintHelp "Processing report. Please wait."

   OpenAcctIdx AcctIdxFileNum, NumIdxRecs
   OpenAcctFile AcctFileNum, NumAccts
   PrnFile = FREEFILE
   ReportFile$ = "ACCTLIST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PrnFile
      
   GOSUB PrintAcctPageHeader

   FOR Cnt = 1 TO NumIdxRecs
      GET AcctIdxFileNum, Cnt, AcctIdx
      GET AcctFileNum, AcctIdx.RecNum, Acct
      FundCode$ = QPTrim$(LEFT$(AcctIdx.AcctNum, FundLen))
      IF INSTR(FundCode$, Lookfor$) THEN
         HowMany = HowMany + 1
         ToPrint$ = SPACE$(80)
         MID$(ToPrint$, 2) = Acct.Num
         MID$(ToPrint$, 18) = Acct.Title
         MID$(ToPrint$, 63) = Acct.Typ
         PRINT #PrnFile, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
           PRINT #PrnFile, FF$
           GOSUB PrintAcctPageHeader
         END IF
      END IF

   NEXT
   PRINT #PrnFile,
   PRINT #PrnFile, HowMany; "Accounts listed."

   CLOSE
   
   SELECT CASE Dev$
      CASE "S"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
      CASE "P"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   END SELECT

   KILL ReportFile$
   ShowCursor

EXIT SUB

PrintAcctPageHeader:
   PRINT #PrnFile, Header$
   PRINT #PrnFile,
   PRINT #PrnFile, "Acct Number     Title                                       Type"
   PRINT #PrnFile, STRING$(78, "-")
   LineCnt = 4
RETURN


END SUB

