'for demo only limits system to 50 total accounts
DECLARE SUB MnuTitle (Title$, X%, Y%, Wdth%, TitColor%)
'GLACCT.BAS

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
DECLARE SUB AddAcct ()
DECLARE SUB EditAcct ()
DECLARE SUB ReSortAcctIndex ()
DECLARE SUB PrintAcctList ()
DECLARE SUB ImportText ()
DECLARE SUB TitleBox (Row, Col, MaxLen, Text$, Cnf AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortAcctIndex ()

DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes


REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

REDIM MChoice$(1 TO 4)

MChoice$(1) = "Add/Change/Delete G/L Accounts"
MChoice$(2) = "Print Account Listing"
MChoice$(3) = "Sort Account Index"
MChoice$(4) = "Quit"
'MChoice$(5) = "Import Text"

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

   'TitleBox 3, Col, MaxLen + 3, "Account Maintenance Menu ", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

   Title$ = "Account Maintenance Menu "
   MnuTitle Title$, Row, Col, MaxLen, 15
   PrintTitle User$
   PrintHelp "Chart of Accounts"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
         EditAcct
      CASE 2
         PrintAcctList
      CASE 3
         ReSortAcctIndex
      CASE 4
         CLS : END
      'CASE 5
      '  ImportText
      CASE ELSE
   END SELECT

LOOP

HideCursor                                'Be nice and turn off mouse cursor
DispLoad
RUN "glmaint"

SUB EditAcct

   ':LOCATE 2, 1: PRINT FRE(-2)
   
   'SHARED Cnf AS Config                   'Color & monitor info
   REDIM Form$(0, 0)                      'Form array holds data while editing
   REDIM Fld(0) AS FieldInfo              'Field info Array
   REDIM Choice$(0 TO 4, 1)               'for pick lists

   Choice$(0, 0) = "3"                    'choice for field #3 (Acct Type)
   Choice$(1, 0) = "Asset"
   Choice$(2, 0) = "Liability"
   Choice$(3, 0) = "Revenue"
   Choice$(4, 0) = "Expenditure"
   
   DIM Frm AS FormInfo                    'Form editing info

   LibName$ = "GL.QSL"                    'Screen library
   FormName$ = "EDACCT"

                                          'Display the screen
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

   Action = 1

   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

      IF SaveFlag > 0 THEN                'Flag for blank fields
         Frm.FldNo = SaveFlag             'see checksave
         SaveFlag = 0
      END IF

      IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field

         IF Frm.FldNo = 1 THEN            'Blank the form if back on acct fld
            Form$(1, 0) = ""
            Form$(2, 0) = ""
            Form$(3, 0) = ""
         END IF
 
         SELECT CASE Frm.PrevFld          'user is moving around in form
            CASE 1
               IF Frm.FldEdited THEN      'If the field was edited

                  Form$(1, 0) = QPTrim$(Form$(1, 0))   'trim the field &
                  IF LEN(Form$(1, 0)) <> TotLen THEN   'make sure len is ok
                     Form$(1, 0) = ""
                     Frm.FldNo = 1
                  
                  ELSE                    'format & look up acct
                     Form$(1, 0) = FmtAcct$(Form$(1, 0), FundLen%, AcctLen%, DetLen%)
                     RecordNum = FindAcct(Form$(1, 0))
                     IF RecordNum > 0 THEN
                        GOSUB GetAcct
                        Frm.FldNo = 2
                        EMode = 1         'account exits - set edit mode
                     ELSE
                        EMode = 0         'account doesn't exist
                     END IF

                  END IF

               END IF

            CASE ELSE

         END SELECT

         Action = 1

      END IF

      SELECT CASE Frm.KeyCode             'Check the user's key presses

         CASE F10Key                      'SaveButton
            SaveFlag = 0
            GOSUB CheckEdit               'Check that no fields are blank
            IF SaveFlag = 0 THEN          'All ok. Save Record
               GOSUB SaveRec
               GOSUB ClearForm
            END IF
            
         CASE F3Key                       'Delete Button
            Killed = 0
            IF RecordNum > 0 THEN         'make sure we have a record
               GOSUB DeleteRec
               IF Killed = -1 THEN
                  GOSUB ClearForm
               END IF
            END IF

         CASE EscKey
            EXIT DO
            'Check to see if form was edited before ESC was pressed

      END SELECT

   LOOP UNTIL Frm.KeyCode = EscKey


EXIT SUB

GetAcct:
   OpenAcctFile AcctFileNum, NumAccts
      GET AcctFileNum, RecordNum, Acct
      Form$(2, 0) = Acct.Title
      Form$(3, 0) = Acct.Typ
   CLOSE AcctFileNum
   
RETURN


SaveRec:
   LOCATE 1, 1, 0
   PrintHelp "Updating files. Please wait."
   OpenAcctFile AcctFileNum, NumAccts

   SELECT CASE EMode
      CASE 0                              'new record
         IF NumAccts < 50 THEN
           Rec = NumAccts + 1
           Acct.Deleted = 0
           Acct.Num = Form$(1, 0)
           Acct.Title = Form$(2, 0)
           Acct.Typ = Form$(3, 0)
           Acct.FrstTran = 0
           Acct.LastTran = 0
           Acct.PYAct = 0
           Acct.BegBal = 0
           Acct.Bgt = 0
           Acct.Bal = 0
           Acct.Encumb = 0
           Acct.MTD = 0
           Acct.YTD = 0
           Acct.NYEst = 0
           Acct.NYReq = 0
           Acct.NYRec = 0
           Acct.NYApp = 0
           Acct.Res = ""
         ELSE
           'ok = MsgBox("GL.QSL", "DEMO")
            CLS
            QPrintRC "Demo Version limited to 50 Total Accounts.", 12, 1, 11
            QPrintRC "Press any key to continue.", 13, 1, 11
            K$ = INPUT$(1)
            EXIT SUB
         END IF
      CASE 1                              'existing record
         Rec = RecordNum
         Acct.Title = Form$(2, 0)
         Acct.Typ = Form$(3, 0)

   END SELECT
   
   PUT AcctFileNum, Rec, Acct
   CLOSE
   SortAcctIndex

RETURN


CheckEdit:                                'Make sure there are no blank fields
   FOR Cnt = 1 TO 3
      IF LEN(QPTrim$(Form$(Cnt, 0))) = 0 THEN
         ok = MsgBox("GL.QSL", "NOBLANK")
         SaveFlag = Cnt
         EXIT FOR
      END IF

   NEXT
RETURN


DeleteRec:                             'delete the current record
   LOCATE 1, 1, 0
   IF MsgBox("GL.QSL", "OK2DEL") = 2 THEN  'Make sure (2=yes)
      IF Acct.LastTran = 0 THEN        'can't delete accout with trans
         PrintHelp "Updating files. Please wait."

         OpenAcctFile AcctFileNum, NumAccts

         'GET AcctFileNum, RecordNum, Acct
         Acct.Deleted = -1
         PUT AcctFileNum, RecordNum, Acct

         CLOSE AcctFileNum

         Killed = -1

         SortAcctIndex

         PrintHelp "G/L Account Maintenance"

      ELSE                                'Acct has transactions
         PLAY "L32O3EC"
         ok = MsgBox("GL.QSL", "CANTDEL")

      END IF
   END IF

RETURN


ClearForm:                               'reset form for more editing
   LOCATE 1, 1, 0
   FOR Cnt = 1 TO 3
      Form$(Cnt, 0) = ""
   NEXT

   PrintHelp "G/L Account Maintenance"
   Frm.FldNo = 1
   Action = 1

RETURN


END SUB

SUB ImportText

   SHARED TotLen%, FundLen%, AcctLen%, DetLen%
   SHARED Acct AS GLAcctRecType
   
   ImportFileName$ = "ACCTLIST.TXT"
   ImportFile = FREEFILE
   OPEN ImportFileName$ FOR INPUT AS #ImportFile

   OpenAcctFile AcctFileNum, NumAccts

   CLS
   LOCATE 1, 1
   PRINT "Importing Account File"

   DO

      RecordsIn = RecordsIn + 1
      LOCATE 2, 1, 0
      PRINT "Processing Record: "; RecordsIn

      '--this imports a comma-delimited file
      'INPUT #ImportFile, AcctNum$, AcctName$, AcctType$

      '--this imports a text file
      LINE INPUT #ImportFile, LineIn$
         LineLen = LEN(LineIn$)

         Acct.Deleted = 0
         Acct.Num = QPTrim$(LEFT$(LineIn$, 15))
         Acct.Title = QPTrim(MID$(LineIn$, 19, 35))
         Acct.Typ = MID$(LineIn$, 65, 1)
         Acct.FrstTran = 0
         Acct.LastTran = 0
         Acct.PYAct = 0
         Acct.BegBal = 0
         Acct.Bgt = 0
         Acct.Bal = 0
         Acct.Encumb = 0
         Acct.MTD = 0
         Acct.YTD = 0
         Acct.NYEst = 0
         Acct.NYReq = 0
         Acct.NYRec = 0
         Acct.NYApp = 0
         Acct.Res = ""
      
      PUT AcctFileNum, RecordsIn, Acct

   LOOP UNTIL EOF(ImportFile)

   PRINT "File import complete."; RecordsIn; "records were imported."
   PRINT "Press any key to continue."
   K$ = INPUT$(1)

   CLOSE
   
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
   PRNFile = FREEFILE
   ReportFile$ = "ACCTLIST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile
      
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
         PRINT #PRNFile, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
           PRINT #PRNFile, FF$
           GOSUB PrintAcctPageHeader
         END IF
      END IF

   NEXT

   PRINT #PRNFile,
   PRINT #PRNFile, HowMany; "Accounts listed."
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

PrintAcctPageHeader:
   PRINT #PRNFile, Header$
   PRINT #PRNFile,
   PRINT #PRNFile, "Acct Number     Title                                       Type"
   PRINT #PRNFile, STRING$(78, "-")
   LineCnt = 4
RETURN


END SUB

SUB ResetPntrs

   OpenAcctFile AcctFileNum, NumAccts
   FOR Cnt = 1 TO NumAccts
      GET AcctFileNum, Cnt, Acct
      Acct.FrstTran = 0
      Acct.LastTran = 0
      PUT AcctFileNum, Cnt, Acct
   NEXT

END SUB

SUB ReSortAcctIndex

LibFile2Scrn "GL.QSL", "SORT", MonoCode, Attribute, ErrorCode
PrintHelp "Re-sort account index."

LOCATE 1, 1, 0
AnyKey = MGetKey(Row, Col)

SELECT CASE AnyKey
   CASE 13, 1000
      QPrintRC SPACE$(40), 12, 21, 7
      QPrintRC "Sorting index.  Please wait.", 12, 21, 15
      SortAcctIndex
   CASE 27, 1001
      EXIT SUB
END SELECT

END SUB

