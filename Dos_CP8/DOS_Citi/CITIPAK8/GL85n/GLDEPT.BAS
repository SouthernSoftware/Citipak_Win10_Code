DECLARE SUB ImportText ()
'GLDept.BAS
DEFINT A-Z

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'PRO7.BI'                      'QuickPak Pro Routines
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
''$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

DECLARE SUB DeptMaint ()
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB OpenDeptIdx (DeptIdxFile%, NumActiveDepts%)
DECLARE SUB PrintDeptList ()
DECLARE SUB SortDeptIndex ()
DECLARE SUB ReSortDeptIndex ()
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintHelp (Help$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)

DIM SHARED User$
DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Dept AS GLDeptRecType
DIM SHARED DeptIdx AS GLDeptIndexType

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes


'STOP
'ImportText


REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

REDIM MChoice$(1 TO 4)

MChoice$(1) = "Add/Change/Delete Depts "
MChoice$(2) = "Print Dept Listing"
MChoice$(3) = "Sort Dept Index"
MChoice$(4) = "Quit"

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

   'TitleBox 3, Col, MaxLen + 3, "Dept Maintenance Menu", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   
   Title$ = "Dept Maintenance Menu "
   MnuTitle Title$, Row, Col, MaxLen, 15
   PrintTitle User$
   PrintHelp "Department Maintenance"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
         ClrBG
         DeptMaint
      CASE 2
         ClrBG
         PrintDeptList
      CASE 3
         ClrBG
         ReSortDeptIndex
      CASE 4
         CLS : END
      CASE ELSE
   END SELECT

LOOP

ClrBG
HideCursor                                'Be nice and turn off mouse cursor
DispLoad
RUN "glmaint"

SUB DeptMaint

   'LOCATE , , 0

   SHARED Cnf AS Config
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo

   DIM Frm AS FormInfo

   LibName$ = "GL.QSL"                    'Display the form
   FormName$ = "EDDEPT"
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   
   Action = 1 ':LOCATE 1, 1: PRINT FRE(-2)
   
   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

      IF SaveFlag > 0 THEN                'see CheckSave sub
         Frm.FldNo = SaveFlag
         SaveFlag = 0
      END IF

      IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field

         'Blank DeptNum field if coming back in
         IF Frm.FldNo = 1 THEN
            Form$(1, 0) = ""
            Form$(2, 0) = ""
            EMode = 0
         END IF

         SELECT CASE Frm.PrevFld

            CASE 1                        'Dept Field

               IF Frm.FldEdited THEN      'If the field was edited
                  Form$(1, 0) = QPTrim$(Form$(1, 0))
                  SaveField 1, Form$(), Fld(), BadFld
                  IF LEN(Form$(1, 0)) <> AcctLen THEN
                     Form$(1, 0) = ""
                     Frm.FldNo = 1
                  ELSE
                     RecordNum = FindDept(Form$(1, 0))
                     IF RecordNum > 0 THEN
                        GOSUB GetDept
                        Frm.FldNo = 2
                        EMode = 1
                     END IF
                  END IF
               END IF

               CASE ELSE
                  'what else?
            END SELECT
         
         Action = 1

      END IF
      
      SELECT CASE Frm.KeyCode             'Check the user's key presses
     
         CASE F10Key                      'SaveButton
            SaveFlag = 0
            GOSUB CheckSave               'Check that no fields are blank
            IF SaveFlag = 0 THEN          'All ok. Save Record
               GOSUB Save
            END IF

         CASE F3Key
            GOSUB DelDept

         CASE ESC
            'Check to see if form was edited before ESC was pressed

      END SELECT
      
   LOOP UNTIL Frm.KeyCode = ESC

   ERASE Form$
   ERASE Fld
   ERASE Choice$

EXIT SUB


GetDept:
   OpenDeptFile DeptFileNum, NumDepts
   GET DeptFileNum, RecordNum, Dept
   Form$(2, 0) = Dept.Title
   CLOSE
   EMode = 1
RETURN


CheckSave:
   FOR Cnt = 1 TO 2                       '2 editable fields on form
      IF LEN(LTRIM$(RTRIM$(Form$(Cnt, 0)))) = 0 THEN
         ok = MsgBox("GL.QSL", "NOBLANK") 'No blank fields!
         SaveFlag = Cnt                   'Set save flag to blank field #
         EXIT FOR
      END IF
   NEXT
RETURN


Save:
   OpenDeptFile DeptFileNum, NumDepts

   SELECT CASE EMode
   CASE 0
      Rec = NumDepts + 1
   CASE 1
      Rec = RecordNum
   END SELECT

   Dept.Deleted = 0
   Dept.DeptNum = Form$(1, 0)
   Dept.Title = Form$(2, 0)

   PUT DeptFileNum, Rec, Dept
   CLOSE

   SortDeptIndex

   FOR Cnt = 1 TO 2
      Form$(Cnt, 0) = ""
   NEXT
   
   Frm.FldNo = 1
   Action = 1
   'EMode = 0

RETURN


DelDept:
   
   FOR Cnt = 1 TO 2
      Form$(Cnt, 0) = ""
   NEXT

   IF RecordNum = 0 THEN RETURN

   OpenDeptFile DeptFileNum, NumDepts
   Dept.Deleted = -1
   PUT DeptFileNum, RecordNum, Dept
   CLOSE
   
   SortDeptIndex
   
   Dept.Deleted = 0
   Frm.FldNo = 1
   Action = 1
   EMode = 0

RETURN


END SUB

SUB ImportText
  
  SHARED Dept AS GLDeptRecType

  ImportFileName$ = "GLD.ASC"
  ImportFile = FREEFILE
  OPEN ImportFileName$ FOR INPUT AS #ImportFile

  OpenDeptFile DeptFileNum, NumDepts

  CLS
  LOCATE 1, 1
  PRINT "Importing Department File"

  DO

    RecordsIn = RecordsIn + 1
    LOCATE 2, 1, 0
    PRINT "Processing Record: "; RecordsIn

    '--this imports a comma-delimited file
    INPUT #ImportFile, DeptNum$, DeptName$

    '--this imports a text file
    'LINE INPUT #ImportFile, LineIn$
    '--text file
    'LineLen = LEN(LineIn$)
    'Acct.Num = QPTrim$(LEFT$(LineIn$, 15))
    'Acct.Title = QPTrim(MID$(LineIn$, 18, 25))
    'Acct.Typ = MID$(LineIn$, 63, 1)

    Dept.Deleted = 0
    Dept.DeptNum = DeptNum$
    Dept.Title = DeptName$

    PUT DeptFileNum, RecordsIn, Dept

  LOOP UNTIL EOF(ImportFile)

  PRINT "File import complete."; RecordsIn; "records were imported."
  PRINT "Press any key to continue."
  k$ = INPUT$(1)

  CLOSE
  
  SortDeptIndex


END SUB

SUB PrintDeptList
   
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
   'Choice$(4, 0) = "Prn#3"
   
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   PrintHelp "Print Dept Listing"

   Action = 1

   DO                                      'Allow user to edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode             'Check the user's key presses

         CASE F10Key                      'SaveButton
            Lookfor$ = RTRIM$(Form$(1, 0))
            IF LEN(LTRIM$(RTRIM$(Form$(1, 0)))) = 0 THEN
               Dev$ = "S"
               LPTNo = 1
            ELSE
               Dev$ = LEFT$(Form$(1, 0), 1)
               LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
            END IF

         CASE ESC
            EXIT SUB

      END SELECT

   LOOP UNTIL Frm.KeyCode = F10Key
   
   LOCATE , , 0

   '--Define strings used for printing
   DIM Desc$(1)
   Header$ = "Master Dept Listing"
   MaxLines = 55
   FF$ = CHR$(12)

   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
   PrintHelp "Processing report. Please wait."

   OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
   OpenDeptFile DeptFileNum, NumDepts

   PRNFile = FREEFILE
   ReportFile$ = "DeptLIST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile
      
   GOSUB PrintDeptPageHeader

   FOR Cnt = 1 TO NumDIdxRecs
     GET DeptIdxFileNum, Cnt, DeptIdx
     GET DeptFileNum, DeptIdx.RecNum, Dept
     ToPrint$ = SPACE$(80)
     MID$(ToPrint$, 2) = Dept.DeptNum
     MID$(ToPrint$, 18) = Dept.Title
     PRINT #PRNFile, ToPrint$
     LineCnt = LineCnt + 1
     IF LineCnt > MaxLines THEN
       PRINT #PRNFile, FF$
       GOSUB PrintDeptPageHeader
     END IF

   NEXT

   PRINT #PRNFile,
   PRINT #PRNFile, NumDIdxRecs; "Depts listed."
   PRINT #PRNFile, FF$
   CLOSE
   
   SELECT CASE Dev$
      CASE "S"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
      CASE "P"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   END SELECT

   KILL ReportFile$
   
EXIT SUB

PrintDeptPageHeader:
  PRINT #PRNFile, Header$
  PRINT #PRNFile,
  PRINT #PRNFile, " Dept Number     Title"
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

SUB ReSortDeptIndex

LibFile2Scrn "GL.QSL", "SORT", MonoCode, Attribute, ErrorCode
PrintHelp "Re-sort Dept index."

LOCATE 1, 1, 0
AnyKey = MGetKey(Row, Col)

SELECT CASE AnyKey
   CASE 13, 1000
      QPrintRC SPACE$(40), 12, 21, 7
      QPrintRC "Sorting index.  Please wait.", 12, 21, 15
      SortDeptIndex
   CASE 27, 1001
      EXIT SUB
END SELECT

END SUB

SUB SortDeptIndex

   IF Exist("GLDept.IDX") THEN                'Delete index file if it exits
      KILL "GLDept.IDX"
   END IF

   OpenDeptIdx DeptIdxFileNum, NumADepts   'Open the Index file

   OpenDeptFile DeptFileNum, NumDepts         'Open the Dept file

   IF LOF(DeptFileNum) = 0 THEN               'get out if nothing to do
      CLOSE
      EXIT SUB
   END IF
   
   REDIM IdxBuff(1 TO NumDepts) AS GLDeptIndexType  'dim the buffer array

   FOR Cnt = 1 TO NumDepts                    'Load the buffer with Depts
      GET DeptFileNum, Cnt, Dept
      'Dept.Deleted = 0
      'PUT DeptFileNum, Cnt, Dept
      IF Dept.Deleted = 0 THEN                'Get only active accounts
         GoodDepts = GoodDepts + 1
         IdxBuff(GoodDepts).DeptNum = Dept.DeptNum
         IdxBuff(GoodDepts).RecNum = Cnt
      END IF
   NEXT

   CLOSE DeptFileNum                          'Close the Dept file

   IF GoodDepts = 0 THEN                      'Exit if all are deleted
      CLOSE DeptIdxFileNum
      EXIT SUB
   END IF

   REDIM PRESERVE IdxBuff(1 TO GoodDepts) AS GLDeptIndexType  'redim with
                                                              'just good Depts
   SortT IdxBuff(1), GoodDepts, 0, 10, 0, 8 'sort the $#&+ out of it!

   FOR Cnt = 1 TO GoodDepts                   'write out to the index file
      GET DeptIdxFileNum, Cnt, DeptIdx
      DeptIdx.DeptNum = IdxBuff(Cnt).DeptNum
      DeptIdx.RecNum = IdxBuff(Cnt).RecNum
      PUT DeptIdxFileNum, Cnt, DeptIdx
   NEXT

   CLOSE DeptIdxFileNum                       'close index
 

END SUB

