'GLFUND.BAS

DEFINT A-Z

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'PRO7.BI'                      'QuickPak Pro Routines
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'RPTPRN.BI'                    'Report printer routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

'Subs in this module
DECLARE SUB FundMaint ()
DECLARE SUB PrintFundList ()
DECLARE SUB ReSortFundIndex ()
DECLARE SUB SortFundIndex ()
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS Config)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)

DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()

DIM SHARED User$
DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Fund AS GLFundRecType
DIM SHARED FundIdx AS GLFundIndexType

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes

REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices
REDIM MChoice$(1 TO 4)
MChoice$(1) = "Add/Change/Delete Funds"
MChoice$(2) = "Print Fund Listing"
MChoice$(3) = "Sort Fund Index"
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

   'TitleBox 3, Col, MaxLen + 3, "Fund Maintenance Menu ", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

   Title$ = "Fund Maintenance Menu "
   MnuTitle Title$, Row, Col, MaxLen, 15
   PrintTitle User$
   PrintHelp "G/L Fund Maintenance"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
         ClrBG
         FundMaint
      CASE 2
         ClrBG
         PrintFundList
      CASE 3
         ClrBG
         ReSortFundIndex
      CASE 4
         CLS : END
      CASE ELSE
   END SELECT

LOOP

ClrBG
HideCursor                                'Be nice and turn off mouse cursor
DispLoad
RUN "glmaint"

SUB FundMaint

   LOCATE , , 0

   'SHARED Cnf AS Config
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo

   DIM Frm AS FormInfo

   LibName$ = "GL.QSL"                    'Display the form
   FormName$ = "EDFUND"
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   
   Action = 1 ':LOCATE 1, 1: PRINT FRE(-2)
   
   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

      IF SaveFlag > 0 THEN                'see CheckSave sub
         Frm.FldNo = SaveFlag
         SaveFlag = 0
      END IF

      IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field

         'Blank FundNum field if coming back in
         IF Frm.FldNo = 1 THEN
            Form$(1, 0) = ""
            Form$(2, 0) = ""
            EMode = 0
         END IF

         SELECT CASE Frm.PrevFld

            CASE 1                        'Fund Field

               IF Frm.FldEdited THEN      'If the field was edited
                  Form$(1, 0) = QPTrim$(Form$(1, 0))
                  SaveField 1, Form$(), Fld(), BadFld
                  IF LEN(Form$(1, 0)) <> FundLen THEN
                     Form$(1, 0) = ""
                     Frm.FldNo = 1
                  ELSE
                     RecordNum = FindFund(Form$(1, 0))
                     IF RecordNum > 0 THEN
                        GOSUB GetFund
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
            GOSUB DelFund

         CASE ESC
            'Check to see if form was edited before ESC was pressed

      END SELECT
      
   LOOP UNTIL Frm.KeyCode = ESCKey

   ERASE Form$
   ERASE Fld
   ERASE Choice$

EXIT SUB


GetFund:
   OpenFundFile FundFileNum, NumFunds
   GET FundFileNum, RecordNum, Fund
   Form$(2, 0) = Fund.Title
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
   OpenFundFile FundFileNum, NumFunds

   SELECT CASE EMode
   CASE 0
      Rec = NumFunds + 1
   CASE 1
      Rec = RecordNum
   END SELECT

   Fund.Deleted = 0
   Fund.FundNum = Form$(1, 0)
   Fund.Title = Form$(2, 0)

   PUT FundFileNum, Rec, Fund
   CLOSE

   SortFundIndex

   FOR Cnt = 1 TO 2
      Form$(Cnt, 0) = ""
   NEXT
   
   Frm.FldNo = 1
   Action = 1
   'EMode = 0

RETURN


DelFund:
   
   FOR Cnt = 1 TO 2
      Form$(Cnt, 0) = ""
   NEXT

   IF RecordNum = 0 THEN RETURN

   OpenFundFile FundFileNum, NumFunds
   Fund.Deleted = -1
   PUT FundFileNum, RecordNum, Fund
   CLOSE
   
   SortFundIndex
   
   Fund.Deleted = 0
   Frm.FldNo = 1
   Action = 1
   EMode = 0

RETURN


END SUB

SUB PrintFundList
  
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
   PrintHelp "Print Fund Listing"

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
               LPTNo = VAL(RIGHT$(RTRIM$(Form$(1, 0)), 1))
            END IF

         CASE ESCKey
            EXIT SUB

      END SELECT

   LOOP UNTIL Frm.KeyCode = F10Key
   
   LOCATE , , 0

   'Define vars used for printing
   MaxLines = 55
   FF$ = CHR$(12)
   Header$ = "Master Fund Listing"

   LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
   PrintHelp "Processing report. Please wait."

   OpenFundIdx FundIdxFileNum, NumFIdxRecs
   OpenFundFile FundFileNum, NumFunds
   PRNFile = FREEFILE
   ReportFile$ = "FUNDLIST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile
      
   GOSUB PrintFundPageHeader

   FOR Cnt = 1 TO NumFIdxRecs

      GET FundIdxFileNum, Cnt, FundIdx
      GET FundFileNum, FundIdx.RecNum, Fund

      HowMany = HowMany + 1

      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 2) = Fund.FundNum
      MID$(ToPrint$, 18) = Fund.Title
      PRINT #PRNFile, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PRNFile, FF$
        GOSUB PrintFundPageHeader
      END IF

   NEXT
   PRINT #PRNFile,
   PRINT #PRNFile, HowMany; "Funds listed."
   PRINT #PRNFile, FF$

   CLOSE
   
   SELECT CASE Dev$
      CASE "S"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
      CASE "P"
         'LibFile2Scrn "GL.QSL", "BAKGRD", MonoCode, Attribute, ErrorCode
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   END SELECT

   KILL ReportFile$
   

EXIT SUB


PrintFundPageHeader:
  PRINT #PRNFile, Header$
  PRINT #PRNFile,
  PRINT #PRNFile, " Fund Number     Title"
  PRINT #PRNFile, STRING$(80, "-")
  LineCnt = 4
RETURN


END SUB

SUB ReSortFundIndex

LibFile2Scrn "GL.QSL", "SORT", MonoCode, Attribute, ErrorCode
PrintHelp "Re-sort fund index."

LOCATE 1, 1, 0
AnyKey = MGetKey(Row, Col)

SELECT CASE AnyKey
   CASE 13, 1000
      QPrintRC SPACE$(40), 12, 21, 7
      QPrintRC "Sorting index.  Please wait.", 12, 21, 15
      SortFundIndex
   CASE 27, 1001
      EXIT SUB
END SELECT

END SUB

