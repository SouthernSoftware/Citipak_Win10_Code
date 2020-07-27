DECLARE SUB printbgttrx ()
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
'DECLARE SUB TitleBox (Row, Col, MaxLen, Text$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, X%, Y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortAcctIndex ()
DECLARE SUB ResetAcct ()
DECLARE SUB PrintIndex ()
DECLARE SUB PrintAcctListRaw ()
DECLARE SUB SearchforDupes ()
DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()

DECLARE SUB FixPRIF ()
DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType

TYPE LastIFRecType
  GLAcct AS STRING * 14       '9 AS tranacct$
  TranDate AS STRING * 8     '8 AS trandate$
  Desc AS STRING * 20        '20 AS trandesc$
  CrAmt AS DOUBLE            '8 AS cramt$
  DrAmt AS DOUBLE            '8 AS dramt$
  Ref AS STRING * 8          '8 AS detail$
  Src AS STRING * 8          '8 AS source$
  filler AS STRING * 11      '4 AS nexttr$
END TYPE

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes

REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

'REDIM MChoice$(1 TO 4)
REDIM MChoice$(1 TO 8)

MChoice$(1) = "Add/Change/Delete G/L Accounts"
MChoice$(2) = "Print Account File"
MChoice$(3) = "Search for Duplicate Accts"
MChoice$(4) = "Re-sort Indexes"
MChoice$(5) = "Re-Link Trans"
MChoice$(6) = "Re-Link Bgt Trans"
MChoice$(7) = "Quit"
MChoice$(8) = "Post Last PR Interface"

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

printbgttrx
'MoveBgt
'ResetAcct
'GOTO jumpout

CLS
DO

   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode

   'TitleBox 3, Col, MaxLen + 3, "Account Maintenance Menu ", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   Title$ = "G/L Util Menu"
   MnuTitle Title$, Row, Col, MaxLen, 15
   
   PrintTitle User$
   PrintHelp "GLUTIL1"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
         ClrBG
         'EditAcct
      CASE 2
         ClrBG
         PrintAcctListRaw
      CASE 3
        SearchforDupes
      CASE 4
        ReSortAcctIndex
      CASE 5
        ReLinkTrans
      CASE 6
        RelinkBgtTrans
      CASE 7
         CLS : END
      CASE 8
        FixPRIF

      CASE ELSE
   END SELECT

LOOP


jumpout:
HideCursor                                'Be nice and turn off mouse cursor
CLS
END

SUB FixPRIF

   DIM LIFRec AS LastIFRecType
   LRecLen = LEN(LIFRec)

   DIM GLTR AS GLTransRecType
   RecLen = LEN(GLTR)

   OPEN "LASTIF.DAT" FOR RANDOM AS #1 LEN = LRecLen
   OPEN "PRIF.DAT" FOR RANDOM AS #2 LEN = RecLen

   'OPEN "IF.PRN" FOR OUTPUT AS #3

   NumOFRecs = LOF(1) \ LRecLen

   FOR Cnt = 1 TO NumOFRecs

      GET #1, Cnt, LIFRec

       TAcctNum$ = QPTrim$(LIFRec.GLAcct)
       ALen = LEN(TAcctNum$)
       IF ALen = 7 THEN TAcctNum$ = TAcctNum$ + "00"
       NAcctNum$ = FmtAcct$(TAcctNum$, 2, 3, 4)
       '
       GLTR.TrDate = Date2Num(LIFRec.TranDate)
       GLTR.AcctNum = NAcctNum$ 'LIFRec.GLAcct
       GLTR.Desc = LIFRec.Desc
       GLTR.Ref = LIFRec.Ref
       GLTR.DrAmt = Round#(LIFRec.DrAmt)
       GLTR.CrAmt = Round#(LIFRec.CrAmt)
       GLTR.Src = LIFRec.Src

       TotDr# = TotDr# + GLTR.DrAmt
       TotCr# = TotCr# + GLTR.CrAmt
     
     PUT #2, Cnt, GLTR


   NEXT

   CLOSE

   CLS
   PRINT TotDr#
   PRINT TotCr#
   PRINT "Posting..."
   Post2GL "PRIF.DAT", BadTrans%
   
END SUB

SUB PackAcctFile

'   CLS
'   COLOR 11
'   LOCATE 1, 30
'   PRINT "Pack Account File"
'
'   COLOR 10
'   LOCATE 2, 1
'   PRINT STRING$(80, CHR$(196))
'
'   DIM Acct AS BAAcctRecType
'   DIM NAcct AS BAAcctRecType
'
'   AcctRecLen = LEN(Acct)
'
'   ESCKey = 27
'   F10Key = 68
'
'   OpenBAAcctFile BAAcctFile%, NumAccts%, ARecLen
'
'   NewAcctFile = FREEFILE
'   OPEN "baacct.new" FOR RANDOM AS NewAcctFile LEN = AcctRecLen
'
'   GoodRecs = 0
'
'   FOR I = 1 TO NumAccts
'      GET #1, I, Acct
'      LOCATE 10, 1: PRINT "Processing Record# :"; I
'
'      GET #1, I, Acct
'
'      IF LEN(QPTrim$(Acct.Num)) > 0 THEN
'         GoodRecs = GoodRecs + 1
'         NAcct.Num = Acct.Num
'         NAcct.Title = Acct.Title
'         NAcct.Typ = Acct.Typ
'         NAcct.Bal = Acct.Bal
'         NAcct.MTD = Acct.MTD
'         NAcct.YTD = Acct.YTD
'         NAcct.Bgt = Acct.Bgt
'         NAcct.FrstTran = Acct.FrstTran
'         NAcct.LastTran = Acct.LastTran
'         NAcct.BegBal = Acct.BegBal
'         'NAcct.ReCalcBal = Acct.ReCalcBal  'unrem for 7-9
'         'NAcct.Encumb = Acct.Encumb        'unrem for 7-9
'
'         PUT NewAcctFile, GoodRecs, NAcct
'
'      END IF
'
'   NEXT
'
'   CLOSE
'
'   NAME "baacct.dat" AS "baacct.old"
'   NAME "baacct.new" AS "baacct.dat"
'
'   LOCATE 23, 1: COLOR 14
'   PRINT "Press any key to exit."
'   K$ = INPUT$(1)
'
'   COLOR 7
'   CLS
'
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

         CASE F10key                      'SaveButton
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

   LOOP UNTIL Frm.KeyCode = F10key
   
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
         linecnt = linecnt + 1
         IF linecnt > MaxLines THEN
           PRINT #PRNFile, FF$
           GOSUB PrintAcctPageHeader
         END IF
      END IF

   NEXT

   PRINT #PRNFile,
   PRINT #PRNFile, HowMany; "Accounts listed."
   'PRINT #PRNFile, acnt
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
   linecnt = 4
RETURN


END SUB

SUB PrintAcctListRaw

   CLS
   
   'LOCATE , , 0
   'REDIM Form$(0, 0)
   'REDIM Fld(0) AS FieldInfo
   'REDIM Choice$(0 TO 3, 1)
   'DIM Frm AS FormInfo
   'LibName$ = "GL.QSL"
   'FormName$ = "PRNACCT"
   '
   'Choices for Vendor Status field
   'Choice$(0, 0) = "2"
   'Choice$(1, 0) = "Screen"
   'Choice$(2, 0) = "Prn#1"
   'Choice$(3, 0) = "Prn#2"
   '
   'LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   'PrintHelp "Print Account Listing"
   'Action = 1
   'DO                                      'Allow user to edit the form
   '
   '   EditForm Form$(), Fld(), Frm, Cnf, Action
   '
   '   SELECT CASE Frm.KeyCode             'Check the user's key presses
   '
   '      CASE F10Key                      'SaveButton
   '         Lookfor$ = RTRIM$(Form$(1, 0))
   '         IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN
   '            Dev$ = "S"
   '            LPTNo = 1
   '         ELSE
   '            Dev$ = LEFT$(Form$(2, 0), 1)
   '            LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
   '         END IF
   '
   '      CASE EscKey
   '         EXIT SUB
   '
   '   END SELECT
   '
   'LOOP UNTIL Frm.KeyCode = F10Key
   '
   LOCATE , , 0
  
   '--Define vars used for printing
   Header$ = "Account File"
   FF$ = CHR$(12)
   MaxLines = 55

   '--Background screen
   'LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, ErrorCode
   'PrintHelp "Processing report. Please wait."

   OpenAcctIdx AcctIdxFileNum, NumIdxRecs
   OpenAcctFile AcctFileNum, NumAccts
   PRNFile = FREEFILE
   ReportFile$ = "ACCTDUMP.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile

   GOSUB PrintAcctPageHeaderRaw

   FOR Cnt = 1 TO NumAccts              'for raw printout
      GET AcctFileNum, Cnt, Acct
         'IF QPTrim$(Acct.Num) = "10-420-7400" THEN
           HowMany = HowMany + 1
           ToPrint$ = SPACE$(80)
           MID$(ToPrint$, 2) = Acct.Num
           MID$(ToPrint$, 18) = Acct.Title
           MID$(ToPrint$, 63) = Acct.Typ
           MID$(ToPrint$, 65) = STR$(Cnt)
           MID$(ToPrint$, 70) = STR$(Acct.Deleted)
           PRINT #PRNFile, ToPrint$
           linecnt = linecnt + 1
           IF linecnt > MaxLines THEN
             PRINT #PRNFile, FF$
             GOSUB PrintAcctPageHeaderRaw
           END IF
         'END IF
      'END IF

   NEXT

   PRINT #PRNFile,
   PRINT #PRNFile, HowMany; "Accounts listed."
   PRINT #PRNFile, acnt
   PRINT #PRNFile, FF$
   CLOSE

   'SELECT CASE Dev$
   '   CASE "S"
         PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
   '   CASE "P"
   '      PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   'END SELECT

   'KILL ReportFile$
   ShowCursor

EXIT SUB

PrintAcctPageHeaderRaw:
   PRINT #PRNFile, Header$
   PRINT #PRNFile,
   PRINT #PRNFile, "Acct Number     Title                                       Type/Rec#/Deleted"
   PRINT #PRNFile, STRING$(78, "-")
   linecnt = 4
RETURN



END SUB

SUB printbgttrx

  LOCATE 1, 1, 0

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "New"
  Choice$(2, 0) = "Existing"
  Choice$(0, 1) = "2"
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Prn#1"
  Choice$(3, 1) = "Prn#2"

  LibName$ = "GLUTIL.QSL"           'screen library
  FormName$ = "PRNTR"         'form to edit
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode     'Check the user's key presses
      CASE EscKey
        ExitFlag = True
        EXIT DO
      CASE F10key
        ExitFlag = False
        EXIT DO
    END SELECT

  LOOP

  IF ExitFlag THEN EXIT SUB

  IF Form$(1, 0) = "E" THEN
    IF Exist("trlist.prn") THEN
      RptFileName$ = "trlist.prn"
      GOTO PrintRpt
    END IF
  END IF
  
  DIM Trans AS GLTransRecType
  TransRecLen = LEN(Trans)
  BgtTransFile = FREEFILE
  OPEN "BGTTRANS.DAT" FOR RANDOM AS BgtTransFile LEN = TransRecLen
  NumTrans& = LOF(BgtTransFile) \ TransRecLen

  '--open a report file to print to
  RptFile = FREEFILE
  RptFileName$ = "TRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  QPrintRC "Processing.    % Complete", 25, 2, -1

  FOR TrRec& = 1 TO NumTrans&
     Complete! = (TrRec& / NumTrans&) * 100
     Pct$ = FUsing(STR$(Complete!), "###")
     QPrintRC Pct$, 25, 14, -1

     GET BgtTransFile, TrRec&, Trans

     ToPrint$ = SPACE$(132)
     LSET ToPrint$ = Trans.AcctNum
     MID$(ToPrint$, 16) = ConvDateStr$(Num2Date(Trans.TrDate))
     MID$(ToPrint$, 24) = LEFT$(Trans.Desc, 13)
     MID$(ToPrint$, 39) = Trans.Ref
     MID$(ToPrint$, 47) = FUsing$(STR$(Trans.DrAmt), "########.##")
     MID$(ToPrint$, 60) = FUsing$(STR$(Trans.CrAmt), "########.##")
     MID$(ToPrint$, 73) = Trans.Src
     MID$(ToPrint$, 84) = "Tr#:" + STR$(TrRec&)
     'MID$(ToPrint$, 96) = "Nx:" + STR$(Trans.NextTran)
     PRINT #RptFile, ToPrint$

     TCnt& = TCnt& + 1
     Debits# = Round#(Debits#) + Round#(Trans.DrAmt)
     Credits# = Round#(Credits#) + Round#(Trans.CrAmt)

  NEXT

  Diff# = Round#(Debits#) - Round#(Credits#)
  PRINT #RptFile,
  PRINT #RptFile, "File Totals"
  PRINT #RptFile, "---------------"
  PRINT #RptFile, "Total Records  : "; FUsing$(STR$(TCnt&), "##########,")
  PRINT #RptFile, "Debit Total    : "; FUsing$(STR$(Debits#), "##########,.##")
  PRINT #RptFile, "Credit Total   : "; FUsing$(STR$(Credits#), "##########,.##")
  PRINT #RptFile, "Balance        : "; FUsing$(STR$(Diff#), "##########,.##")

  CLOSE

  IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN
    EntryPoint = 2
    Dev$ = "S"
    LPTNo = 1
  ELSE
    EntryPoint = 5
    Dev$ = LEFT$(Form$(2, 0), 1)
    LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
  END IF

PrintRpt:
  RptTitle$ = "List Transaction Records"
  EntryPoint = 2
  PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint






END SUB

SUB PrintIndex
   LOCATE 1, 1, 0
   OpenAcctIdx AcctIdxFileNum, NumIdxRecs

   PRNFile = FREEFILE
   ReportFile$ = "TEST.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile

   ToPrint$ = SPACE$(80)

   FOR Cnt = 1 TO NumIdxRecs
    GET AcctIdxFileNum, Cnt, AcctIdx
    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 2) = AcctIdx.AcctNum
    MID$(ToPrint$, 18) = STR$(AcctIdx.RecNum)
    PRINT #PRNFile, ToPrint$
   NEXT

   PRINT #PRNFile, FF$
   CLOSE

   'SELECT CASE Dev$
   '   CASE "S"
         PrintRptFile "Testing", ReportFile$, LPTNo, RetCode, 2
   '   CASE "P"
   '      PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
   'END SELECT

   'KILL ReportFile$
   'ShowCursor


END SUB

SUB ResetAcct

   PRNFile = FREEFILE
   ReportFile$ = "TEMP.PRN"
   OPEN ReportFile$ FOR OUTPUT AS #PRNFile
   
   '--set P/Y Balance to 0
   CLS
   LOCATE 1, 1, 0
   OpenAcctFile AcctFileNum, NumAccts
   FOR Cnt = 1 TO NumAccts
      GET AcctFileNum, Cnt, Acct
      PRINT #PRNFile, "Account: "; Acct.Num; " Beg Bal: "; Acct.BegBal
      'Acct.BegBal = 0
      'PUT AcctFileNum, Cnt, Acct
   NEXT

   CLOSE
   PrintRptFile "GLUTIL", ReportFile$, LPTNo, RetCode, 0

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

SUB SearchforDupes

  CLS
  LOCATE 1, 1, 0

  '--Define vars used for printing
  Header$ = "Duplicate Account Listing"
  FF$ = CHR$(12)
  MaxLines = 55

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs

  PRNFile = FREEFILE
  ReportFile$ = "DUPEACCT.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PRNFile

  FOR Cnt = 1 TO NumIdxRecs
    GET AcctIdxFileNum, Cnt, AcctIdx
      AccttoFind$ = AcctIdx.AcctNum
      LOCATE 1, 1: PRINT "Searching: " + AccttoFind$
        FOR SCnt = 1 TO NumIdxRecs
        GET AcctIdxFileNum, SCnt, AcctIdx
        IF AcctIdx.AcctNum = AccttoFind$ AND SCnt <> Cnt THEN
          BEEP
          DCnt = DCnt + 1
          GOSUB LogDupe
        END IF
      NEXT
  NEXT

  IF DCnt = 0 THEN
    PRINT #PRNFile, "No duplicate account numbers found."
  END IF
  PRINT #PRNFile, FF$
  CLOSE
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
  KILL ReportFile$
  ShowCursor

EXIT SUB

LogDupe:
   PRINT #PRNFile, AccttoFind$, AcctIdx.RecNum
RETURN


END SUB

