DECLARE SUB changedate ()
DECLARE SUB ListPOTrans ()
DECLARE SUB ReLinkPOTrans ()
'GLUTIL.BAS (c) Southern Software, Inc.
'by Jim Brown

DEFINT A-Z
DECLARE SUB EditTransRec ()
'DECLARE SUB ChangeDate3 ()
'DECLARE SUB TransChangeAcct ()
'DECLARE SUB SetFieldValue ()
DECLARE SUB UnMarkRecs ()
DECLARE SUB ViewLog ()
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, X%, Y%, Wdth%, TitColor%)
DECLARE SUB relink ()
DECLARE SUB KillTrans ()
DECLARE SUB MarkRecs (mark)
DECLARE SUB ListTransFile ()
DECLARE SUB ChangeDateSR ()
DECLARE SUB ClearQuery ()
DECLARE SUB ListMTrans ()
DECLARE SUB UpdateQuery ()
DECLARE SUB MarkQuery ()
DECLARE SUB AcctSub ()
DECLARE SUB TransWrk ()

DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, R%, c%, clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()
DECLARE SUB ChangeBgtDate ()

'---
DECLARE SUB Srch4DupeAccts ()
DECLARE SUB EdAcct ()
DECLARE SUB Fixbgted ()
DECLARE SUB DateScan ()
DECLARE SUB GetFYDates (Fy1BegDate%, FY1EndDate%, FY2BegDate%, FY2EndDate%)

CONST False = 0, True = NOT False

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations
''$INCLUDE: 'pro7.bi'

DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Trans AS GLTransRecType

'$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
'$INCLUDE: 'init.bi'
  
'changedate
'setfieldvalue

  LOCATE 1, 1, 0
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it
  
  REDIM MChoice$(1 TO 21)
  MChoice$(1) = "Print Transaction File"
  MChoice$(2) = "Scan for Dates out of Range"
  MChoice$(3) = "Edit Transactions"
  MChoice$(4) = "Update Query"
  MChoice$(5) = "List Tagged Trans"
  MChoice$(6) = "Tag Records"
  MChoice$(7) = "Un-Tag Records"
  MChoice$(8) = "Clear All Tagged Records"
  MChoice$(9) = "Delete Tagged Records"
  MChoice$(10) = "Relink G/L Trans Database"
  MChoice$(11) = "Relink Bgt Trans Database"
  MChoice$(12) = "Search & Replace Trans Date"
  MChoice$(13) = "Search for Duplicate Accts"
  MChoice$(14) = "Edit Account"
  'MChoice$(14) = "Fix GJ Edit File"
  MChoice$(15) = "Fix Bgt Edit File"
  MChoice$(16) = "Fix Bgt Transaction Date"
  MChoice$(17) = "View G/L Util Log"
  MChoice$(18) = "Re-Link AP Ledger"
  MChoice$(19) = "Quit"
  MChoice$(20) = "ReLink POTrans"
  MChoice$(21) = "List PO Trans"

  MaxLen = 0  'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MChoice$)               'Find max menu width
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = 6'((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "General Ledger Utility Menu"

  'GOTO StartHere
  'SetFieldValue

  DO
    
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode

    Title$ = "Utility Menu "
    MnuTitle Title$, Row, Col, MaxLen, 15
    PrintTitle User$
    PrintHelp Help$
    
    LOCATE Row, Col, 0
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf
    
    IF ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
      CASE 1
        ListTransFile
      CASE 2
        DateScan
      CASE 3
        EditTransRec
      CASE 4
        MarkQuery
      CASE 5
        ListMTrans
      CASE 6
        MarkRecs True
      CASE 7
        MarkRecs False
      CASE 8
        ClearQuery
      CASE 9
        KillTrans
      CASE 10
        ReLinkTrans
      CASE 11
        RelinkBgtTrans
      CASE 12
        ChangeDateSR
        'ChangeDateMarked
      CASE 13
        Srch4DupeAccts
      CASE 14
        EdAcct
      'CASE 14
      '  BEEP
      CASE 15
        Fixbgted
      CASE 16
        ChangeBgtDate
      CASE 17
        ViewLog
      CASE 18
        RUN "APLINK"
      CASE 19
        CLS
        END
      CASE 20
        ReLinkPOTrans
      CASE 21
        ListPOTrans
    END SELECT
    
  LOOP

  ClrBG
  HideCursor '--Be nice and turn the mouse cursor Off
  'CLS
  'END

  DispLoad

  RUN "GLSYS"
  END

'StartHere:

'   DIM POTrans AS GLTransRecType
'   TransRecLen = LEN(POTrans)
'   POTransFile = FREEFILE
'   OPEN "POTRANS.DAT" FOR RANDOM AS #1 LEN = TransRecLen
'   OPEN "POTRANS.NEW" FOR RANDOM AS #10 LEN = TransRecLen
'
'   NumTrans& = LOF(POTransFile) \ TransRecLen
'   FOR Cnt = 1 TO NumTrans&
'     SELECT CASE Cnt
'     CASE 51 TO 52, 605 TO 616, 643 TO 646
'     CASE ELSE
'       GET #1, Cnt, POTrans
'       PUT #10, , POTrans
'     END SELECT
'   NEXT
'   CLOSE

'   TransRecLen = LEN(Trans)
'   OPEN "GLTRANS.DAT" FOR RANDOM AS #1 LEN = TransRecLen
'   OPEN "GLTRANS.NEW" FOR RANDOM AS #10 LEN = TransRecLen
'   NumTrans& = LOF(1) / TransRecLen
'   FOR Cnt& = 1 TO NumTrans&
'     GET #1, Cnt&, Trans
'     ThisAcct$ = LEFT$(Trans.AcctNum, 4)
'     IF (INSTR(ThisAcct$, "12-") > 0) OR (INSTR(ThisAcct$, "32-") > 0) THEN
'       PUT #10, , Trans
'     END IF
'   NEXT
'   CLOSE

SUB AcctSub
'
'  DIM AcctIdx AS GLAcctNdxType
'  SHARED Acct AS GLAcctRecType
'
'  PrnFile = FREEFILE
'  RptFileName$ = "acctsub.prn"
'  OPEN RptFileName$ FOR OUTPUT AS PrnFile
'
'
'  CLS
'  OpenAcctFile GLAcctFile, NumAccts
'
'  FOR Cnt = 1 TO NumAccts
'    GET GLAcctFile, Cnt, Acct
'    Acct.BegBal = 0
'    PUT GLAcctFile, Cnt, Acct
'  NEXT
'
'  CLOSE
'
'  EXIT SUB
'
'
'---reset budget
'   OpenAcctFile GLAcctFile, NumAccts
'   FOR Cnt = 1 TO NumAccts
'      LOCATE 12, 1
'      PRINT "Setting budget amounts to zero. Processing: "; Acct.Num
'      GET GLAcctFile, Cnt, Acct
'      Acct.Bgt = 0
'      PUT GLAcctFile, Cnt, Acct
'   NEXT
'   CLOSE
'
END SUB

SUB ChangeBgtDate

  LibName$ = "GLUTIL.QSL"
  FormName$ = "S&RDATE"

  SHARED Choice$()
  REDIM Choice$(0 TO 1, 1)
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "BG - Budget Transactions"

  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo              'Field info Array
  DIM Frm AS FormInfo                    'Form editing info

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  GoodRec = False
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode             'Check the user's key presses
      CASE F10key                      'SaveButton
        ExitFlag = False
        EXIT DO
      CASE EscKey
        ExitFlag = True
        EXIT DO
    END SELECT
  LOOP

  IF ExitFlag THEN EXIT SUB

  ODate = Date2Num(Form$(1, 0))
  NDate = Date2Num(Form$(3, 0))
  Mod$ = Form$(2, 0)
  Changed = 0

  '--Open the Budget Transaction File.
  DIM BTrans AS GLTransRecType
  BTransRecLen = LEN(BTrans)
  BTransFile = FREEFILE
  OPEN "BGTTRANS.DAT" FOR RANDOM AS BTransFile LEN = BTransRecLen
  NumBTrans& = LOF(BTransFile) \ BTransRecLen
  
  '--open a report file to print to
  RptFile = FREEFILE
  RptFileName$ = "glutil.log"
  OPEN RptFileName$ FOR APPEND AS RptFile
  PRINT #RptFile,
  PRINT #RptFile, "Budget Date change procedure ran on " + DATE$ + " @ " + TIME$
  
  QPrintRC SPACE$(80), 25, 1, -1
  QPrintRC "Scanning file ", 25, 2, -1

  FOR TrRec& = 1 TO NumBTrans&
     Done# = (TrRec& / NumBTrans&) * 100
     Pct$ = FUsing(STR$(Done#), "###") + "% Complete"
     QPrintRC Pct$, 25, 15, -1

     GET BTransFile, TrRec&, BTrans
      IF BTrans.TrDate = ODate THEN
        GOSUB LogBgtDateChange
      END IF

  NEXT

  REDIM R$(1 TO 4)
  R$(1) = "The dates of the above transactions were changed to " + Num2Date(NDate)
  R$(2) = "Bgt Transaction(s) Changed : " + FUsing$(STR$(Changed), "#########")
  R$(3) = "Debit Total  : " + FUsing$(STR$(GTDebit#), "########,.##")
  R$(4) = "Credit Total : " + FUsing$(STR$(GTCredit#), "########,.##")

  IF Changed > 0 THEN
    FOR i = 1 TO 4
      PRINT #RptFile, R$(i)
    NEXT
    PRINT #RptFile,
  ELSE
    PRINT #RptFile, "No transactions were changed."
  END IF

  CLOSE

  LibFile2Scrn LibName$, "OKBOX", MonoCode, Attribute, ErrorCode
  FOR i = 2 TO 4
    QPrintRC R$(i), 7 + i, 17, -1
  NEXT
  K$ = INPUT$(1)


EXIT SUB


LogBgtDateChange:
   '--Record change in log
   ToPrint$ = SPACE$(132)
   LSET ToPrint$ = BTrans.AcctNum
   MID$(ToPrint$, 17) = Num2Date$(BTrans.TrDate)
   MID$(ToPrint$, 28) = LEFT$(BTrans.Desc, 15)
   MID$(ToPrint$, 46) = FUsing$(STR$(BTrans.DrAmt), "########,.##")
   MID$(ToPrint$, 59) = FUsing$(STR$(BTrans.CrAmt), "########,.##")
   MID$(ToPrint$, 72) = BTrans.Src
   PRINT #RptFile, ToPrint$

   '--Change date in transaction file
   BTrans.TrDate = NDate
   PUT BTransFile, TrRec&, BTrans

   Changed = Changed + 1
   GTDebit# = GTDebit# + BTrans.DrAmt
   GTCredit# = GTCredit# + BTrans.CrAmt

   '--Show change on screen
   M$ = "Changed: " + FUsing(STR$(Changed), "#####")
   QPrintRC M$, 25, 45, -1

RETURN
  

END SUB

SUB changedate

  OpenTransFile GLTransFile, NumTrans&

  NewDate = Date2Num("10/31/96")

  QPrintRC SPACE$(80), 25, 1, -1
  QPrintRC "Scanning file ", 25, 2, -1

  FOR TrRec& = 1 TO NumTrans&
     Done# = (TrRec& / NumTrans&) * 100
     Pct$ = FUsing(STR$(Done#), "###") + "% Complete"
     QPrintRC Pct$, 25, 15, -1
     GET GLTransFile, TrRec&, Trans
     Trans.TrDate = NewDate
     PUT GLTransFile, TrRec&, Trans
  NEXT

  CLOSE

END SUB

SUB ChangeDateMarked

  CLS
  COLOR 11
  Col = 1

  LOCATE 1, 1, 0

  QPrintRC "Change Marked Transaction Dates", 1, 1, 10

  QPrintRC "Esc=Cancel Enter=Go", 3, 1, 15
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  
  QPrintRC SPACE$(80), 3, 1, 15

  LOCATE 3, 1
  INPUT "New Date: ", D$
  NewDate = Date2Num(D$)
  ND$ = Num2Date(NewDate)

  Msg$ = "Enter the password to Change Marked Transaction Dates to: " + ND$
  QPrintRC Msg$, 3, 1, 15

  FOR KCnt = 1 TO 3
    K$ = INPUT$(1)
    IF K$ = CHR$(27) THEN EXIT SUB
    P$ = P$ + K$
  NEXT
  P$ = "YES"
  IF P$ = "YES" THEN

    OpenTransFile TransFileNum, NumTrans&

    QPrintRC SPACE$(80), 3, 1, 15
    LOCATE 3, Col
    PRINT USING "Total Records:   ######"; NumTrans&

    DrAmt# = 0
    CrAmt# = 0
    RCnt& = 0

    FOR RecNo = 1 TO NumTrans&
     GET TransFileNum, RecNo, Trans
     IF Trans.Marked = True THEN

        '--Set to New Date
        Trans.TrDate = NewDate
        PUT TransFileNum, RecNo, Trans

        '--Summarize
        RCnt& = RCnt& + 1
        DrAmt# = DrAmt# + Trans.DrAmt
        CrAmt# = CrAmt# + Trans.CrAmt

        LOCATE 4, Col, 0
        PRINT USING "Records Changed: ######"; RCnt&

      END IF
    NEXT

    CLOSE

    Diff# = Round#(DrAmt#) - Round#(CrAmt#)
    Bal# = ABS(Diff#)

    QPrintRC "Changed Transaction Summary:", 6, 1, 10

    LOCATE 7, Col
    PRINT USING "Total Debits:  #########,.##"; DrAmt#
    LOCATE 8, Col
    PRINT USING "Total Credits: #########,.##"; CrAmt#

    QPrintRC "Press any key to continue.", 11, 1, 15
    K$ = INPUT$(1)
    
  ELSE
    BEEP
    QPrintRC "Operation aborted. You must have the password perform this opeartion", 5, 1, 12
    QPrintRC "Press any key to continue", 6, 1, 14
    K$ = INPUT$(1)
  END IF
    
  
END SUB

SUB ChangeDateSR

  LibName$ = "GLUTIL.QSL"
  FormName$ = "S&RDATE"
  SHARED Choice$()
  REDIM Choice$(0 TO 9, 1)
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "AP"
  Choice$(2, 0) = "CR"
  Choice$(3, 0) = "CD"
  Choice$(4, 0) = "GJ"
  Choice$(5, 0) = "PR"
  Choice$(6, 0) = "UB"
  Choice$(7, 0) = "TB"
  Choice$(8, 0) = "AR"
  Choice$(9, 0) = "CK"
  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo              'Field info Array
  DIM Frm AS FormInfo                    'Form editing info

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  GoodRec = False
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode             'Check the user's key presses
      CASE F10key                      'SaveButton
        ExitFlag = False
        EXIT DO
      CASE EscKey
        ExitFlag = True
        EXIT DO
    END SELECT
  LOOP

  IF ExitFlag THEN EXIT SUB
  
  ODate = Date2Num(Form$(1, 0))
  NDate = Date2Num(Form$(3, 0))
  Mod$ = Form$(2, 0)
  Changed = 0
  
  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FREEFILE
  RptFileName$ = "glutil.log"
  OPEN RptFileName$ FOR APPEND AS RptFile
  PRINT #RptFile,
  PRINT #RptFile, "Date change procedure ran on " + DATE$ + " @ " + TIME$
  PRINT #RptFile, "Transactions changed:"

  QPrintRC SPACE$(80), 25, 1, -1
  QPrintRC "Scanning file ", 25, 2, -1

  FOR TrRec& = 1 TO NumTrans&
     Done# = (TrRec& / NumTrans&) * 100
     Pct$ = FUsing(STR$(Done#), "###") + "% Complete"
     QPrintRC Pct$, 25, 15, -1
     GET GLTransFile, TrRec&, Trans
      IF Trans.TrDate = ODate THEN
        TMod$ = LEFT$(Trans.Src, 2)
        IF TMod$ = Mod$ THEN
           GOSUB LogDateChange
        END IF
     END IF
  NEXT

  REDIM R$(1 TO 4)
  R$(1) = "The dates of the above transactions were changed to " + Num2Date(NDate)
  R$(2) = "Trans Changed  : " + FUsing$(STR$(Changed), "#########")
  R$(3) = "Total Debits   : " + FUsing$(STR$(Debits#), "########,.##")
  R$(4) = "Total Credits  : " + FUsing$(STR$(Credits#), "########,.##")

  IF Changed > 0 THEN
    FOR i = 1 TO 4
      PRINT #RptFile, R$(i)
    NEXT
    PRINT #RptFile,
  ELSE
    PRINT #RptFile, "No transactions were changed."
  END IF
  
  CLOSE

  LibFile2Scrn LibName$, "OKBOX", MonoCode, Attribute, ErrorCode
  FOR i = 2 TO 4
    QPrintRC R$(i), 7 + i, 17, -1
  NEXT
  K$ = INPUT$(1)


EXIT SUB


LogDateChange:
   '--Record change in log
   ToPrint$ = SPACE$(132)
   LSET ToPrint$ = Trans.AcctNum
   MID$(ToPrint$, 17) = Num2Date$(Trans.TrDate)
   MID$(ToPrint$, 28) = LEFT$(Trans.Desc, 15)
   'MID$(ToPrint$, 36) = Trans.Ref
   MID$(ToPrint$, 46) = FUsing$(STR$(Trans.DrAmt), "########,.##")
   MID$(ToPrint$, 59) = FUsing$(STR$(Trans.CrAmt), "########,.##")
   MID$(ToPrint$, 72) = Trans.Src
   PRINT #RptFile, ToPrint$

   '--Change date in transaction file
   Trans.TrDate = NDate
   PUT GLTransFile, TrRec&, Trans
   Changed = Changed + 1
   Debits# = Debits# + Trans.DrAmt
   Credits# = Credits# + Trans.CrAmt

   '--Show change on screen
   M$ = "Changed: " + FUsing(STR$(Changed), "#####")
   QPrintRC M$, 25, 45, -1

RETURN



END SUB

SUB ClearQuery

  CLS
  COLOR 11
  Col = 1

  LOCATE 1, 1, 0

  QPrintRC "Clear Update Query", 1, 1, 10

  QPrintRC "Esc=Cancel Enter=Go", 3, 1, 15
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB

  OpenTransFile TransFileNum, NumTrans&
  
  LOCATE 3, Col
  PRINT USING "Total Records:   ######"; NumTrans&
  
  DrAmt# = 0
  CrAmt# = 0
  RCnt& = 0

  FOR RecNo& = 1 TO NumTrans&

    '--Set Flag to False
    GET TransFileNum, RecNo&, Trans
    Trans.Marked = 0
    PUT TransFileNum, RecNo&, Trans

    '--Summarize
    RCnt& = RCnt& + 1
    DrAmt# = DrAmt# + Trans.DrAmt
    CrAmt# = CrAmt# + Trans.CrAmt
    
    '--PLAYING
    'DrAmt# = Trans.DrAmt
    'CrAmt# = Trans.CrAmt
    'CrAmt# = Trans.DrAmt
    'DrAmt# = Trans.CrAmt


    LOCATE 4, Col, 0
    PRINT USING "Records Cleared: ######"; RCnt&
  NEXT

  CLOSE

  Diff# = Round#(DrAmt#) - Round#(CrAmt#)
  Bal# = ABS(Diff#)

  QPrintRC "File Totals:", 6, 1, 10
  
  LOCATE 7, Col
  PRINT USING "Total Debits:  #########,.##"; DrAmt#
  LOCATE 8, Col
  PRINT USING "Total Credits: #########,.##"; CrAmt#

  LOCATE 10, Col

  IF Diff# = 0 THEN
    COLOR 10
    DrCr$ = CHR$(1) + " Transactions are in Balance!"
    PRINT DrCr$
  ELSEIF Diff# > 0 THEN
    COLOR 14
    DrCr$ = " Debit"
    PRINT "Transactions are out of balance!"
    PRINT USING "########,.##"; Bal#;
    PRINT DrCr$
  ELSE
    COLOR 12
    DrCr$ = " Credit"
    PRINT "Transactions are out of balance!"
    PRINT USING "########,.##"; Bal#;
    PRINT DrCr$
  END IF
  
  QPrintRC "Press ESC to exit.", 23, 1, 15
  DO UNTIL INKEY$ = CHR$(27): LOOP

END SUB

SUB DateScan

   COLOR 15
   CLS
   PRINT "Scanning for Invalid Dates."
   
   GetFYDates Fy1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
   
   OpenTransFile GLTransFile, NumTrans&

   LogFile = FREEFILE
   LogfileName$ = "DATESCAN.LOG"
   OPEN LogfileName$ FOR APPEND AS #LogFile

   PRINT #LogFile,
   PRINT #LogFile, "Date Scan started @ " + DATE$ + " @ "; TIME$
   PRINT #LogFile, "Scanning date outside of " + Num2Date(Fy1BegDate) + " and " + Num2Date(FY2EndDate)
   
   FOR TrRec& = 1 TO NumTrans&
      Done# = (TrRec& / NumTrans&) * 100
      LOCATE 3, 1, 0
      PRINT USING "Scanning transaction file. ###% complete."; Done#

      GET GLTransFile, TrRec&, Trans

      IF Trans.TrDate < Fy1BegDate OR Trans.TrDate > FY2EndDate THEN
         BadDate = BadDate + 1
         GOSUB ShowTrans
      END IF

   NEXT

   'LOCATE 4, 1
   'PRINT STRING$(80, " ")
   'LOCATE 4, 1
   '
   ''IF GoodCnt& = NumTrans& THEN
   ''  PRINT "No date problems in transaction file."
   ''  PRINT #LogFile, "No date problems in transaction file."
   ''END IF
   '
   'IF BadDate > 0 THEN
   '   PRINT BadDate; "date(s) are out of range."
   '   PRINT #LogFile, BadDate; "date(s) are out of range."
   '   PRINT #LogFile, USING "Total Debits  : ########,.##"; TotDebits#
   '   PRINT #LogFile, USING "Total Credits : ########,.##"; TotCredits#
   'ELSE
   '   PRINT "No date problems in transaction file."
   '   PRINT #LogFile, "No date problems in transaction file."
   'END IF

  ' PRINT
  ' PRINT "Press any key to continue."
  ' K$ = INPUT$(1)

   CLOSE
   
   RptTitle$ = "Date Scan"
   PrintRptFile RptTitle$, LogfileName$, LPTNo, RetCode%, EntryPoint
   KILL LogfileName$


   EXIT SUB

ShowTrans:
   PRINT #LogFile,
   PRINT #LogFile, "Record: "; TrRec&
   PRINT #LogFile, "Date:   "; Num2Date(Trans.TrDate)
   PRINT #LogFile, "Desc:   "; Trans.Desc
   PRINT #LogFile, "Dr Amt: "; Trans.DrAmt
   PRINT #LogFile, "Cr Amt: "; Trans.CrAmt
   PRINT #LogFile, "Src:    "; Trans.Src
   PRINT #LogFile,
   TotDebits# = TotDebits# + Trans.DrAmt
   TotCredits# = TotCredits# + Trans.CrAmt
RETURN
   

END SUB

SUB EdAcct

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

         CASE F10key                      'SaveButton
            SaveFlag = 0
            GOSUB CheckEdit               'Check that no fields are blank
            IF SaveFlag = 0 THEN          'All ok. Save Record
               GOSUB SaveAcctRec
               GOSUB ClearAcctForm
            END IF

         CASE F3Key                       'Delete Button
            Killed = 0
            IF RecordNum > 0 THEN         'make sure we have a record
               GOSUB DeleteRec
               IF Killed = -1 THEN
                  GOSUB ClearAcctForm
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
      'RecordNum = 1183
      GET AcctFileNum, RecordNum, Acct
      Form$(2, 0) = Acct.Title
      Form$(3, 0) = Acct.Typ
   CLOSE AcctFileNum

RETURN


SaveAcctRec:
   LOCATE 1, 1, 0
   PrintHelp "Updating files. Please wait."
   OpenAcctFile AcctFileNum, NumAccts

   SELECT CASE EMode
      CASE 0                              'new record
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

      'IF Acct.LastTran = 0 THEN        'can't delete accout with trans

         PrintHelp "Updating files. Please wait."

         OpenAcctFile AcctFileNum, NumAccts

         'GET AcctFileNum, RecordNum, Acct
         Acct.Deleted = -1
         PUT AcctFileNum, RecordNum, Acct

         CLOSE AcctFileNum

         Killed = -1

         SortAcctIndex

         PrintHelp "G/L Account Maintenance"

      'ELSE                                'Acct has transactions
      '   PLAY "L32O3EC"
      '   ok = MsgBox("GL.QSL", "CANTDEL")
      '
      'END IF
   END IF

RETURN


ClearAcctForm:                               'reset form for more editing
   LOCATE 1, 1, 0
   FOR Cnt = 1 TO 3
      Form$(Cnt, 0) = ""
   NEXT

   PrintHelp "G/L Account Maintenance"
   Frm.FldNo = 1
   Action = 1

RETURN



END SUB

SUB EditTransRec

  LibName$ = "GLUTIL.QSL"
  FormName$ = "EDTRAN"
  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo              'Field info Array
  DIM Frm AS FormInfo                    'Form editing info
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  GoodRec = False
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field
      IF Frm.FldNo = 1 THEN            'Blank the form if back on acct fld
        GOSUB ClearForm
      END IF

      SELECT CASE Frm.PrevFld          'user is moving around in form
        CASE 1
          IF Frm.FldEdited THEN      'If the field was edited
            RecNum& = VAL(Form$(1, 0))
            IF RecNum& > 0 THEN
              BadAcct = False
              GOSUB GetTrans
            END IF
          END IF
        CASE 2
        '  IF Frm.FldNo <> 1 THEN
        '    Rec = FindAcct(Form$(2, 0))
        '    IF Rec = 0 THEN
        '      ok = MsgBox("GLUTIL", "INVALID")
        '      BadAcct = True
        '    END IF
        '   END IF

          IF Frm.FldEdited THEN      'If the field was edited
            'LinkWarn = True
            ok = MsgBox("GLUTIL", "WARNLINK")
          END IF
        CASE ELSE
      END SELECT
      Action = 1
    END IF

    SELECT CASE Frm.KeyCode
      CASE F10key
        '--make sure we have a valid rec num when user presses f10
        IF GoodRec THEN
          
          '--Make sure we have a valid acct number
          Rec = FindAcct(Form$(2, 0))
          IF Rec = 0 THEN
            ok = MsgBox("GLUTIL", "INVALID")
            BadAcct = True
            Frm.FldNo = 2
          ELSE
            BadAcct = False
          END IF

          'IF LinkWarn THEN
          'END IF

          '--If the acct is ok then we can save
          IF NOT BadAcct THEN
            GOSUB SaveRec
            GOSUB ClearForm
          END IF
          
        END IF
    END SELECT

  LOOP UNTIL Frm.KeyCode = EscKey

EXIT SUB


GetTrans:
  OpenTransFile TransFile, NumTrans&
  IF RecNum& > 0 AND RecNum& <= NumTrans& THEN
    GET TransFile, RecNum&, Trans
    Form$(2, 0) = Trans.AcctNum
    Form$(3, 0) = Num2Date(Trans.TrDate)
    Form$(4, 0) = Trans.Desc
    Form$(5, 0) = Trans.Ref
    Form$(6, 0) = STR$(Trans.DrAmt)
    Form$(7, 0) = STR$(Trans.CrAmt)
    Form$(8, 0) = Trans.Src
    CLOSE TransFile
    GoodRec = True
  ELSE
    ok = MsgBox("GL.QSL", "BADREC")
    Frm.FldNo = 1
    Form$(1, 0) = ""
    GoodRec = False
    GOSUB ClearForm
  END IF
RETURN

SaveRec:
  LOCATE 1, 1, 0
  'PrintHelp "Updating files. Please wait."
  OpenTransFile TransFile, NumTrans&
  Trans.AcctNum = Form$(2, 0)
  Trans.TrDate = Date2Num(Form$(3, 0))
  Trans.Desc = Form$(4, 0)
  Trans.Ref = Form$(5, 0)
  Trans.DrAmt = Value(Form$(6, 0), ErrCode)
  Trans.CrAmt = Value(Form$(7, 0), ErrCode)
  Trans.Src = Form$(8, 0)
  PUT TransFile, RecNum&, Trans
  CLOSE
RETURN

ClearForm:
  Form$(1, 0) = ""
  Form$(2, 0) = ""
  Form$(3, 0) = ""
  Form$(4, 0) = ""
  Form$(5, 0) = ""
  Form$(6, 0) = ""
  Form$(7, 0) = ""
  Form$(8, 0) = ""
  Frm.FldNo = 1
  Action = 1
RETURN

END SUB

SUB Fixbgted

   CLS
   LOCATE 1, 1, 0
   QPrintRC "Fix Unformatted Account - Budget", 1, 1, -1
   QPrintRC "Esc=Cancel, any other key to continue.", 2, 1, -1
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN
      EXIT SUB
   END IF

   DIM BgtEdit AS TrEditRecType
   BgtEdLen = LEN(BgtEdit)
   BgtEditFileNum = FREEFILE
   OPEN "BGTED.dat" FOR RANDOM ACCESS READ WRITE SHARED AS BgtEditFileNum LEN = BgtEdLen
   NumEdTrans = LOF(BgtEditFileNum) \ BgtEdLen

   FOR Cnt = 1 TO NumEdTrans
      Done! = (Cnt / NumEdTrans) * 100

      LOCATE 4, 1, 0
      PRINT USING "Processing File. ###% Complete."; Done!

      GET BgtEditFileNum, Cnt, BgtEdit
      AcctNum$ = StripAcct$(BgtEdit.AcctNum)
      BgtEdit.AcctNum = FmtAcct$(AcctNum$, FundLen%, AcctLen%, DetLen%)
      PUT BgtEditFileNum, Cnt, BgtEdit

   NEXT

   CLOSE

   QPrintRC "Operation complete. Press any key to continue.", 5, 1, -1
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN
      EXIT SUB
   END IF
   
END SUB

SUB FixGJEd

   CLS
   LOCATE 1, 1, 0
   QPrintRC "Fix Unformatted Account - G/J", 1, 1, -1
   QPrintRC "Esc=Cancel, any other key to continue.", 2, 1, -1
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN
      EXIT SUB
   END IF

   DIM GJEdit AS TrEditRecType
   GJEdLen = LEN(GJEdit)
   GJEditFileNum = FREEFILE
   OPEN "GLGJED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJEditFileNum LEN = GJEdLen
   NumEdTrans = LOF(GJEditFileNum) \ GJEdLen
   
   FOR Cnt = 1 TO NumEdTrans
      Done! = (Cnt / NumEdTrans) * 100

      LOCATE 4, 1, 0
      PRINT USING "Processing File. ###% Complete."; Done!

      GET GJEditFileNum, Cnt, GJEdit
      AcctNum$ = StripAcct$(GJEdit.AcctNum)
      GJEdit.AcctNum = FmtAcct$(AcctNum$, FundLen%, AcctLen%, DetLen%)
      PUT GJEditFileNum, Cnt, GJEdit

   NEXT

   CLOSE

   QPrintRC "Operation complete. Press any key to continue.", 5, 1, -1
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN
      EXIT SUB
   END IF


END SUB

SUB KillTrans

  CLS
  LOCATE 1, 1, 0
  COLOR 11: PRINT "Delete Marked Transactions"
  LOCATE 3, 1
  COLOR 15
  PRINT "This utility removes Marked transaction records from the history file."
  PRINT "Transactions check in but they don't check out."
  COLOR 12
  PRINT "BACKUP BEFORE RUNNING THIS OPERATION!!!"
  COLOR 14
  PRINT "Press Esc to cancel or any other key to continue."
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN
    EXIT SUB
  END IF

  IF Exist("gltrans.old") THEN
    CLS
    COLOR 11: PRINT "Delete Marked Transactions"
    LOCATE 3, 1
    COLOR 12
    PRINT "WARNING!! A Backup file from a prior operation exits. Kill it? (Y/N)"
    DO
      K$ = UCASE$(INPUT$(1))
    LOOP UNTIL K$ = "Y" OR K$ = "N"
    SELECT CASE K$
      CASE "Y"
        KILL "gltrans.old"
      CASE "N"
        EXIT SUB
    END SELECT
  END IF

  CLS
  COLOR 11: PRINT "Deleting Marked Transactions"

  COLOR 15

  DIM Tr AS GLTransRecType
  DIM NTr AS GLTransRecType

  '--open a new GLtrans.file
  NTrRecLen = LEN(NTr)
  NGLTransFile = FREEFILE
  OPEN "ngltrans.dat" FOR RANDOM ACCESS READ WRITE SHARED AS NGLTransFile LEN = NTrRecLen '85
  NumNewTrans& = LOF(NGLTransFile) \ NTrRecLen

  '--open a log file to list transactions removed
  LogFile = FREEFILE
  OPEN "glutil.log" FOR APPEND AS LogFile
  PRINT #LogFile,
  PRINT #LogFile, "Removed transactions procedure started @ " + DATE$ + " " + TIME$

  OpenTransFile GLTransFile, NumTrans&

  FOR TrRec& = 1 TO NumTrans&
    GET GLTransFile, TrRec&, Tr

    Complete# = (TrRec& / NumTrans&) * 100
    LOCATE 3, 1: PRINT USING "Processing. Please wait. ###% Complete."; Complete#

    IF Tr.Marked = 0 THEN
      '--copy good records to the new file
      NewTrans& = NewTrans& + 1
      NTr.AcctNum = Tr.AcctNum
      NTr.TrDate = Tr.TrDate
      NTr.Desc = Tr.Desc
      NTr.CrAmt = Tr.CrAmt
      NTr.DrAmt = Tr.DrAmt
      NTr.Ref = Tr.Ref
      NTr.Src = Tr.Src
      NTr.NextTran = Tr.NextTran
      PUT NGLTransFile, NewTrans&, NTr
    ELSE
      '-check em out
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = Tr.AcctNum
      MID$(ToPrint$, 13) = ConvDateStr1$(Num2Date(Tr.TrDate))
      MID$(ToPrint$, 24) = LEFT$(Tr.Desc, 19)
      MID$(ToPrint$, 42) = STR$(Tr.CrAmt)
      MID$(ToPrint$, 52) = STR$(Tr.DrAmt)
      MID$(ToPrint$, 62) = Tr.Ref
      MID$(ToPrint$, 72) = Tr.Src
      PRINT #LogFile, ToPrint$

      '--Keep track of what's gone
      Gone& = Gone& + 1
      DebitsGone# = DebitsGone# + Tr.DrAmt
      CreditsGone# = CreditsGone# + Tr.CrAmt

    END IF

  NEXT

  'screen
  LOCATE 4, 1
  PRINT "Transactions removed: "; Gone&
  PRINT "Total Debits removed: " + STR$(DebitsGone#)
  PRINT "Total Credits removed: " + STR$(CreditsGone#)

  'log
  PRINT #LogFile,
  PRINT #LogFile, "Transactions removed: "; Gone&
  PRINT #LogFile, "Total Debits removed: " + STR$(DebitsGone#)
  PRINT #LogFile, "Total Credits removed: " + STR$(CreditsGone#)
  PRINT #LogFile,

  CLOSE

  NAME "gltrans.dat" AS "gltrans.old"
  NAME "ngltrans.dat" AS "gltrans.dat"

  PRINT "Press any key to continue."
  K$ = INPUT$(1)
  CLS

  ReLinkTrans
  
END SUB

SUB ListMTrans
   
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
    IF Exist("mtrlist.prn") THEN
      RptFileName$ = "mtrlist.prn"
      GOTO PrintMRpt
    END IF
  END IF
  
  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FREEFILE
  RptFileName$ = "MTRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS RptFile
  
  QPrintRC "Processing.    % Complete", 25, 2, -1

  FOR TrRec& = 1 TO NumTrans&
    Complete! = (TrRec& / NumTrans&) * 100
    Pct$ = FUsing(STR$(Complete!), "###")
    QPrintRC Pct$, 25, 14, -1

    GET GLTransFile, TrRec&, Trans

    IF Trans.Marked THEN

      Marked& = Marked& + 1
      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = Trans.AcctNum
      MID$(ToPrint$, 16) = Num2Date(Trans.TrDate)
      MID$(ToPrint$, 27) = LEFT$(Trans.Desc, 10)
      MID$(ToPrint$, 38) = Trans.Ref
      MID$(ToPrint$, 48) = FUsing$(STR$(Trans.DrAmt), "########,.##")
      MID$(ToPrint$, 61) = FUsing$(STR$(Trans.CrAmt), "########,.##")
      MID$(ToPrint$, 74) = Trans.Src
      MID$(ToPrint$, 84) = "T#:" + STR$(TrRec&)
      MID$(ToPrint$, 96) = "Nx:" + STR$(Trans.NextTran)
      PRINT #RptFile, ToPrint$

      Debits# = Round#(Debits#) + Round#(Trans.DrAmt)
      Credits# = Round#(Credits#) + Round#(Trans.CrAmt)
      Diff# = Round#(Debits#) - Round#(Credits#)

    END IF

  NEXT

  PRINT #RptFile,
  PRINT #RptFile, "File Totals"
  PRINT #RptFile, "---------------"
  PRINT #RptFile, "Marked Records : "; FUsing$(STR$(Marked&), "##########,")
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

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT
  
PrintMRpt:
   RptTitle$ = "List Marked Transactions"
   PrintRptFile RptTitle$, RptFileName$, LPTNo, RetCode%, EntryPoint

END SUB

SUB ListPOTrans

  TDate = Date2Num("01-01-2001")
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
      GOTO PrintPORpt
    END IF
  END IF

  'OpenTransFile GLTransFile, NumTrans&
  DIM POTrans AS GLTransRecType
  TransRecLen = LEN(POTrans)
  POTransFile = FREEFILE
  OPEN "POTRANS.DAT" FOR RANDOM SHARED AS POTransFile LEN = TransRecLen
  NumTrans& = LOF(POTransFile) \ TransRecLen

  'OPEN "POTRANS.NEW" FOR RANDOM SHARED AS #10 LEN = TransRecLen
  '--open a report file to print to
  RptFile = FREEFILE
  RptFileName$ = "TRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  QPrintRC "Processing.    % Complete", 25, 2, -1

  FOR TrRec& = 1 TO NumTrans&
     Complete! = (TrRec& / NumTrans&) * 100
     Pct$ = FUsing(STR$(Complete!), "###")
     QPrintRC Pct$, 25, 14, -1

     GET POTransFile, TrRec&, POTrans
     'IF POTrans.TrDate >= TDate THEN
     ' PUT #10, , POTrans
       ToPrint$ = SPACE$(132)
       LSET ToPrint$ = POTrans.AcctNum
       MID$(ToPrint$, 16) = ConvDateStr$(Num2Date(POTrans.TrDate))
       MID$(ToPrint$, 24) = LEFT$(POTrans.Desc, 13)
       MID$(ToPrint$, 39) = POTrans.Ref
  '     IF INSTR(POTrans.Ref, "3410") > 0 THEN
  '       STOP
  '     END IF
       MID$(ToPrint$, 47) = FUsing$(STR$(POTrans.DrAmt), "########.##")
       MID$(ToPrint$, 60) = FUsing$(STR$(POTrans.CrAmt), "########.##")
       MID$(ToPrint$, 73) = POTrans.Src
       MID$(ToPrint$, 84) = "Tr#:" + STR$(TrRec&)
       'MID$(ToPrint$, 96) = "Nx:" + STR$(POTrans.NextTran)
       PRINT #RptFile, ToPrint$

       TCnt& = TCnt& + 1
       Debits# = Round#(Debits#) + Round#(POTrans.DrAmt)
       Credits# = Round#(Credits#) + Round#(POTrans.CrAmt)
    'END IF

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

PrintPORpt:
  RptTitle$ = "List Transaction Records"
  EntryPoint = 2
  PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint



END SUB

SUB ListTransFile

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
  
  OpenTransFile GLTransFile, NumTrans&

  '--open a report file to print to
  RptFile = FREEFILE
  RptFileName$ = "TRLIST.PRN"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  QPrintRC "Processing.    % Complete", 25, 2, -1

  FOR TrRec& = 1 TO NumTrans&
     Complete! = (TrRec& / NumTrans&) * 100
     Pct$ = FUsing(STR$(Complete!), "###")
     QPrintRC Pct$, 25, 14, -1

     GET GLTransFile, TrRec&, Trans

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

SUB MarkQuery

  LibName$ = "GL.QSL"           'screen library
  FormName$ = "QRYTR"         'form to edit

  SHARED Choice$()

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 8, 0 TO 2)
  DIM Frm AS FormInfo

  DIM FundIdx AS GLFundIndexType

  Choice$(0, 0) = "5"           'Choices for Acct Type
  Choice$(1, 0) = "All"
  Choice$(2, 0) = "Asset"
  Choice$(3, 0) = "Liability"
  Choice$(4, 0) = "Revenue"
  Choice$(5, 0) = "Expenditure"

  Choice$(0, 1) = "10"           'Choices for Module
  Choice$(1, 1) = "Any"
  Choice$(2, 1) = "AR"
  Choice$(3, 1) = "AP"
  Choice$(4, 1) = "CD"
  Choice$(5, 1) = "CR"
  Choice$(6, 1) = "GJ"
  Choice$(7, 1) = "PR"
  Choice$(8, 1) = "VC"

  Choice$(0, 2) = "12"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Prn#1"
  Choice$(3, 2) = "Prn#2"

  MaxLines = 60
  FF$ = CHR$(12)
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  RptTitle$ = "Transaction Query"
  TotDr# = 0
  TotCr# = 0

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  Form$(1, 0) = Num2Date(GetBegDate)
  Form$(2, 0) = Num2Date(GetEndDate)
  GetFundCodes Form$(3, 0), Form$(4, 0)
  Form$(5, 0) = "All"

  Action = 1

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

    SELECT CASE Frm.KeyCode     'Check the user's key presses

      CASE F10key 'SaveButton
        LowDate = Date2Num(Form$(1, 0))
        HighDate = Date2Num(Form$(2, 0))
        StartFund$ = QPTrim$(Form$(3, 0))
        EndFund$ = QPTrim$(Form$(4, 0))
        AType$ = QPTrim$(Form$(5, 0))
        AcctCode$ = QPTrim$(Form$(6, 0))
        ObjCode$ = QPTrim$(Form$(7, 0))
        TSrch$ = QPTrim$(Form$(8, 0))
        TRef$ = QPTrim$(Form$(9, 0))
        Src$ = QPTrim$(Form$(10, 0))
        PostDate$ = QPTrim$(Form$(11, 0))

        IF LEN(LTRIM$(RTRIM$(Form$(12, 0)))) = 0 THEN
          Dev$ = "S"
          LPTNo = 1
        ELSE
          Dev$ = LEFT$(Form$(12, 0), 1)
          LPTNo = VAL(RIGHT$(RTRIM$(Form$(12, 0)), 1))
        END IF

      CASE ESC
        EXIT SUB

    END SELECT

  LOOP UNTIL Frm.KeyCode = F10key

  IF AType$ = "All" THEN
    AType$ = ""
    AT$ = "All"
  ELSE
    AType$ = LEFT$(AType$, 1)
    AT$ = AType$
  END IF

  IF Src$ = "Any" THEN
    Src$ = ""
    Sc$ = "Any "
  ELSE
    Sc$ = Src$
  END IF

  IF AcctCode$ = "" THEN
    AC$ = "All "
  ELSE
    AC$ = AcctCode$
  END IF

  IF ObjCode$ = "" THEN
    Obj$ = "All "
  ELSE
    Obj$ = ObjCode$
  END IF

  LPDate = Date2Num("01/01/1900")
  HPDate = Date2Num("09/05/96")

  'End of Input
  '=====================================================
  'Start Report Processing

  LOCATE 1, 1, 0

  ReportFile$ = "QUERY.PRN"
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFIdxRecs

  QPrintRC SPACE$(80), 25, 1, 112
  QPrintRC " Searching:", 25, 1, -1

  GOSUB PrintQueryPageHeader

  FOR Cnt = 1 TO NumGLAccts     'NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIdx

    AcctNumber$ = QPTrim$(AcctIdx.AcctNum)
    QPrintRC AcctNumber$, 25, 13, -1

    Fund$ = LEFT$(AcctNumber$, FundLen)
    Dept$ = MID$(AcctNumber$, FundLen + 2, AcctLen)
    Det$ = RIGHT$(AcctNumber$, DetLen)

    IF Fund$ >= StartFund$ AND Fund$ <= EndFund$ THEN
      IF INSTR(Dept$, AcctCode$) AND INSTR(Det$, ObjCode$) THEN
        GET AcctFileNum, AcctIdx.RecNum, Acct
        IF INSTR(Acct.Typ, AType$) THEN

          NextTr& = Acct.FrstTran               'get the first trans for this acct

          DO UNTIL NextTr& = 0  'keep going 'til we run out of trans
            GET TransFileNum, NextTr&, Trans

            IF Trans.TrDate >= LowDate AND Trans.TrDate <= HighDate THEN

              TrDateStamp$ = MID$(Trans.Src, 3, 6)

              IF INSTR(TrDateStamp$, PostDate$) THEN
              'PDate = Date2Num(TrDateStamp$)
              'IF PDate >= LPDate AND PDate <= HPDate THEN

                IF INSTR(Trans.Src, Src$) THEN

                 'TSrch$ = "CK"
                 'SrchLen = LEN(TSrch$)
                 'IF LEFT$(Trans.Desc, SrchLen) = TSrch$ THEN  'kill me
                 IF INSTR(UCASE$(Trans.Desc), UCASE$(TSrch$)) THEN   'kill me
                  IF INSTR(Trans.Ref, TRef$) THEN
                   'Tag em
                   Match = Match + 1
                   Trans.Marked = -1
                   PUT TransFileNum, NextTr&, Trans

                   TotDr# = TotDr# + Round#(Trans.DrAmt)
                   TotCr# = TotCr# + Round#(Trans.CrAmt)

                   ToPrint$ = SPACE$(100)
                   LSET ToPrint$ = ConvDateStr1(Num2Date(Trans.TrDate))
                   MID$(ToPrint$, 12) = Trans.AcctNum
                   MID$(ToPrint$, 28) = LEFT$(Trans.Desc, 15)
                   MID$(ToPrint$, 45) = Trans.Ref
                   IF Trans.DrAmt <> 0 THEN
                     MID$(ToPrint$, 55) = FUsing$(STR$(Trans.DrAmt), CommaFmt$)
                   ELSE
                     MID$(ToPrint$, 55) = ""
                   END IF
                   IF Trans.CrAmt <> 0 THEN
                     MID$(ToPrint$, 69) = FUsing$(STR$(Trans.CrAmt), CommaFmt$)
                   ELSE
                     MID$(ToPrint$, 69) = ""
                   END IF
                   MID$(ToPrint$, 82) = RecNum$
                   MID$(ToPrint$, 90) = Trans.Src
                   PRINT #PrnFile, ToPrint$
                   LineCnt = LineCnt + 1
                   IF LineCnt > MaxLines THEN
                      PRINT #PrnFile, FF$
                      GOSUB PrintQueryPageHeader
                   END IF
                  END IF
                 END IF 'kill me
                END IF
              END IF
            END IF

            NextTr& = Trans.NextTran            'Get the next transaction

          LOOP
        END IF
      END IF
    END IF
  NEXT

  Diff# = Round#(TotDr#) - Round#(TotCr#)
  Bal# = ABS(Diff#)

  IF Diff# = 0 THEN
    DrCr$ = ""
  ELSEIF Diff# > 0 THEN
    DrCr$ = " Dr"
  ELSE
    DrCr$ = " Cr"
  END IF

  PRINT #PrnFile, STRING$(78, "-")
  PRINT #PrnFile, "Totals";
  PRINT #PrnFile, TAB(54); USING "#########,.##"; Round#(TotDr#);
  PRINT #PrnFile, TAB(68); USING "#########,.##"; Round#(TotCr#)
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Summary"
  PRINT #PrnFile, USING "Balance:                   #########,.##"; Round#(Bal#); : PRINT #PrnFile, DrCr$
  PRINT #PrnFile, USING "Number of transactions:        #####,"; Match
  PRINT #PrnFile, FF$

  CLOSE

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
    CASE ELSE
  END SELECT

  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  ' ReportFile$

EXIT SUB

PrintQueryPageHeader:
  PRINT #PrnFile, RptTitle$
  PRINT #PrnFile,
  PRINT #PrnFile, "Date Range: " + Form$(1, 0) + " thru " + Form$(2, 0) + SPACE$(2) + "Module: " + Sc$
  PRINT #PrnFile, "Funds " + StartFund$ + " thru " + EndFund$ + "  Account: " + AC$ + "  Object: " + Obj$ + "  Account Type: " + AT$
  PRINT #PrnFile, "Date        Acct No        Desc             Ref             Debit         Credit          Source"
  PRINT #PrnFile, "------------------------------------------------------------------------------------------------"
  LineCnt = 6
RETURN




END SUB

SUB MarkRecs (mark)
  
  LibName$ = "GLUTIL.QSL"
  FormName$ = "TAGTR"
  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo              'Field info Array
  DIM Frm AS FormInfo                    'Form editing info
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode             'Check the user's key presses
      CASE F10key                      'SaveButton
        ExitFlag = False
        EXIT DO
      CASE EscKey
        ExitFlag = True
        EXIT DO
    END SELECT
  LOOP

  IF ExitFlag THEN EXIT SUB
  
  First& = VAL(Form$(1, 0))
  Last& = VAL(Form$(2, 0))
  
  OpenTransFile TransFileNum, NumTrans&
  IF Last& > NumTrans& THEN
    Last& = NumTrans&
  END IF

  LibFile2Scrn LibName$, "BOX", MonoCode, Attribute, ErrorCode

  FOR RecNo& = First& TO Last&

    '--Set Flag to False
    GET TransFileNum, RecNo&, Trans
    Trans.Marked = mark
    PUT TransFileNum, RecNo&, Trans

    '--Summarize
    RCnt& = RCnt& + 1
    DrAmt# = DrAmt# + Trans.DrAmt
    CrAmt# = CrAmt# + Trans.CrAmt

    LOCATE 8, 17, 0
    PRINT USING "Records Marked: ######"; RCnt&

  NEXT

  CLOSE

  Diff# = Round#(DrAmt#) - Round#(CrAmt#)
  Bal# = ABS(Diff#)

  QPrintRC "File Totals:", 6, 1, 10

  LOCATE 9, 17
  PRINT USING "Total Debits:  #########,.##"; DrAmt#
  LOCATE 10, 17
  PRINT USING "Total Credits: #########,.##"; CrAmt#

END SUB

SUB setfieldvalue

  STOP

  CLS
  LOCATE 1, 1, 0
  QPrintRC "Set Value in Transaction File.", 1, 1, 10
  QPrintRC "Esc=Cancel Enter=Go", 3, 1, 15
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB

  OpenTransFile TransFileNum, NumTrans&
  
  FOR T& = 31763 TO 32131
    GET TransFileNum, T&, Trans
    Trans.Src = "PR120496"
    PUT TransFileNum, T&, Trans
  NEXT
  CLOSE

END SUB

SUB Srch4DupeAccts
  
  LOCATE 1, 1, 0
  LibFile2Scrn "GLUTIL", "BOX", MonoCode, Attribute, ErrorCode

  '--Define vars used for printing
  Header$ = "Duplicate Account Listing"
  FF$ = CHR$(12)
  MaxLines = 55

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs

  PrnFile = FREEFILE
  ReportFile$ = "DUPEACCT.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  FOR Cnt = 1 TO NumIdxRecs
    GET AcctIdxFileNum, Cnt, AcctIdx
      AccttoFind$ = AcctIdx.AcctNum
      QPrintRC "Searching: " + AccttoFind$, 8, 17, -1
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
    PRINT #PrnFile, "No duplicate account numbers found."
  END IF
  PRINT #PrnFile, FF$
  CLOSE
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
  KILL ReportFile$
  ShowCursor

EXIT SUB

LogDupe:
   PRINT #PrnFile, AccttoFind$, AcctIdx.RecNum
RETURN

END SUB

SUB TransWrk

  SHARED Trans AS GLTransRecType

  OpenTransFile TransFileNum, NumTrans&

  GET TransFileNum, 26834, Trans
  NewAmt# = Round(Trans.DrAmt) + 72
  Trans.DrAmt = Round(NewAmt#)
  Trans.CrAmt = 0
  PUT TransFileNum, 26834, Trans

  GET TransFileNum, 26857, Trans
  NewAmt# = Round(Trans.CrAmt) + 72
  Trans.DrAmt = 0
  Trans.CrAmt = Round(NewAmt#)
  PUT TransFileNum, 26857, Trans
  
  CLOSE TransFileNum



END SUB

SUB UnMarkRecs

  CLS
  COLOR 11
  Col = 1

  LOCATE 1, 1, 0

  QPrintRC "Un-Mark Records", 1, 1, 10

  QPrintRC "Esc=Cancel Enter=Go", 3, 1, 15
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB

  ok = False
  DO
    CLS
    INPUT "Enter the First Record to Mark  :", FRec$
    INPUT "Enter the Last Record to Mark   :", LRec$

    First& = VAL(FRec$)
    Last& = VAL(LRec$)
    PRINT "OK to Mark "; First&; " to "; Last&; "Esc=Cancel"
    K$ = INPUT$(1)

    IF K$ = CHR$(27) THEN
      QFlag = True
      EXIT DO
    END IF

    IF K$ = "Y" THEN ok = True

  LOOP UNTIL ok

  IF QFlag THEN EXIT SUB

  OpenTransFile TransFileNum, NumTrans&
  IF Last& > NumTrans& THEN
    Last& = NumTrans&
  END IF

  FOR RecNo& = First& TO Last&

    '--Set Flag to False
    GET TransFileNum, RecNo&, Trans
    Trans.Marked = False
    PUT TransFileNum, RecNo&, Trans

    '--Summarize
    RCnt& = RCnt& + 1
    DrAmt# = DrAmt# + Trans.DrAmt
    CrAmt# = CrAmt# + Trans.CrAmt

    LOCATE 4, Col, 0
    PRINT USING "Records Marked: ######"; RCnt&

  NEXT

  CLOSE

  Diff# = Round#(DrAmt#) - Round#(CrAmt#)
  Bal# = ABS(Diff#)

  QPrintRC "File Totals:", 6, 1, 10

  LOCATE 7, Col
  PRINT USING "Total Debits:  #########,.##"; DrAmt#
  LOCATE 8, Col
  PRINT USING "Total Credits: #########,.##"; CrAmt#

  QPrintRC "Press ESC to exit.", 23, 1, 15
  DO UNTIL INKEY$ = CHR$(27): LOOP

END SUB

SUB ViewLog

  LOCATE 1, 1, 0

  IF Exist("GLUTIL.LOG") THEN
     RptTitle$ = "G/L Utility Log"
     RptFileName$ = "GLUTIL.LOG"
     LPTPort = 1
     EntryPoint = 2
     PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint
  ELSE
    CLS
    PRINT "No Log file.  Press any key to continue."
    K$ = INPUT$(1)
  END IF

END SUB

