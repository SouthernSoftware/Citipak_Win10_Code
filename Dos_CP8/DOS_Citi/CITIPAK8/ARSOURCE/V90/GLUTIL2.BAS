DECLARE SUB TransWork ()
DECLARE SUB Copy2NewFile ()
DECLARE SUB ClearTrLink ()
DECLARE SUB ExportMarked ()
DECLARE SUB PostExtFile (PostFileName$)
DECLARE SUB UnMarkTrans ()
DECLARE SUB ViewMarkedTrans ()
DECLARE SUB MarkTrans ()
DEFINT A-Z

'--subs in this module
DECLARE SUB PackAcctFile ()
DECLARE SUB AcctSub ()
DECLARE SUB CarmenOK ()
DECLARE SUB ChangeDate ()
DECLARE SUB DateScan ()
DECLARE SUB DateFixCustom ()
DECLARE FUNCTION FmtDate$ (D$)
DECLARE SUB ReLinkTrans ()
DECLARE SUB ListTrans ()
DECLARE SUB edittrans ()
DECLARE SUB RemoveTrans ()
DECLARE SUB IFLook ()
DECLARE SUB TQuery ()

'--subs in other modules
DECLARE SUB PrintScreenTitle (Title$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'DECLARE SUB Prn2Scrn (FileName$, TopRow, BotRow, clr, PrnFlag)
'DECLARE SUB Prn2Scrn (FileName$, TopRow%, BotRow%, clr, ExitCode)
'DECLARE SUB Prn2Prn (ReportFile$, Title$(), TitLines%, MaxLines%, LPTNo%, exitFlag%)
'DECLARE SUB TextIn (t$, Max%, NumOnly%, CapsOn%, ExitCode%, Colr%)
DECLARE FUNCTION GetBegDate% ()
DECLARE FUNCTION GetEndDate% ()


''$INCLUDE: 'fileview.bi'
'$INCLUDE: 'pro7.bi'
'$INCLUDE: 'gl.bi'
'$INCLUDE: 'glaux.bi'

CONST False = 0, True = NOT False

'Copy2NewFile
'TransWork
END
'appva
'END
'PostExtFile "TR2POST.DAT"
'CLS
'AcctSub
'END
'edittrans
'ClearTrLink


   DO
      CLS
      COLOR 10
      LOCATE 3, 27: PRINT "GL Utility Menu"
      COLOR 11
      LOCATE 5, 20: PRINT "1. List Transaction History File"
      LOCATE 6, 20: PRINT "2. Relink Transaction File"
      LOCATE 7, 20: PRINT "3. Change Transaction Date"
      LOCATE 8, 20: PRINT "4. Scan for Invalid Dates"
      LOCATE 9, 20: PRINT "5. Mark Transactions"
      LOCATE 10, 20: PRINT "6. Un-Mark Transactions"
      LOCATE 11, 20: PRINT "7. View Marked Transactions"
      LOCATE 12, 20: PRINT "8. Delete Marked Transactions"
      LOCATE 13, 20: PRINT "9. View Log File"
      
      COLOR 15
      LOCATE 20, 20: PRINT "Selection No (Esc=Quit) "

      LOCATE 20, 43, 1
      COLOR 14
      K$ = INPUT$(1)

      CLS
      LOCATE 1, 1, 0

      SELECT CASE K$
         CASE "1"
            ListTrans
         CASE "2"
            ReLinkTrans
         CASE "3"
            ChangeDate
            'DateFixCustom
         CASE "4"
            'DateScan
         CASE "5"
            MarkTrans
         CASE "6"
            UnMarkTrans
         CASE "7"
            ViewMarkedTrans
         CASE "8"
            RemoveTrans
         CASE "9"
           IF Exist("GLUTIL.LOG") THEN
              RptTitle$ = "G/L Utility Log"
              RptFileName$ = "GLUTIL.LOG"
              LPTPort = 1
              EntryPoint = 2
              PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint
            END IF
           'ExportMarked
      END SELECT

   LOOP UNTIL K$ = CHR$(27)

END

SUB AcctSub

   'DIM AcctIdx AS GLAcctNdxType
   DIM Acct AS GLAcctRecType

   CLS
   OpenAcctFile GLAcctFile, NumAccts

   'PrnFile = FREEFILE
   'RptFileName$ = "acctlist.prn"
   'OPEN RptFileName$ FOR OUTPUT AS PrnFile

   'AcctNum$ = "10-200-04"
   'AcctRec = FindAcct(AcctNum$)

   FOR Cnt = 1 TO NumAccts
    GET GLAcctFile, Cnt, Acct
      Acct.FrstBTran = 0
      Acct.LastBTran = 0
    PUT GLAcctFile, Cnt, Acct
   
   NEXT

   CLOSE
   EXIT SUB

   
'GET GLAcctFile, AcctRec, Acct
'      'fix for yancey
'      'LOCATE 1, 1: PRINT Acct.Num
'      'IF LEFT$(Acct.Num, 9) = "10-624-74" THEN 'STOP
'      '   Acct.Bgt = 2000
'      '   PUT GLAcctFile, Cnt, Acct
'      'END IF
'
   'NEXT

'   CLOSE
   
'   'Prn2Scrn RptFileName$, 6, 21, 15, PrnFlag%

'   INPUT "Enter a number: ", Num1#
'   INPUT "Amount to subtract from the first number: ", Num2#
'   PRINT "The Answer is: "; Num1# - Num2#
'   PRINT "Duh!"
'
'   QPrintRC "Press any key to continue.", 25, 1, -1
'   K$ = INPUT$(1)

'***************************************************************************
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

'--List account records by index -----------------------
'   OpenGLAcctIdx GLAcctIdx, ActiveAccts
'   OpenAcctFile GLAcctFile, NumAccts
'   FOR Cnt = 1 TO ActiveAccts
'      GET GLAcctIdx, Cnt, AcctIdx
'      GET GLAcctFile, AcctIdx.RecNum, Acct
'      PRINT #PrnFile, Acct.Num; Acct.Title; Acct.BegBal; Acct.Bgt; Acct.ReCalcBal; Acct.FrstTran; Acct.LastTran
'   NEXT
'   CLOSE
   
END SUB

SUB appva

  PrintScreenTitle "Mark Transactions: Special"
  LOCATE 3, 1
  COLOR 15
  PRINT "Press any key to contine. Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  LOCATE 3, 1: PRINT SPACE$(80)

  DIM Trans AS GLTransRecType
  TransRecLen = LEN(Trans)
  OpenTransFile GLTransFile, NumTrans&

  DebitsMarked# = 0
  CreditsMarked# = 0
  
  FOR T& = 1 TO NumTrans&
     GET GLTransFile, T&, Trans
     Mod$ = LEFT$(Trans.Src, 2)

     IF Mod$ = "AP" AND T& < 4239 THEN
       Trans.Marked = -1
       PUT GLTransFile, T&, Trans
       
       '--Running Total of marked transactions
       DebitsMarked# = DebitsMarked# + Trans.DrAmt
       CreditsMarked# = CreditsMarked# + Trans.CrAmt

     END IF


     LOCATE 5, 1
     PRINT USING "Debits  :########,.##"; DebitsMarked#
     PRINT USING "Credits :########,.##"; CreditsMarked#

  NEXT

  PRINT
  PRINT "Press Esc to exit. Any other key to continue."
  Q$ = INPUT$(1)

CLOSE




END SUB

SUB ChangeDate

   CLS
   PrintScreenTitle "Change Transaction Date"
   LOCATE 3, 1
   COLOR 15
   PRINT "Press any key to contine. Esc=Cancel"
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN EXIT SUB
   LOCATE 3, 1: PRINT SPACE$(80)
   COLOR 11
   LOCATE 4, 1

   INPUT "Enter the old date to change: (##/##/##) ", OldDate$
   INPUT "Enter the new date:           (##/##/##) ", NewDate$
   INPUT "Enter the Module:    (AR,GJ,AP,PR,CD,CR) ", M$
   PRINT
   PRINT "Press Esc to Cancel, any other key to continue."
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN
      EXIT SUB
   END IF
   
   ODate = Date2Num(OldDate$)
   NDate = Date2Num(NewDate$)
   Mod$ = UCASE$(M$)
   Changed = 0

   DIM Trans AS GLTransRecType
   OpenTransFile GLTransFile, NumTrans&

   '--open a report file to print to
   RptFile = FREEFILE
   RptFileName$ = "glutil.log"
   OPEN RptFileName$ FOR APPEND AS RptFile
   PRINT #RptFile,
   PRINT #RptFile, "Date change procedure ran on " + DATE$ + " @ " + TIME$
   PRINT #RptFile, "Transactions changed:"

   FOR TrRec& = 1 TO NumTrans&
      Done# = (TrRec& / NumTrans&) * 100
      LOCATE 8, 1
      PRINT USING "Scanning transaction file. ###% complete.         "; Done#
      GET GLTransFile, TrRec&, Trans
      IF Trans.TrDate = ODate THEN
         TMod$ = LEFT$(Trans.Src, 2)
         IF TMod$ = Mod$ THEN
            GOSUB LogDateChange
         END IF
      END IF
   NEXT
   
   IF Changed > 0 THEN
      PRINT #RptFile, "The dates of the above transactions were changed to " + Num2Date(NDate)
      PRINT #RptFile, "Trans Changed  : "; FUsing$(STR$(Changed), "#########")
      PRINT #RptFile, "Total Debits   : "; FUsing$(STR$(Debits#), "########,.##")
      PRINT #RptFile, "Total Credits  : "; FUsing$(STR$(Credits#), "########,.##")
      PRINT #RptFile,
   ELSE
      PRINT #RptFile, "No transactions were changed."
   END IF
   
   CLOSE

   LOCATE 10, 1
   PRINT "Total Debits   : "; FUsing$(STR$(Debits#), "########,.##")
   PRINT "Total Credits  : "; FUsing$(STR$(Credits#), "########,.##")
   PRINT "Press any key to continue."
   K$ = INPUT$(1)
   
   'RptTitle$ = "Date Change Log"
   'EntryPoint = 2
   'PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint
   
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
   LOCATE 9, 1
   PRINT USING "Transaction Dates Changed: #####"; Changed

RETURN


END SUB

SUB ClearTrLink

   DIM Trans AS GLTransRecType

   OpenTransFile GLTransFile, NumTrans&

   '-Set the pointers in the transaction file to zero
   FOR TCnt& = 1 TO NumTrans&
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 3, 1, 0
      PRINT USING "Initializing transaction file. ###% complete."; Done#
      GET GLTransFile, TCnt&, Trans
      Trans.NextTran = 0
      PUT GLTransFile, TCnt&, Trans
   NEXT

   CLOSE



END SUB

SUB Copy2NewFile

  CLS
  PrintScreenTitle "Copying Transactions"
  PRINT "Press Esc to cancel or any other key to continue."
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN
    EXIT SUB
  END IF
  
  DIM Tr AS GLTransRecType
  DIM NTr AS GLTransRecType

  LDate = Date2Num("10/01/96")
  HDate = Date2Num("10/31/96")
  Mod$ = "AP"

  '--open a new GLtrans.file
  NTrRecLen = LEN(NTr)
  NGLTransFile = FREEFILE
  OPEN "gltrans.exp" FOR RANDOM ACCESS READ WRITE SHARED AS NGLTransFile LEN = NTrRecLen '85
  NumNewTrans& = LOF(NGLTransFile) \ NTrRecLen

  '--open a log file to list transactions removed
  LogFile = FREEFILE
  OPEN "glutil.log" FOR APPEND AS LogFile
  PRINT #LogFile,
  PRINT #LogFile, "Copying began @ " + DATE$ + " " + TIME$

  OpenTransFile GLTransFile, NumTrans&

  FOR TrRec& = 1 TO NumTrans&
    GET GLTransFile, TrRec&, Tr

    Complete# = (TrRec& / NumTrans&) * 100
    LOCATE 3, 1: PRINT USING "Processing. Please wait. ###% Complete."; Complete#

    IF Tr.TrDate >= LDate AND Tr.TrDate <= HDate THEN
      IF LEFT$(Tr.Src, 2) = Mod$ THEN
        '--copy transactions to the new file
        
        NTr.AcctNum = Tr.AcctNum
        NTr.TrDate = Tr.TrDate
        NTr.Desc = Tr.Desc
        NTr.CrAmt = Tr.CrAmt
        NTr.DrAmt = Tr.DrAmt
        NTr.Ref = Tr.Ref
        NTr.Src = Tr.Src
        NTr.NextTran = Tr.NextTran
        PUT NGLTransFile, NewTrans&, NTr

        NewTrans& = NewTrans& + 1
        Debits# = Debits# + Tr.DrAmt
        Credits# = Credits# + Tr.CrAmt

      END IF
    END IF

  NEXT

  'screen
  LOCATE 4, 1
  PRINT "Transactions moved: "; NewTrans&
  PRINT "Total Debits moved: " + STR$(Debits#)
  PRINT "Total Credits moved: " + STR$(Credits#)

  'log
  PRINT #LogFile,
  PRINT #LogFile, "Transactions exported: "; NewTrans&
  PRINT #LogFile, "Total Debits  : " + STR$(DebitsGone#)
  PRINT #LogFile, "Total Credits : " + STR$(CreditsGone#)
  PRINT #LogFile,

  CLOSE

  PRINT "Press any key to continue."
  K$ = INPUT$(1)
  CLS
  
END SUB

SUB DateFixCustom

   COLOR 15
   CLS
   PrintScreenTitle "Custom Date Fix"

   '--open a report file to print to
   RptFile = FREEFILE
   RptFileName$ = "TRLIST.PRN"
   OPEN RptFileName$ FOR OUTPUT AS RptFile
   
   DIM Trans AS GLTransRecType
   OpenTransFile GLTransFile, NumTrans&
   Rec$ = "       "

   FOR TrRec& = 1 TO NumTrans&
      LSET Rec$ = STR$(TrRec&)
      QPrintRC "Processing Record" + Rec$, 4, 1, 15
      GET GLTransFile, TrRec&, Trans

      'IF Trans.TrDate >= Date2Num("09/01/96") THEN
      '  Mo$ = LEFT$(Num2Date(Trans.TrDate), 2)
      '  Mth = VAL(Mo$)
      '  IF Mth >= 7 AND Mth <= 12 THEN
      '    NewYr$ = "95"
      '  ELSE
      '    NewYr$ = "96"
      '  END IF
      '  ND$ = LEFT$(Num2Date(Trans.TrDate), 8) + NewYr$
      '  Cnt = Cnt + 1
      '  Trans.TrDate = Date2Num(ND$)
      '  PUT GLTransFile, TrRec&, Trans
      'END IF

      IF QPTrim$(Trans.Desc) = "From Utility" THEN
        GOSUB PrintDetail
      END IF

  NEXT

  'LOCATE 5, 1
  'PRINT Cnt; "Dates Changed"
  'PRINT "Press any key to continue"
  'K$ = INPUT$(1)

  PRINT #RptFile, USING "Total Debits:  ########,.## "; TotDr#
  PRINT #RptFile, USING "Total Credits: ########,.## "; TotCr#
  CLOSE

  RptTitle$ = "Transaction Listing"
  EntryPoint = 2
  PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint



EXIT SUB


PrintDetail:
  ToPrint$ = SPACE$(132)
  LSET ToPrint$ = Trans.AcctNum
  MID$(ToPrint$, 13) = Num2Date(Trans.TrDate)
  MID$(ToPrint$, 25) = LEFT$(Trans.Desc, 13)
  MID$(ToPrint$, 38) = Trans.Ref
  MID$(ToPrint$, 48) = FUsing$(STR$(Trans.DrAmt), "########,.##")
  MID$(ToPrint$, 61) = FUsing$(STR$(Trans.CrAmt), "########,.##")
  MID$(ToPrint$, 74) = Trans.Src
  PRINT #RptFile, ToPrint$
  TotDr# = TotDr# + Trans.DrAmt
  TotCr# = TotCr# + Trans.CrAmt

RETURN






END SUB

SUB edittrans

DIM Trans AS GLTransRecType
OpenTransFile GLTransFile, NumTrans&

GET GLTransFile, 34852, Trans
Trans.DrAmt = 0
Trans.CrAmt = 0
PUT GLTransFile, 34852, Trans

GET GLTransFile, 34869, Trans
Trans.DrAmt = 0
Trans.CrAmt = 0
PUT GLTransFile, 34859, Trans

CLOSE
END SUB

SUB ExportMarked

  CLS
  PrintScreenTitle "Exporting Marked Transactions"
  COLOR 15

  DIM Tr AS GLTransRecType
  DIM NTr AS GLTransRecType

  '--open a new GLtrans.file
  NTrRecLen = LEN(NTr)
  NGLTransFile = FREEFILE
  OPEN "gltrans.exp" FOR RANDOM ACCESS READ WRITE SHARED AS NGLTransFile LEN = NTrRecLen '85
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

    IF Tr.Marked = -1 THEN
      '--copy marked records to the new file
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
    'ELSE
    '  '-check em out
    '  ToPrint$ = SPACE$(80)
    '  LSET ToPrint$ = Tr.AcctNum
    '  MID$(ToPrint$, 13) = ConvDateStr1$(Num2Date(Tr.TrDate))
    '  MID$(ToPrint$, 24) = LEFT$(Tr.Desc, 19)
    '  MID$(ToPrint$, 42) = STR$(Tr.CrAmt)
    '  MID$(ToPrint$, 52) = STR$(Tr.DrAmt)
    '  MID$(ToPrint$, 62) = Tr.Ref
    '  MID$(ToPrint$, 72) = Tr.Src
    '  PRINT #LogFile, ToPrint$
    '
    '  '--Keep track of what's gone
    '  Gone& = Gone& + 1
    '  DebitsGone# = DebitsGone# + Tr.DrAmt
    '  CreditsGone# = CreditsGone# + Tr.CrAmt
    '
    END IF

  NEXT

  'screen
  'LOCATE 4, 1
  'PRINT "Transactions removed: "; Gone&
  'PRINT "Total Debits removed: " + STR$(DebitsGone#)
  'PRINT "Total Credits removed: " + STR$(CreditsGone#)

  'log
  PRINT #LogFile,
  PRINT #LogFile, "Transactions removed: "; Gone&
  PRINT #LogFile, "Total Debits removed: " + STR$(DebitsGone#)
  PRINT #LogFile, "Total Credits removed: " + STR$(CreditsGone#)
  PRINT #LogFile,

  CLOSE

  'NAME "gltrans.dat" AS "gltrans.old"
  'NAME "ngltrans.dat" AS "gltrans.dat"

  PRINT "Press any key to continue."
  K$ = INPUT$(1)
  CLS

  'ReLinkTrans



END SUB

SUB ListTrans

   PrintScreenTitle "Transaction Listing"
   LOCATE 3, 1
   COLOR 15
   PRINT "(C)reate New Report, (V)iew Current Report,  Esc=Cancel"
   K$ = INPUT$(1)
   SELECT CASE K$
    CASE CHR$(27)
      EXIT SUB
    CASE "V", "v"
      IF Exist("trlist.prn") THEN
        RptFileName$ = "trlist.prn"
        GOTO PrintRpt
      END IF
   END SELECT
   LOCATE 3, 1: PRINT SPACE$(80)

   DIM Trans AS GLTransRecType
   OpenTransFile GLTransFile, NumTrans&

   '--open a report file to print to
   RptFile = FREEFILE
   RptFileName$ = "TRLIST.PRN"
   OPEN RptFileName$ FOR OUTPUT AS RptFile
   
   FOR TrRec& = 1 TO NumTrans&
      Complete! = (TrRec& / NumTrans&) * 100
      LOCATE 3, 1
      PRINT USING "Processing... ###% Complete."; Complete!

      GET GLTransFile, TrRec&, Trans

      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = Trans.AcctNum
      MID$(ToPrint$, 13) = Num2Date(Trans.TrDate)
      MID$(ToPrint$, 25) = LEFT$(Trans.Desc, 13)
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

   NEXT
   
   PRINT #RptFile, "Total Debits   : "; FUsing$(STR$(Debits#), "##########,.##")
   PRINT #RptFile, "Total Credits  : "; FUsing$(STR$(Credits#), "##########,.##")
   PRINT #RptFile, "Difference     : "; FUsing$(STR$(Diff#), "##########,.##")

   CLOSE

PrintRpt:
   RptTitle$ = "Transaction Listing"
   EntryPoint = 2
   PrintRptFile RptTitle$, RptFileName$, LPTPort%, RetCode%, EntryPoint

END SUB

SUB MarkTrans

  PrintScreenTitle "Mark Transactions"
  LOCATE 3, 1
  COLOR 15
  PRINT "Press any key to contine. Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  LOCATE 3, 1: PRINT SPACE$(80)

  DIM Trans AS GLTransRecType
  TransRecLen = LEN(Trans)
  OpenTransFile GLTransFile, NumTrans&

  DebitsMarked# = 0
  CreditsMarked# = 0

  DO
Top:
     LowRec& = 0
     HighRec& = 0

     CLS
     LOCATE 3, 1
     PRINT "Mark Transaction Records"
     INPUT "FIRST Record :", L$: LowRec& = VAL(L$)
     INPUT "LAST Record  :", H$: HighRec& = VAL(H$)
     PRINT "OK to Mark?"
     PRINT "Y=Yes, Any other key=No, Esc=Quit"
     K$ = UCASE$(INPUT$(1))

     IF K$ = CHR$(27) THEN EXIT SUB
     IF K$ <> "Y" THEN GOTO Top
     IF LowRec& <= 0 OR HighRec& <= 0 THEN EXIT DO
     IF HighRec& < LowRec& THEN EXIT DO
     IF HighRec& > NumTrans& THEN HighRec& = NumTrans&

     LOCATE 4, 1
     PRINT "                                       "
     PRINT "Marked Totals                          "

     FOR TrRec& = LowRec& TO HighRec&
        GET GLTransFile, TrRec&, Trans
        Trans.Marked = -1
        PUT GLTransFile, TrRec&, Trans

        '--Running Total of marked transactions
        DebitsMarked# = DebitsMarked# + Trans.DrAmt
        CreditsMarked# = CreditsMarked# + Trans.CrAmt

        LOCATE 5, 1
        PRINT USING "Debits  :########,.##"; DebitsMarked#
        PRINT USING "Credits :########,.##"; CreditsMarked#

     NEXT

     PRINT
     PRINT "Press Esc to exit. Any other key to continue."
     Q$ = INPUT$(1)
     DebitsMarked# = 0
     CreditsMarked# = 0

  LOOP UNTIL Q$ = CHR$(27)

CLOSE


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
'   DIM Acct AS GLAcctRecType
'   DIM NAcct AS GLAcctRecType
'
'   AcctRecLen = LEN(Acct)
'
'   ESCKey = 27
'   F10Key = 68
'
'   OpenAcctFile GLAcctFile%, NumAccts%
'
'   NewAcctFile = FREEFILE
'   OPEN "GLacct.new" FOR RANDOM AS NewAcctFile LEN = AcctRecLen
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
'   NAME "GLacct.dat" AS "GLacct.old"
'   NAME "GLacct.new" AS "GLacct.dat"
'
'   LOCATE 23, 1: COLOR 14
'   PRINT "Press any key to exit."
'   K$ = INPUT$(1)
'
'   COLOR 7
'   CLS
'
END SUB

SUB PostExtFile (PostFileName$)

  CLS
  PRINT "Ready to Post: " + PostFileName$
  PRINT "Press <Esc> to Cancel, any other key to continue."
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB

  Post2GL PostFileName$, BadTrans
  IF BadTrans > 0 THEN
    LOCATE 3, 1
    PRINT "Bad Transactions: "; BadTrans
  ELSE
    LOCATE 3, 1
    PRINT "Posting Successful."
  END IF

  PRINT "Press any key to exit."
  K$ = INPUT$(1)

END SUB

SUB PrintScreenTitle (Title$)

TLen = LEN(Title$)
Col = (80 - TLen) \ 2
TColor = 11
TLineColor = 10

QPrintRC Title$, 1, Col, TColor
QPrintRC STRING$(80, CHR$(196)), 2, 1, TLineColor


END SUB

SUB ReLinkTrans

   PrintScreenTitle "Relink Accounting Database"
   LOCATE 3, 1
   COLOR 15
   PRINT "Press any key to contine. Esc=Cancel"
   K$ = INPUT$(1)
   IF K$ = CHR$(27) THEN EXIT SUB
   LOCATE 3, 1: PRINT SPACE$(80)

   DIM Trans AS GLTransRecType
   DIM Acct AS GLAcctRecType

   OpenTransFile GLTransFile, NumTrans&
   OpenAcctFile GLAcctFile, NumAccts

   LOCK GLTransFile
   LOCK GLAcctFile

   LogFile = FREEFILE
   OPEN "glutil.log" FOR APPEND AS #LogFile
   PRINT #LogFile,
   PRINT #LogFile, "Accounting Database relink started @ " + DATE$ + " @ "; TIME$

   COLOR 15
   '-Set the pointers in the transaction file to zero
   FOR TCnt& = 1 TO NumTrans&
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 3, 1, 0
      PRINT USING "Initializing transaction file. ###% complete."; Done#
      GET GLTransFile, TCnt&, Trans
      Trans.NextTran = 0
      PUT GLTransFile, TCnt&, Trans
   NEXT

   LOCATE 3, 1, 0
   PRINT STRING$(80, 32)

   '-Set the pointers in the account file to zero
   FOR Cnt = 1 TO NumAccts
      Done# = (Cnt / NumAccts) * 100
      LOCATE 3, 1, 0
      PRINT USING "Initializing account file. ###% complete."; Done#
      GET GLAcctFile, Cnt, Acct
      Acct.FrstTran = 0
      Acct.LastTran = 0
      PUT GLAcctFile, Cnt, Acct
   NEXT

   LOCATE 3, 1, 0
   PRINT STRING$(80, 32)

   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 3, 1, 0
      PRINT USING "Relinking database. ###% complete."; Done#

      GET GLTransFile, TCnt&, Trans

      '-Find the record number of the account
      AcctRecNum = FindAcct(Trans.AcctNum)

      '-If we find the account
      IF AcctRecNum > 0 THEN
         GET GLAcctFile, AcctRecNum, Acct

         '-Check out the pointer to the first transaction
         SELECT CASE Acct.FrstTran

           '-If this is the first transaction for this account
           CASE 0
               '-Set first and last pointers to this transaction
               Acct.FrstTran = TCnt&
               Acct.LastTran = TCnt&
               PUT GLAcctFile, AcctRecNum, Acct

            '-If there are already transactions for this account
            CASE IS > 0
               '-Remember the pointer to the last transaction.
               Prev& = Acct.LastTran

               '-Set the last trans pointer to this transaction
               Acct.LastTran = TCnt&
               PUT GLAcctFile, AcctRecNum, Acct

               '-Get the last previous transaction and set its
               '-next tran pointer to this transaction
               GET GLTransFile, Prev&, Trans
               Trans.NextTran = TCnt&
               PUT GLTransFile, Prev&, Trans

               'update running balance here

         END SELECT
      
      ELSE  '-could not find the account
         BadTran = BadTran + 1
         LOCATE 13, 1
         PRINT USING "Orphaned transactions: #####"; BadTran

         '-Keep a list of orphaned transactions.
         GOSUB Logit

      END IF

   NEXT

   '-we're done
   UNLOCK GLTransFile
   UNLOCK GLAcctFile

   '-Tell user we're done.
   IF BadTran > 0 THEN
      '-Errors in trans file
      PRINT #LogFile, "Relink encountered ophans. Completed @ " + DATE$ + " @ " + TIME$
      QPrintRC "There were errors in transaction file!", 4, 1, 15
      QPrintRC "Orphaned transactions listed in LINK.LOG file.", 5, 1, 15
      QPrintRC "Press any key to continue.", 6, 1, 15
      K$ = INPUT$(1)
   ELSE
      '-No errors in trans file
      PRINT #LogFile, "Relink of Accounting Databases successful. " + DATE$ + " @ " + TIME$
      QPrintRC "Accounting database relink successful.", 3, 1, 15
      QPrintRC "Press any key to continue.", 4, 1, 15
      K$ = INPUT$(1)
   END IF

   CLOSE

EXIT SUB

Logit:
   ToPrint$ = SPACE$(132)
   LSET ToPrint$ = Trans.AcctNum
   MID$(ToPrint$, 18) = Num2Date(Trans.TrDate)
   MID$(ToPrint$, 30) = LEFT$(Trans.Desc, 15)
   MID$(ToPrint$, 50) = Trans.Ref
   MID$(ToPrint$, 60) = FUsing(STR$(Trans.DrAmt), "#######,.##")
   MID$(ToPrint$, 70) = FUsing(STR$(Trans.CrAmt), "#######,.##")
   MID$(ToPrint$, 80) = "Record:" + STR$(TCnt&)
   PRINT #LogFile, ToPrint$
RETURN


END SUB

SUB RemoveTrans

  CLS
  PrintScreenTitle "Delete Marked Transactions"
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
    PrintScreenTitle "Delete Marked Transactions"
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
  PrintScreenTitle "Deleting Marked Transactions"
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

SUB TransWork

   COLOR 15
   CLS
   PrintScreenTitle "Transaction Work"

   'BegDate = GetBegDate%
   'EndDate = GetEndDate%

   DIM Trans AS GLTransRecType
   OpenTransFile GLTransFile, NumTrans&
   
   LogFile = FREEFILE
   OPEN "glutil.log" FOR APPEND AS #LogFile
   PRINT #LogFile,
   PRINT #LogFile, "G/L Util" + DATE$ + " @ "; TIME$
   
   FOR TrRec& = 1 TO NumTrans&
      Done# = (TrRec& / NumTrans&) * 100
      LOCATE 3, 1, 0
      PRINT USING "Scanning transaction file. ###% complete."; Done#

      GET GLTransFile, TrRec&, Trans

      'IF Trans.TrDate < BegDate OR Trans.TrDate > EndDate THEN
      IF LEFT$(Trans.Src, 2) = "PR" THEN
        
        TCount = TCount + 1
        Trans.Ref = " "
        PUT GLTransFile, TrRec&, Trans
        GOSUB ShowTrans

      END IF

   NEXT
   
   LOCATE 4, 1
   PRINT STRING$(80, " ")
   LOCATE 4, 1

   'IF GoodCnt& = NumTrans& THEN
   '   PRINT "No date problems in transaction file."
   '   PRINT #LogFile, "No date problems in transaction file."
   'END IF
   '
   'IF TCount > 0 THEN
   '   PRINT TCount; "date(s) are out of range."
   '   PRINT #LogFile, TCount; "date(s) are out of range."
   '   PRINT #LogFile, USING "Total Debits  : ########,.##"; TotDebits#
   '   PRINT #LogFile, USING "Total Credits : ########,.##"; TotCredits#
   'END IF
   '
   'PRINT
   PRINT "Press any key to continue."
   K$ = INPUT$(1)

   CLOSE

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

SUB UnMarkTrans

PrintScreenTitle "Un-Mark Transactions"
LOCATE 3, 1
COLOR 15
PRINT "Press any key to contine. Esc=Cancel"
K$ = INPUT$(1)
IF K$ = CHR$(27) THEN EXIT SUB
LOCATE 3, 1: PRINT SPACE$(80)

DIM Trans AS GLTransRecType
TransRecLen = LEN(Trans)
OpenTransFile GLTransFile, NumTrans&

DO
UMTop:
   LowRec& = 0
   HighRec& = 0

   CLS
   LOCATE 3, 1
   PRINT "Un-Mark Transaction Records"
   INPUT "FIRST Record :", L$: LowRec& = VAL(L$)
   INPUT "LAST Record  :", H$: HighRec& = VAL(H$)
   PRINT "OK to UnMark?"
   PRINT "Y=Yes, Any other key=No, Esc=Quit"
   K$ = UCASE$(INPUT$(1))

   IF K$ = CHR$(27) THEN EXIT SUB
   IF K$ <> "Y" THEN GOTO UMTop
   IF LowRec& <= 0 OR HighRec& <= 0 THEN EXIT DO
   IF HighRec& < LowRec& THEN EXIT DO
   IF HighRec& > NumTrans& THEN HighRec& = NumTrans&

   LOCATE 4, 1
   PRINT "                                       "
   PRINT "Un-Marked Totals                          "

   FOR TrRec& = LowRec& TO HighRec&

      GET GLTransFile, TrRec&, Trans
      Trans.Marked = 0
      PUT GLTransFile, TrRec&, Trans

      '--Running Total of marked transactions
      DebitsMarked# = DebitsMarked# + Trans.DrAmt
      CreditsMarked# = CreditsMarked# + Trans.CrAmt

      LOCATE 5, 1
      PRINT USING "Debits  :########,.##"; DebitsMarked#
      PRINT USING "Credits :########,.##"; CreditsMarked#

   NEXT

   PRINT
   PRINT "Press Esc to exit."
   Q$ = INPUT$(1)

LOOP UNTIL Q$ = CHR$(27)

CLOSE


END SUB

SUB ViewMarkedTrans


  PrintScreenTitle "View Marked Transactions"
  LOCATE 3, 1
  COLOR 15
  PRINT "Press any key to contine. Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  LOCATE 3, 1: PRINT SPACE$(80)

   DIM Trans AS GLTransRecType
   OpenTransFile GLTransFile, NumTrans&

   '--open a report file to print to
   RptFile = FREEFILE
   TransRptFileName$ = "TRLIST.PRN"
   OPEN TransRptFileName$ FOR OUTPUT AS RptFile

   FOR TrRec& = 1 TO NumTrans&

      Complete! = (TrRec& / NumTrans&) * 100
      LOCATE 3, 1
      PRINT USING "Processing... ###% Complete."; Complete!

      GET GLTransFile, TrRec&, Trans

      IF Trans.Marked = -1 THEN
         ToPrint$ = SPACE$(132)
         LSET ToPrint$ = Trans.AcctNum
         MID$(ToPrint$, 16) = ConvDateStr1$(Num2Date(Trans.TrDate))
         MID$(ToPrint$, 25) = LEFT$(Trans.Desc, 12)
         MID$(ToPrint$, 38) = Trans.Ref
         MID$(ToPrint$, 47) = FUsing$(STR$(Trans.DrAmt), "########,.##")
         MID$(ToPrint$, 59) = FUsing$(STR$(Trans.CrAmt), "########,.##")
         MID$(ToPrint$, 72) = Trans.Src
         MID$(ToPrint$, 82) = "T#:" + STR$(TrRec&)
         MID$(ToPrint$, 94) = "Nx:" + STR$(Trans.NextTran)
         PRINT #RptFile, ToPrint$

         Count = Count + 1
         Debits# = Debits# + Trans.DrAmt
         Credits# = Credits# + Trans.CrAmt
      END IF

   NEXT

   PRINT #RptFile, "Marked Trans   : "; FUsing$(STR$(Count), "##########")
   PRINT #RptFile, "Total Debits   : "; FUsing$(STR$(Debits#), "##########,.##")
   PRINT #RptFile, "Total Credits  : "; FUsing$(STR$(Credits#), "##########,.##")

   CLOSE

   RptTitle$ = "Marked Transactions"
   EntryPoint = 2
   PrintRptFile RptTitle$, TransRptFileName$, LPTPort%, RetCode%, EntryPoint
   
END SUB

SUB ViewMarkTrans

   COLOR 15
   CLS

   PrintScreenTitle "View Marked Transactions"

   DIM Trans AS GLTransRecType
   OpenTransFile GLTransFile, NumTrans&

   '--open a report file to print to
   RptFile = FREEFILE
   TransRptFileName$ = "TRLIST.PRN"
   OPEN TransRptFileName$ FOR OUTPUT AS RptFile

   FOR TrRec& = 1 TO NumTrans&

      Complete! = (TrRec& / NumTrans&) * 100
      LOCATE 3, 1
      PRINT USING "Processing... ###% Complete."; Complete!

      GET GLTransFile, TrRec&, Trans

      IF Trans.Marked = -1 THEN
         ToPrint$ = SPACE$(132)
         LSET ToPrint$ = Trans.AcctNum
         MID$(ToPrint$, 11) = Num2Date(Trans.TrDate)
         MID$(ToPrint$, 20) = LEFT$(Trans.Desc, 15)
         MID$(ToPrint$, 36) = Trans.Ref
         MID$(ToPrint$, 46) = FUsing$(STR$(Trans.DrAmt), "########,.##")
         MID$(ToPrint$, 59) = FUsing$(STR$(Trans.CrAmt), "########,.##")
         MID$(ToPrint$, 72) = Trans.Src
         MID$(ToPrint$, 82) = "T#:" + STR$(TrRec&)
         MID$(ToPrint$, 94) = "Nx:" + STR$(Trans.NextTran)
         PRINT #RptFile, ToPrint$

         Count = Count + 1
         Debits# = Debits# + Trans.DrAmt
         Credits# = Credits# + Trans.CrAmt
      END IF

   NEXT

   PRINT #RptFile, "Marked Trans   : "; FUsing$(STR$(Count), "##########")
   PRINT #RptFile, "Total Debits   : "; FUsing$(STR$(Debits#), "##########,.##")
   PRINT #RptFile, "Total Credits  : "; FUsing$(STR$(Credits#), "##########,.##")

   CLOSE

   CLS
   COLOR 11

   LOCATE 2, 1: PRINT "Marked Transaction Totals"
   LOCATE 4, 1: PRINT "Acct      Date     Description     Ref              Debit       Credit  PostRef"
   LOCATE 5, 1: PRINT STRING$(80, CHR$(196))

   LOCATE 23, 1: PRINT "Esc to exit. F7 to Print."

   'Prn2Scrn TransRptFileName$, 6, 21, 15, PrnFlag%
   'IF PrnFlag = 1 THEN GOSUB PrintHist
   'ConvDateStr1$

END SUB

