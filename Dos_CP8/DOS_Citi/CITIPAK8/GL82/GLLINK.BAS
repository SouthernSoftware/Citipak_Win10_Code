DEFINT A-Z

'--subs in this module
DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()

'--subs in other modules
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)

'$INCLUDE: 'pro7.bi'
'$INCLUDE: 'gl.bi'
'$INCLUDE: 'glaux.bi'

CONST False = 0, True = NOT False

'ReLinkTrans

SUB RelinkBgtTrans

   COLOR 15

   LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
   QPrintRC "Relink Budget Transaction Database", 5, 10, 15
   PrintHelp "System Operations.  Please wait."
   
   DIM BTrans AS GLTransRecType
   DIM Acct AS GLAcctRecType

   TransRecLen = LEN(BTrans)
   BgtTransFile = FREEFILE
   OPEN "BGTTRANS.DAT" FOR RANDOM AS BgtTransFile LEN = TransRecLen
   NumTrans& = LOF(BgtTransFile) \ TransRecLen
   
   OpenAcctFile GLAcctFile, NumAccts

   LOCK BgtTransFile
   LOCK GLAcctFile

   LogFile = FREEFILE
   LogFileName$ = "GLLINK.LOG"
   OPEN LogFileName$ FOR APPEND AS #LogFile
   PRINT #LogFile,
   PRINT #LogFile, "Budget Database relink started @ " + DATE$ + " @ "; TIME$

   '-Set the pointers in the transaction file to zero
   FOR TCnt& = 1 TO NumTrans&
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Initializing transaction file. ###% complete."; Done#
      GET BgtTransFile, TCnt&, BTrans
      BTrans.NextTran = 0
      PUT BgtTransFile, TCnt&, BTrans
   NEXT

   LOCATE 6, 10, 0
   PRINT STRING$(61, 32)

   '-Set the budget pointers in the account file to zero
   FOR Cnt = 1 TO NumAccts
      Done# = (Cnt / NumAccts) * 100
      LOCATE 6, 10, 0
      PRINT USING "Initializing account file. ###% complete."; Done#
      GET GLAcctFile, Cnt, Acct
      Acct.FrstBTran = 0
      Acct.Bgt = 0
      Acct.LastBTran = 0
      PUT GLAcctFile, Cnt, Acct
   NEXT

   LOCATE 6, 10, 0
   PRINT STRING$(61, 32)

   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Relinking database. ###% complete."; Done#

      GET BgtTransFile, TCnt&, BTrans

      '-Find the record number of the account
      AcctRecNum = FindAcct(BTrans.AcctNum)

      '-If we find the account
      IF AcctRecNum > 0 THEN
         GET GLAcctFile, AcctRecNum, Acct
         
         
         '-Check out the pointer to the first transaction
         



         '--update the Acct's Budget Balance
         SELECT CASE Acct.Typ
            CASE "A", "E"
               Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.DrAmt) - Round#(BTrans.CrAmt)
            CASE "L", "R"
               Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.CrAmt) - Round#(BTrans.DrAmt)
         END SELECT
         PUT GLAcctFile, AcctRecNum, Acct

         SELECT CASE Acct.FrstBTran
           '-If this is the first transaction for this account
           CASE 0
               '-Set first and last pointers to this transaction
               Acct.FrstBTran = TCnt&
               Acct.LastBTran = TCnt&
               PUT GLAcctFile, AcctRecNum, Acct

            CASE IS > 0  '-There are already transactions for this account
               '-Remember the pointer to the last transaction.
               Prev& = Acct.LastBTran
               '-Set the last trans pointer to this transaction
               Acct.LastBTran = TCnt&
               PUT GLAcctFile, AcctRecNum, Acct

               '-Get the last previous transaction and set its
               '-next tran pointer to this transaction
               GET BgtTransFile, Prev&, BTrans
               BTrans.NextTran = TCnt&
               PUT BgtTransFile, Prev&, BTrans
            CASE ELSE
         END SELECT


      ELSE  '-could not find the account
         BadTran = BadTran + 1
         LOCATE 14, 10
         PRINT USING "Orphaned transactions: #####"; BadTran
         GOSUB LogBgtTrans '-Keep a list of orphaned transactions.

      END IF
   NEXT

   '-we're done
   UNLOCK BgtTransFile
   UNLOCK GLAcctFile

   '-Tell user we're done.
   IF BadTran > 0 THEN
      '-Errors in trans file
      PRINT #LogFile, "Relink encountered ophans. Completed @ " + DATE$ + " @ " + TIME$
      'QPrintRC "There were errors in transaction file!", 7, 10, 15
      'QPrintRC "Orphaned transactions listed in LINK.LOG file.", 8, 10, 15
      'QPrintRC "Press any key to continue.", 9, 10, 15
      'K$ = INPUT$(1)
   ELSE
      '-No errors in trans file
      PRINT #LogFile, "Relink of Budget Database successful. " + DATE$ + " @ " + TIME$
      'QPrintRC "Accounting database relink successful.", 7, 10, 15
      'QPrintRC "Press any key to continue.", 7, 10, 15
      'K$ = INPUT$(1)
   END IF

   CLOSE

EXIT SUB

LogBgtTrans:
   ToPrint$ = SPACE$(132)
   LSET ToPrint$ = BTrans.AcctNum
   MID$(ToPrint$, 18) = Num2Date(BTrans.TrDate)
   MID$(ToPrint$, 30) = LEFT$(BTrans.Desc, 15)
   MID$(ToPrint$, 50) = BTrans.Ref
   MID$(ToPrint$, 60) = FUsing(STR$(BTrans.DrAmt), "#######,.##")
   MID$(ToPrint$, 70) = FUsing(STR$(BTrans.CrAmt), "#######,.##")
   MID$(ToPrint$, 80) = "Record:" + STR$(TCnt&)
   PRINT #LogFile, ToPrint$
RETURN


END SUB

SUB ReLinkPOTrans

   COLOR 15

   LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
   QPrintRC "Relink Purchase Order Transaction Database", 5, 10, 15
   PrintHelp "System Operations.  Please wait."

   DIM Acct AS GLAcctRecType

   OpenAcctFile GLAcctFile, NumAccts

   DIM POTrans AS GLTransRecType
   TransRecLen = LEN(POTrans)
   POTransFile = FREEFILE
   OPEN "POTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS POTransFile LEN = TransRecLen
   NumTrans& = LOF(POTransFile) \ TransRecLen
   
   'LOCK POTransFile
   'LOCK GLAcctFile

   LogFile = FREEFILE
   LogFileName$ = "GLLINK.LOG"
   OPEN LogFileName$ FOR APPEND AS #LogFile

   PRINT #LogFile,
   PRINT #LogFile, "Accounting Database relink started @ " + DATE$ + " @ "; TIME$

   '-Set the pointers in the transaction file to zero
   FOR TCnt& = 1 TO NumTrans&
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Initializing transaction file. ###% complete."; Done#
      GET POTransFile, TCnt&, POTrans
      POTrans.NextTran = 0
      PUT POTransFile, TCnt&, POTrans
   NEXT

   LOCATE 6, 10, 0
   PRINT STRING$(61, 32)

   '-Set the po pointers in the account file to zero
   FOR Cnt = 1 TO NumAccts
      Done# = (Cnt / NumAccts) * 100
      LOCATE 6, 10, 0
      PRINT USING "Initializing account file. ###% complete."; Done#
      GET GLAcctFile, Cnt, Acct
      Acct.FrstPTran = 0
      Acct.Encumb = 0
      Acct.LastPTran = 0
      PUT GLAcctFile, Cnt, Acct
   NEXT

   LOCATE 6, 10, 0
   PRINT STRING$(61, 32)

   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Relinking. ###% complete."; Done#

      'IF TCnt& = 11881 THEN STOP

      GET POTransFile, TCnt&, POTrans

      '-Find the record number of the account
      AcctRecNum = FindAcct(POTrans.AcctNum)

      '-If we find the account
      IF AcctRecNum > 0 THEN
         GET GLAcctFile, AcctRecNum, Acct

         '--update running encumbrance balance here
         SELECT CASE Acct.Typ
          CASE "E", "A"
            Acct.Encumb = Round#(Acct.Encumb) + Round#(POTrans.DrAmt) - Round#(POTrans.CrAmt)
          CASE "L", "R"
            Acct.Encumb = Round#(Acct.Encumb) + Round#(POTrans.CrAmt) - Round#(POTrans.DrAmt)
         END SELECT
         PUT GLAcctFile, AcctRecNum, Acct
         
         '-Check out the pointer to the first transaction
         SELECT CASE Acct.FrstPTran

           '-If this is the first transaction for this account
           CASE 0
               '-Set first and last pointers to this transaction
               Acct.FrstPTran = TCnt&
               Acct.LastPTran = TCnt&
               PUT GLAcctFile, AcctRecNum, Acct

            '-If there are already transactions for this account
            CASE IS > 0
               '-Remember the pointer to the last transaction.
               Prev& = Acct.LastPTran

               '-Set the last trans pointer to this transaction
               Acct.LastPTran = TCnt&
               PUT GLAcctFile, AcctRecNum, Acct

               '-Get the last previous transaction and set its
               '-next tran pointer to this transaction
               GET POTransFile, Prev&, POTrans
               POTrans.NextTran = TCnt&
               PUT POTransFile, Prev&, POTrans
         
         END SELECT
         
      ELSE  '-could not find the account
         BadTran = BadTran + 1
         LOCATE 13, 10
         PRINT USING "Orphaned transactions: #####"; BadTran

         '-Keep a list of orphaned transactions.
         GOSUB LogPO

      END IF

   NEXT

   '-we're done here
   'UNLOCK POTransFile
   'UNLOCK GLAcctFile

   IF BadTran > 0 THEN
      '-Errors in trans file
      PRINT #LogFile,
      PRINT #LogFile, "Orphan Transaction Totals:";
      PRINT #LogFile, TAB(58); FUsing(STR$(BadDebits#), "#######,.##");
      PRINT #LogFile, TAB(70); FUsing(STR$(BadCredits#), "#######,.##")
      PRINT #LogFile, "Relink completed @ " + DATE$ + " @ " + TIME$
      PRINT #LogFile, "Orphan transactions encountered! Call Customer Support."
   ELSE
      '-No errors in trans file
      PRINT #LogFile, "Relink of PO Transaction successful. " + DATE$ + " @ " + TIME$
   END IF

   CLOSE

   '-Tell user we're done.
   IF BadTran > 0 THEN
      '-Errors in trans file
      QPrintRC "Press any key to view log.", 14, 10, 15
      K$ = INPUT$(1)
      PrintRptFile "Link Log", LogFileName$, LPTNo, RetCode%, 2
   ELSE
      '-No errors in trans file
      QPrintRC "Purchase Order transaction relink successful.", 6, 10, 15
      QPrintRC "Press any key to continue.", 7, 10, 15
      K$ = INPUT$(1)
   END IF

EXIT SUB

LogPO:
   ToPrint$ = SPACE$(132)
   LSET ToPrint$ = POTrans.AcctNum
   MID$(ToPrint$, 18) = Num2Date(POTrans.TrDate)
   MID$(ToPrint$, 30) = LEFT$(POTrans.Desc, 15)
   MID$(ToPrint$, 46) = POTrans.Ref
   MID$(ToPrint$, 58) = FUsing(STR$(POTrans.DrAmt), "#######,.##")
   MID$(ToPrint$, 70) = FUsing(STR$(POTrans.CrAmt), "#######,.##")
   MID$(ToPrint$, 85) = "Record:" + STR$(TCnt&)
   PRINT #LogFile, ToPrint$
   BadDebits# = BadDebits# + POTrans.DrAmt
   BadCredits# = BadCredits# + POTrans.CrAmt
RETURN



END SUB

SUB ReLinkTrans

   'PrintScreenTitle "Relink Accounting Database"
   'LOCATE 3, 1
   'COLOR 15
   'PRINT "Press any key to contine. Esc=Cancel"
   'K$ = INPUT$(1)
   'IF K$ = CHR$(27) THEN EXIT SUB
   'LOCATE 3, 1: PRINT SPACE$(80)

   COLOR 15

   LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
   QPrintRC "Relink Transaction Database", 5, 10, 15
   PrintHelp "System Operations.  Please wait."


   'PRINT
   
   DIM Trans AS GLTransRecType
   DIM Acct AS GLAcctRecType

   OpenTransFile GLTransFile, NumTrans&
   OpenAcctFile GLAcctFile, NumAccts

   LOCK GLTransFile
   LOCK GLAcctFile

   LogFile = FREEFILE
   LogFileName$ = "GLLINK.LOG"
   OPEN LogFileName$ FOR APPEND AS #LogFile

   PRINT #LogFile,
   PRINT #LogFile, "Accounting Database relink started @ " + DATE$ + " @ "; TIME$
   
   '-Set the pointers in the transaction file to zero
   FOR TCnt& = 1 TO NumTrans&
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Initializing transaction file. ###% complete."; Done#
      GET GLTransFile, TCnt&, Trans
      Trans.NextTran = 0
      PUT GLTransFile, TCnt&, Trans
   NEXT

   LOCATE 6, 10, 0
   PRINT STRING$(61, 32)

   '-Set the pointers in the account file to zero
   FOR Cnt = 1 TO NumAccts
      Done# = (Cnt / NumAccts) * 100
      LOCATE 6, 10, 0
      PRINT USING "Initializing account file. ###% complete."; Done#
      GET GLAcctFile, Cnt, Acct
      Acct.FrstTran = 0
      Acct.Bal = 0
      Acct.LastTran = 0
      PUT GLAcctFile, Cnt, Acct
   NEXT

   LOCATE 6, 10, 0
   PRINT STRING$(61, 32)

   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Relinking. ###% complete."; Done#

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
         LOCATE 13, 10
         PRINT USING "Orphaned transactions: #####"; BadTran

         'Trans.Marked = -1
         'PUT GLTransFile, TCnt&, Trans
         'Trans.Marked = 0

         '-Keep a list of orphaned transactions.
         GOSUB Logit

      END IF

   NEXT

   '-we're done here
   UNLOCK GLTransFile
   UNLOCK GLAcctFile

   IF BadTran > 0 THEN
      '-Errors in trans file
      PRINT #LogFile,
      PRINT #LogFile, "Orphan Transaction Totals:";
      PRINT #LogFile, TAB(58); FUsing(STR$(BadDebits#), "#######,.##");
      PRINT #LogFile, TAB(70); FUsing(STR$(BadCredits#), "#######,.##")
      PRINT #LogFile, "Relink completed @ " + DATE$ + " @ " + TIME$
      PRINT #LogFile, "Orphan transactions encountered! Call Customer Support."
   ELSE
      '-No errors in trans file
      PRINT #LogFile, "Relink of Accounting Databases successful. " + DATE$ + " @ " + TIME$
   END IF
   
   CLOSE

   '-Tell user we're done.
   IF BadTran > 0 THEN
      '-Errors in trans file
      QPrintRC "Press any key to view log.", 14, 10, 15
      K$ = INPUT$(1)
      PrintRptFile "Link Log", LogFileName$, LPTNo, RetCode%, 2
   ELSE
      '-No errors in trans file
      QPrintRC "Accounting database relink successful.", 6, 10, 15
      QPrintRC "Press any key to continue.", 7, 10, 15
      K$ = INPUT$(1)
   END IF
   
EXIT SUB

Logit:
   ToPrint$ = SPACE$(132)
   LSET ToPrint$ = Trans.AcctNum
   MID$(ToPrint$, 18) = Num2Date(Trans.TrDate)
   MID$(ToPrint$, 30) = LEFT$(Trans.Desc, 15)
   MID$(ToPrint$, 46) = Trans.Ref
   MID$(ToPrint$, 58) = FUsing(STR$(Trans.DrAmt), "#######,.##")
   MID$(ToPrint$, 70) = FUsing(STR$(Trans.CrAmt), "#######,.##")
   MID$(ToPrint$, 85) = "Record:" + STR$(TCnt&)
   PRINT #LogFile, ToPrint$
   BadDebits# = BadDebits# + Trans.DrAmt
   BadCredits# = BadCredits# + Trans.CrAmt
RETURN


END SUB

