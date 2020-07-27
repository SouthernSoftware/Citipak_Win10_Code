DECLARE SUB ExpTrans ()
DEFINT A-Z


'--subs in other modules

'$INCLUDE: 'pro7.bi'
'$INCLUDE: 'gl.bi'
'$INCLUDE: 'glaux.bi'

CONST False = 0, True = NOT False

ExpTrans

SUB ExpTrans


   COLOR 15

   CLS

   PRINT "Export Of General Ledger Budget Transaction File"
   
   DIM Trans AS GLTransRecType
   BgtTransFile = FREEFILE
   BgtTransRecLen = LEN(Trans)
   OPEN "BgtTrans.DAT" FOR RANDOM AS BgtTransFile LEN = BgtTransRecLen
   OPEN "BUDGET.TXT" FOR OUTPUT AS #10
   NumTrans& = LOF(BgtTransFile) / BgtTransRecLen
   
   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Exporting. ###% complete."; Done#

      GET BgtTransFile, TCnt&, Trans
       DebitAmt$ = LTRIM$(STR$(Trans.DrAmt))
       CreditAmt$ = LTRIM$(STR$(Trans.CrAmt))
      PRINT #10, RTRIM$(Trans.AcctNum); "\";
      PRINT #10, RTRIM$(Num2Date(Trans.TrDate)); "\";
      PRINT #10, RTRIM$(Trans.Desc); "\";
      PRINT #10, RTRIM$(Trans.Ref); "\";
      PRINT #10, RTRIM$(DebitAmt$); "\";
      PRINT #10, RTRIM$(CreditAmt$); "\";
      PRINT #10, RTRIM$(Trans.Src)
      NEXT
      LPRINT "Budget Trans Mapping of File BUDGET.TXT"
      LPRINT "Comma Delimited"
      LPRINT
      LPRINT "Acct Number"
      LPRINT "Trans Date"
      LPRINT "Description"
      LPRINT "Reference"
      LPRINT "Debit Amt"
      LPRINT "Credit Amt"
      LPRINT "Source"
      LPRINT CHR$(12);

   '-we're done here
   

   CLOSE

EXIT SUB

Logit:
 
  RETURN


END SUB

