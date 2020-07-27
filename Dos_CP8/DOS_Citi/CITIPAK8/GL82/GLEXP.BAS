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

   PRINT "Export Of General Ledger Transaction File"
   
   DIM Trans AS GLTransRecType
   

   TransRecLen = LEN(Trans)
   TransFileNum = FREEFILE
   OPEN "GLTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TransFileNum LEN = TransRecLen
   NumTrans& = LOF(TransFileNum) \ TransRecLen

   OPEN "GLTRANS.TXT" FOR OUTPUT AS #10

   
   


   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE 6, 10, 0
      PRINT USING "Exporting. ###% complete."; Done#

      GET TransFileNum, TCnt&, Trans
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
      LPRINT "G/L Trans Mapping of File GLTRANS.TXT"
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

