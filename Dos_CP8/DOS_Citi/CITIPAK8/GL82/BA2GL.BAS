DEFINT A-Z

DECLARE SUB ConvertFundFile ()
DECLARE SUB ConvertAcctFile (Vers, TotLen, FundLen, AcctLen, DetLen)
DECLARE SUB ConvertTransfile (Vers, FundLen%, AcctLen%, DetLen%)
DECLARE SUB ConvertBudgetFile (Vers, TotLen, FundLen, AcctLen, DetLen)
DECLARE FUNCTION PadAcct$ (AcctNum$)
DECLARE SUB Relink ()
DECLARE SUB InitBgtVars ()

'$INCLUDE: 'GL.BI'
'$INCLUDE: 'GLAUX.BI'
'$INCLUDE: 'pro7.BI'

'--Change the following for the specific version we're converting

'--7/9 digit Version
'$INCLUDE: 'BARECS.BI'
'TotLen = 9        'total acct length
'FundLen = 2       'length of fund code
'AcctLen = 3       '  "    "  dept   "
'DetLen = 4        '  "    "  object "
'V$ = "Seven/Nine digit account numbers."
'Vers = 7
'Vers = 9

'--7 digit Version
''$INCLUDE: 'BARECS7.BI'
TotLen = 7        'total acct length
FundLen = 2       'length of fund code
AcctLen = 3       '  "    "  dept   "
DetLen = 2        '  "    "  object "
V$ = "Seven digit account numbers."
Ver = 0

''$INCLUDE: 'BAvass.BI'
'TotLen = 9        'total acct length
'FundLen = 2       'length of fund code
'AcctLen = 4       '  "    "  dept   "
'DetLen = 3        '  "    "  object "
'V$ = "Vass custom conversion."
'Vers = 0

'--7/9 digit Version
''$INCLUDE: 'BARECS.BI'
'TotLen = 12        'total acct length
'FundLen = 2       'length of fund code
'AcctLen = 4       '  "    "  dept   "
'DetLen = 4        '  "    "  object "
'V$ = "Brookneal, VA"
'Vers = 0

DIM SHARED Colr AS INTEGER
Colr = 7

CLS
LOCATE 1, 1, 0
QPrintRC "Convert Citipak version 6.x and 7.x to version 8.", 10, 1, Colr
QPrintRC V$, 11, 1, Colr
QPrintRC "Press [Esc] to cancel or any other key to continue.", 12, 1, Colr
K$ = INPUT$(1)
IF K$ = CHR$(27) THEN
   CLS
   END
END IF

CLS

ConvertFundFile
ConvertAcctFile Vers, TotLen, FundLen, AcctLen, DetLen
ConvertTransfile Vers, FundLen, AcctLen, DetLen
'ConvertBudgetFile Vers, TotLen, FundLen, AcctLen, DetLen
Relink

QPrintRC "G/L Conversion complete. Press any key to exit.", 12, 1, Colr
K$ = INPUT$(1)
CLS
END


'--to redo bgt conversion
'KILL "BGTTRANS.DAT"
'InitBgtVars
'ConvertBudgetFile Vers, TotLen, FundLen, AcctLen, DetLen

SUB ConvertAcctFile (Vers, TotLen, FundLen, AcctLen, DetLen)

  SHARED Colr AS INTEGER
   
  DIM BANdx  AS BAAcctNdxType
  DIM BAAcct AS BAAcctRecType
  DIM BABgtPrep AS BABgtPrepRecType
  DIM GLAcct AS GLAcctRecType

  BABgtPrepRecLen = LEN(BABgtPrep)
  BAAcctRecLen = LEN(BAAcct)
  BANdxRecLen = LEN(BANdx)
  GLAcctRecLen = LEN(GLAcct)

  OPEN "baacctdx.dat" FOR RANDOM AS #1 LEN = BANdxRecLen
  OPEN "baacct.dat" FOR RANDOM AS #2 LEN = BAAcctRecLen
  OPEN "glacct.dat" FOR RANDOM AS #3 LEN = GLAcctRecLen
  OPEN "babudprp.dat" FOR RANDOM AS #4 LEN = BABgtPrepRecLen
  NumAccts = LOF(1) \ BANdxRecLen

  QPrintRC "Converting account file.    % Complete.", 3, 1, Colr
  Done$ = SPACE$(4)
  FOR Cnt = 1 TO NumAccts

    RSET Done$ = STR$(INT((Cnt / NumAccts) * 100))
    QPrintRC Done$, 3, 25, Colr

    GET #1, Cnt, BANdx
    GET #2, BANdx.RecNum, BAAcct
    GET #4, BANdx.RecNum, BABgtPrep
    GET #3, Cnt, GLAcct
    
    AcctNum$ = StripAcct(BAAcct.Num)       'remove dashes

    IF Vers = 9 THEN
      AcctNum$ = PadAcct(AcctNum$)
    END IF
    
    AcctNum$ = FmtAcct$(AcctNum$, FundLen, AcctLen, DetLen)

    GLAcct.Deleted = 0
    GLAcct.Num = AcctNum$
    GLAcct.Title = BAAcct.Title
    GLAcct.Typ = BAAcct.Typ
    GLAcct.FrstTran = BAAcct.FrstTran
    GLAcct.LastTran = BAAcct.LastTran
    GLAcct.PYAct = BABgtPrep.PYAct
    GLAcct.BegBal = BAAcct.BegBal
    GLAcct.Bgt = BAAcct.Bgt       '0 '-if history is bad then get budget amt from acct file
    GLAcct.Bal = BAAcct.Bal
    GLAcct.Encumb = 0 'BAAcct.Encumb '-Rem out Enc field for 7 digit version
    GLAcct.MTD = BAAcct.MTD
    GLAcct.YTD = BAAcct.YTD
    GLAcct.NYEst = BABgtPrep.EstAmt
    GLAcct.NYReq = BABgtPrep.ReqAmt
    GLAcct.NYRec = BABgtPrep.RecAmt
    GLAcct.NYApp = BABgtPrep.AppAmt
    GLAcct.FrstBTran = 0            'Pointer to First Budget Trans
    GLAcct.LastBTran = 0            'Pointer to Last Budget Trans

    PUT #3, Cnt, GLAcct

  NEXT

  CLOSE
  QPrintRC "Updating Account Index...", 4, 1, Colr
  SortAcctIndex
  QPrintRC " Done.", 4, 24, Colr

END SUB

SUB ConvertBudgetFile (Vers, TotLen, FundLen, AcctLen, DetLen)

  SHARED Colr AS INTEGER
  DIM Acct AS GLAcctRecType

  DIM BABudget AS BABgtTransType
  BABgtRecLen = LEN(BABudget)
  BABgtFile = FREEFILE
  OPEN "babudget.dat" FOR RANDOM AS BABgtFile LEN = BABgtRecLen
  NumBARecs = LOF(1) \ BABgtRecLen
  
  OpenAcctFile AcctFile, NumAccts

  DIM BgtTrans AS GLTransRecType
  GLBgtFile = FREEFILE
  GLBgtRecLen = LEN(BgtTrans)
  OPEN "bgttrans.dat" FOR RANDOM AS GLBgtFile LEN = GLBgtRecLen
  'NumGLRecs = LOF(GLBgtFile) \ GLBgtRecLen

  Done$ = SPACE$(4)
  QPrintRC "Converting budget file.    % Complete.", 6, 1, Colr
  
  FOR Cnt = 1 TO NumBARecs              'Assign edit file to trans format

    RSET Done$ = STR$(INT((Cnt / NumBARecs) * 100))
    QPrintRC Done$, 6, 24, Colr

    GET BABgtFile, Cnt, BABudget

    IF LEN(QPTrim$(BABudget.Acct)) THEN
      
      AcctNum$ = StripAcct$(BABudget.Acct)

      IF Vers = 9 THEN
        AcctNum$ = PadAcct(AcctNum$)
      END IF
     
      AcctNum$ = FmtAcct$(AcctNum$, FundLen, AcctLen, DetLen)

      AcctRec = FindAcct(AcctNum$)     'BABudget.Acct = AcctNum$
      IF AcctRec > 0 THEN
        GET AcctFile, AcctRec, Acct

        IF BABudget.BgtAmt <> 0 THEN
          SELECT CASE Acct.Typ
            CASE "E"
              IF BABudget.BgtAmt > 0 THEN
                Acct.Bgt = Round#(Acct.Bgt + BABudget.BgtAmt)
                DebitAmt# = Round#(BABudget.BgtAmt)
                CreditAmt# = 0
              ELSE
                Acct.Bgt = Round#(Acct.Bgt + BABudget.BgtAmt)
                DebitAmt# = 0
                CreditAmt# = ABS(Round#(BABudget.BgtAmt))
              END IF
            CASE "R"
              IF BABudget.BgtAmt > 0 THEN
                Acct.Bgt = Round#(Acct.Bgt + BABudget.BgtAmt)
                CreditAmt# = Round#(BABudget.BgtAmt)
                DebitAmt# = 0
              ELSE
                Acct.Bgt = Round#(Acct.Bgt + BABudget.BgtAmt)
                CreditAmt# = 0
                DebitAmt# = ABS(Round#(BABudget.BgtAmt))
              END IF
            CASE ELSE
              DebitAmt# = 0
              CreditAmt# = 0
          END SELECT

          PUT AcctFile, AcctRec, Acct
          
          BgtTrans.AcctRec = AcctRec
          BgtTrans.AcctNum = AcctNum$
          BgtTrans.TrDate = Date2Num(BABudget.TrDate)
          SELECT CASE BABudget.TrType
            CASE "A"
              BgtTrans.Desc = "Amendment"
            CASE "B"
              BgtTrans.Desc = "Bgt Appr'n"
          END SELECT
          BgtTrans.Ref = ""
          BgtTrans.DrAmt = DebitAmt#
          BgtTrans.CrAmt = CreditAmt#
          BgtTrans.Src = "BC" + ConvDateStr$(DATE$)
          BgtTrans.NextTran = 0

          NumBgtTrans = NumBgtTrans + 1
          PUT GLBgtFile, NumBgtTrans, BgtTrans

          '--------------------------------Start linking here
          IF Acct.FrstBTran = 0 THEN       'if first trans for this acct,
             Acct.FrstBTran = NumBgtTrans  'assign first & last pointers to
             Acct.LastBTran = NumBgtTrans  'this transaction
             PUT AcctFile, AcctRec, Acct

          ELSE                             'otherwise
                                           'in the account file..
             Prev = Acct.LastBTran         'remember the prev trans pointer,
             Acct.LastBTran = NumBgtTrans  'reset last trans to this trans
             PUT AcctFile, AcctRec, Acct

                                           'In the trans file...
             GET GLBgtFile, Prev, BgtTrans  'Get the last transaction
             BgtTrans.NextTran = NumBgtTrans     'reset pointer to this trans
             PUT GLBgtFile, Prev, BgtTrans

          END IF
        END IF
      ELSE
        BadCnt = BadCnt + 1
      END IF
    END IF
  NEXT
  'STOP
  CLOSE

END SUB

SUB ConvertFundFile

  SHARED Colr AS INTEGER

  DIM BAFund AS BAFundType
  DIM GLFund AS GLFundRecType
  
  OPEN "bafund.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 64
  NumBAFunds = LOF(1) \ 64

  OpenFundFile FundFileNum%, NumFunds%
  
  CLS
  QPrintRC "Converting fund file.    % Complete.", 1, 1, Colr
  Done$ = SPACE$(4)
  FOR Cnt = 1 TO NumBAFunds

    RSET Done$ = STR$(INT((Cnt / NumBAFunds) * 100))
    QPrintRC Done$, 1, 22, Colr

    GET #1, Cnt, BAFund
    GET FundFileNum, Cnt, GLFund

    GLFund.Deleted = 0
    GLFund.FundNum = QPTrim$(STR$(BAFund.FundCode))
    GLFund.Title = BAFund.FundName

    PUT FundFileNum, Cnt, GLFund

  NEXT

  CLOSE
  QPrintRC "Updating Fund Index...", 2, 1, Colr
  SortFundIndex
  QPrintRC " Done.", 2, 21, Colr

END SUB

SUB ConvertTransfile (Vers, FundLen%, AcctLen%, DetLen%)

  SHARED Colr AS INTEGER

  DIM BATrans AS BATransRecType
  DIM GLTrans AS GLTransRecType
  BARecLen = LEN(BATrans)
  GLRecLen = LEN(GLTrans)

  OPEN "BATRANS.DAT" FOR RANDOM AS #1 LEN = BARecLen
  NumBARecs& = LOF(1) \ BARecLen

  OPEN "GLTRANS.DAT" FOR RANDOM AS #2 LEN = GLRecLen
  NumGLRecs& = LOF(2) \ GLRecLen

  QPrintRC "Converting transaction file.    % complete.", 5, 1, Colr
  Done$ = SPACE$(4)

  FOR Cnt& = 1 TO NumBARecs&
    RSET Done$ = (STR$(INT((Cnt& / NumBARecs&) * 100)))
    QPrintRC Done$, 5, 29, Colr

    GET #1, Cnt&, BATrans
    GET #2, Cnt&, GLTrans
                     
    GLTrans.TrDate = Date2Num(BATrans.TrDate)

    AcctNum$ = StripAcct$(BATrans.AcctNum)

    '--7/9 version ##-###-##(-##) ----------------------------------------
    IF Vers = 9 THEN
      AcctNum$ = PadAcct(AcctNum$)
    END IF
    
    GLTrans.AcctNum = FmtAcct(AcctNum$, FundLen%, AcctLen%, DetLen%)
    
    GLTrans.Desc = BATrans.Desc
    GLTrans.Ref = BATrans.Ref
    GLTrans.DrAmt = Round#(BATrans.DrAmt)
    GLTrans.CrAmt = Round#(BATrans.CrAmt)
    GLTrans.Src = BATrans.Source
    GLTrans.NextTran = BATrans.NextTran
    PUT #2, Cnt&, GLTrans

  NEXT

  CLOSE

END SUB

SUB InitBgtVars
  
  DIM GLAcct AS GLAcctRecType
  GLAcctRecLen = LEN(GLAcct)
  OPEN "glacct.dat" FOR RANDOM AS #1 LEN = GLAcctRecLen
  NumAccts = LOF(1) \ GLAcctRecLen

  QPrintRC "Initializing.    % Complete.", 3, 1, Colr
  Done$ = SPACE$(4)
  FOR Cnt = 1 TO NumAccts

    RSET Done$ = STR$(INT((Cnt / NumAccts) * 100))
    QPrintRC Done$, 3, 14, Colr

    GET #1, Cnt, GLAcct

    GLAcct.Bgt = 0
    GLAcct.FrstBTran = 0
    GLAcct.LastBTran = 0

    PUT #1, Cnt, GLAcct

  NEXT

  CLOSE

END SUB

FUNCTION PadAcct$ (AcctNum$)

  AcctNum$ = QPTrim$(AcctNum$)           'trim it up
  AcctLength = LEN(AcctNum$)             'find the length of the acct
  IF AcctLength = 7 THEN                 'make 'em the same length
    AcctNum$ = AcctNum$ + "00"
  END IF

  PadAcct = AcctNum$

END FUNCTION

SUB Relink '(Row, Col)

   SHARED Colr AS INTEGER
   COLOR Colr
   
   Row = 7: Col = 1
   LOCATE Row, Col: PRINT SPACE$(80)

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

   '-Set the pointers in the transaction file to zero
   FOR TCnt& = 1 TO NumTrans&
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE Row, Col, 0
      PRINT USING "Initializing transaction file. ###% complete."; Done#
      GET GLTransFile, TCnt&, Trans
      Trans.NextTran = 0
      PUT GLTransFile, TCnt&, Trans
   NEXT

   LOCATE Row, Col, 0
   PRINT STRING$(80, 32)

   '-Set the pointers in the account file to zero
   FOR Cnt = 1 TO NumAccts
      Done# = (Cnt / NumAccts) * 100
      LOCATE Row, Col, 0
      PRINT USING "Initializing account file. ###% complete."; Done#
      GET GLAcctFile, Cnt, Acct
      Acct.FrstTran = 0
      Acct.LastTran = 0
      PUT GLAcctFile, Cnt, Acct
   NEXT

   LOCATE Row, Col, 0
   PRINT STRING$(80, 32)

   '-Start the relink process
   FOR TCnt& = 1& TO NumTrans&

      '-Something to look at while this is going on
      Done# = (TCnt& / NumTrans&) * 100
      LOCATE Row, Col, 0
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
         LOCATE Row + 2, Col
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
      QPrintRC "There were errors in transaction file!", Row + 1, Col, Colr
      QPrintRC "Orphaned transactions listed in LINK.LOG file.", Row + 2, Col, Colr
      QPrintRC "Press any key to continue.", Row + 3, Col, Colr
      K$ = INPUT$(1)
   ELSE
      '-No errors in trans file
      PRINT #LogFile, "Relink of Accounting Databases successful. " + DATE$ + " @ " + TIME$
      QPrintRC "Accounting database relink successful.", Row, Col, Colr
      'QPrintRC "Press any key to continue.", Row + 1, Col, Colr
      'K$ = INPUT$(1)
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

