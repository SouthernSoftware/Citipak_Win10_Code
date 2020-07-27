DEFINT A-Z

'--Procedures from Pro Lib
'$INCLUDE: 'pro7.bi'

'--Procedures in APAUX.BAS
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans%, RecLen%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE FUNCTION GetVendorRec (VendorCode$)

'--Procedures in COMNAUX.BAS
DECLARE FUNCTION Round# (DblNum#)

'--Procedures in This module
DECLARE SUB PrintLedger ()
DECLARE SUB PackLedgerFile ()
DECLARE SUB fixvendor ()
DECLARE SUB Working ()
DECLARE SUB Repost2GL ()
DECLARE SUB update81 ()
DECLARE SUB EditAPLTran ()

CONST False = 0, True = NOT False

'$INCLUDE: 'ap.bi'
'$INCLUDE: 'gl.bi'

DIM SHARED Vendor AS VendorRecType

STOP
'Repost2GL              'DANGER!!!!!!
'EditAPLTran            'DANGER!!!!!!
'PackLedgerFile         'Check out before running!!!!!!
'PrintLedger
'Working
'fixvendor
'update81               'RUN TO CONVERT FROM 8.0

'post2gl "tr2post.dat", badtrx%

SUB EditAPLTran

   'BEEP: BEEP: STOP
   CLS
   DIM apledger AS APLedger81RecType
   APLRecLen = LEN(apledger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
   FOR Cnt = 1 TO NumTrans
     GET APLedgerFile, Cnt, apledger

     apledger.PdCheckNum = 0
     apledger.PdCheckDate = 0
   
     PUT APLedgerFile, Cnt, apledger
   
   NEXT

   CLOSE

END SUB

SUB fixvendor

  CLS
  PRINT "Vendor Code Fix"
  PRINT "Press <Esc> to cancel, any other key to continue."
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB

  CLS

  SHARED Vendor AS VendorRecType

  OpenVendorFile VendorFile, NumVRecs

  FOR Cnt = 1 TO NumVRecs
    GET VendorFile, Cnt, Vendor
    LOCATE 1, 1: PRINT Cnt
    NVC$ = UCASE$(Vendor.VNum)
    Vendor.VNum = NVC$
    PUT VendorFile, Cnt, Vendor
  NEXT

  CLOSE
  CLS
  PRINT "Run Vendor Re-sort."

  END

END SUB

SUB PackLedgerFile

   STOP 'Do you know what you're doing?
   'Make sure gosub NewLedgerRec is set right
   
   'CLS
   'LOCATE 1, 1, 0
   'PRINT "Packing A/P Ledger Databases."
   'PRINT "Esc=Cancel. Any other key to continue."
   'K$ = UCASE$(INPUT$(1))
   'IF K$ = CHR$(27) THEN EXIT SUB
   
   DO
      CLS
      LOCATE 1, 1
      PRINT "Packing A/P Ledger Databases."
      LOCATE 3, 1
      INPUT "Purge Date: ", D$
      PurgeDate = Date2Num%(D$)
      IF PurgeDate > 0 THEN
         PRINT "Purge ALL HISTORY before "; Num2Date(PurgeDate); "?"
         K$ = UCASE$(INPUT$(1))
         IF K$ = "Y" THEN EXIT DO
      END IF
   LOOP
   
   '--Open the A/P Ledger file
   DIM apledger AS APLedger81RecType
   APLRecLen = LEN(apledger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
   '--Open the Ledger Distribution file
   DIM APDist AS APDistRecType
   APDRecLen = LEN(APDist)
   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
   
   '--Create a new ledger file to hold the new records
   DIM NAPLedger AS APLedger81RecType
   NAPLRecLen = LEN(apledger)
   NAPLedgerFile = FREEFILE
   OPEN "APLEDGER.NEW" FOR RANDOM AS NAPLedgerFile LEN = NAPLRecLen
   NumNewTrans& = LOF(NAPLedgerFile) \ NAPLRecLen

   '--Create a new ledger distribution file
   DIM NAPDist AS APDistRecType
   NAPDRecLen = LEN(NAPDist)
   NAPDistFile = FREEFILE
   OPEN "APDIST.NEW" FOR RANDOM AS NAPDistFile LEN = NAPDRecLen
   NumNewDistRecs& = LOF(NAPDistFile) \ NAPDRecLen

   FOR Cnt& = 1 TO NumTrans

      LOCATE 3, 1
      PRINT USING "Processing Transaction: #####"; Cnt&;
      PRINT USING " of ##### "; NumTrans
      
      GET APLedgerFile, Cnt&, apledger

      '--Move only transactions greater that the purge date
      '--or any open invoices to new file.
      IF apledger.TrDate > PurgeDate OR apledger.PayCode < 3 THEN
      
      'IF APLedger.TrCode = 1 THEN  '<--keep only invoices
      'IF Cnt& >= 1116 AND Cnt& <= 1145 THEN
      'IF Cnt& < 5520 OR Cnt& > 5521 THEN

         GOSUB NewLedgerRec
         NextDist& = apledger.FrstDist
         IF NextDist& > 0 THEN
           DO
             GET APDistFile, NextDist&, APDist
             NextDist& = APDist.NextDist
             GOSUB NewDistRec
           LOOP UNTIL NextDist& = 0
         END IF

      END IF

   NEXT

   CLOSE
   
   '--keep the old files
   NAME "APLEDGER.DAT" AS "APLEDGER.OLD"
   NAME "APDIST.DAT" AS "APDIST.OLD"

   '--rename the new files
   NAME "APLEDGER.NEW" AS "APLEDGER.DAT"
   NAME "APDIST.NEW" AS "APDIST.DAT"
   
   LOCATE 6, 1
   PRINT "Press any key to continue with relink."
   K$ = INPUT$(1)

   
EXIT SUB


NewLedgerRec:
   'STOP 'see comment below

   NumNewTrans& = NumNewTrans& + 1
   
   '--Version 8.1
   NAPLedger.VIN = apledger.VIN
   NAPLedger.VendorCode = apledger.VendorCode
   NAPLedger.VRecNum = apledger.VRecNum
   NAPLedger.TrDate = apledger.TrDate
   NAPLedger.GLDistDate = apledger.GLDistDate
   NAPLedger.DueDate = apledger.DueDate
   NAPLedger.TrCode = apledger.TrCode
   NAPLedger.DOCNum = apledger.DOCNum
   NAPLedger.PONUM = apledger.PONUM

   '--Fix me according to what you're doing!
   NAPLedger.PayCode = apledger.PayCode
   'NAPLedger.PayCode = 1 '<--Sets invoices back to open!

   '--check these also, set to zero if reopening payable
   NAPLedger.PrintCode = apledger.PrintCode
   NAPLedger.PdCheckNum = apledger.PdCheckNum
   NAPLedger.PdCheckDate = apledger.PdCheckDate

   NAPLedger.Comment = apledger.Comment
   NAPLedger.PSLFlag = apledger.PSLFlag
   NAPLedger.Get1099 = apledger.Get1099
   NAPLedger.Amt = apledger.Amt
   NAPLedger.FrstDist = 0 'apledger.FrstDist  'Relink these
   NAPLedger.LastDist = 0 'apledger.LastDist
   NAPLedger.NextTrans = 0 'apledger.NextTrans
   NAPLedger.Pad = apledger.Pad
   
   PUT NAPLedgerFile, NumNewTrans&, NAPLedger

   LOCATE 4, 1
   PRINT "New Transactions: "; NumNewTrans&

RETURN


NewDistRec:
   NumNewDistRecs& = NumNewDistRecs& + 1

   '--Set ledger record key to new ledger record
   NAPDist.APLedgerRec = NumNewTrans&

   '--These fields stays the same
   NAPDist.DistAcctRec = APDist.DistAcctRec
   NAPDist.DistAcctNum = APDist.DistAcctNum
   NAPDist.DistAmt = APDist.DistAmt
   NAPDist.NextDist = 0

   PUT NAPDistFile, NumNewDistRecs&, NAPDist
   LOCATE 5, 1
   PRINT "New Distributions: "; NumNewDistRecs&

RETURN


END SUB

SUB PrintLedger

  CLS

  DIM apledger AS APLedger81RecType
  APLRecLen = LEN(apledger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs

  RptFile = FREEFILE
  RptFileName$ = "apledger.prn"
  OPEN RptFileName$ FOR OUTPUT AS RptFile
  CommaFmt$ = "#######,.##"
  CommaFmtT$ = "###########,.##"
  RunTotal# = 0

  'PurgeDate = Date2Num%("08/22/96")

  FOR Cnt& = 1 TO NumTrans
    LOCATE 1, 1
    PRINT USING "Processing Ledger Record: #####"; Cnt&

    GET APLedgerFile, Cnt&, apledger
    GET VendorFile, apledger.VRecNum, Vendor

    'IF LEFT$(APLedger.VendorCode, 4) = "USMA" THEN
    'IF APLedger.PayCode < 3 THEN
    'IF APLedger.TrCode = 3 THEN
    'IF APLedger.TrDate <= PurgeDate THEN 'OR APLedger.PayCode < 3 THEN
    'IF APLedger.TrCode = 1 THEN
    'IF APLedger.TrDate >= Date2Num("07/01/96") AND APLedger.TrDate <= Date2Num("07/31/96") THEN
    'IF APLedger.TrDate = Date2Num("01/08/97") THEN

      Cnt = Cnt + 1
      RunTotal# = RunTotal# + apledger.Amt

      ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = STR$(Cnt&)
      MID$(ToPrint$, 2) = "Trans: " + STR$(Cnt&)
      MID$(ToPrint$, 15) = apledger.VendorCode
      MID$(ToPrint$, 27) = Vendor.VName
      'MID$(ToPrint$, 58) = "Trans: " + STR$(Cnt&)
      PRINT #RptFile, ToPrint$

      LSET ToPrint$ = ""
      
      SELECT CASE apledger.TrCode
         CASE 1
            MID$(ToPrint$, 5) = "Invoice " + apledger.DOCNum
         CASE 3
            MID$(ToPrint$, 5) = "Check " + apledger.DOCNum
         CASE 4
            MID$(ToPrint$, 5) = "Purchase Order  " + apledger.DOCNum
         CASE ELSE
            MID$(ToPrint$, 5) = "Unknown Tr Code" + STR$(apledger.TrCode) + apledger.DOCNum
            'APLedger.TrCode = 4
            'PUT APLedgerFile, Cnt&, APLedger
      END SELECT
      MID$(ToPrint$, 54) = "Total Amt: " + FUsing(STR$(apledger.Amt), CommaFmt$)
      PRINT #RptFile, ToPrint$

      LSET ToPrint$ = ""
      MID$(ToPrint$, 5) = "Tr Date: " + Num2Date$(apledger.TrDate)
      MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(apledger.DueDate)
      MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(apledger.GLDistDate)
      PRINT #RptFile, ToPrint$

      LSET ToPrint$ = ""
      SELECT CASE apledger.PayCode
         CASE 1
            Status$ = "Status: Open"
            PdChkDate$ = " "
            PdChkNum$ = " "
         CASE 3
            Status$ = "Status: Paid"
            PdChkDate$ = "Check Date: " + Num2Date(apledger.PdCheckDate)
            PdChkNum$ = "Check Num: " + STR$(apledger.PdCheckNum)
         CASE ELSE
            Status$ = "Status: Invalid Pay Code"
            PdChkDate$ = Num2Date(apledger.PdCheckDate)
            PdChkNum$ = STR$(apledger.PdCheckNum)
      END SELECT

      MID$(ToPrint$, 5) = Status$
      MID$(ToPrint$, 27) = PdChkDate$ 'Num2Date(APLedger.PdCheckDate)
      MID$(ToPrint$, 54) = PdChkNum$
      PRINT #RptFile, ToPrint$

      '--Now print the distribution
      NextDist& = apledger.FrstDist
      DistAmt# = 0

      IF NextDist& > 0 THEN  '--ignore checks, no distribution
         PRINT #RptFile, TAB(40); "Accounting Distribution:"
         DO
            GET APDistFile, NextDist&, APDist

            DistAmt# = DistAmt# + APDist.DistAmt
            TDistAmt# = TDistAmt# + APDist.DistAmt
            ThisRec& = NextDist&
            NextDist& = APDist.NextDist

            ToPrint$ = SPACE$(80)
            'MID$(ToPrint$, 30) = STR$(APDist.APLedgerRec)
            MID$(ToPrint$, 40) = APDist.DistAcctNum
            MID$(ToPrint$, 54) = FUsing(STR$(APDist.DistAmt), CommaFmt$)
            MID$(ToPrint$, 65) = STR$(APDist.APLedgerRec) + "/" + STR$(ThisRec&) + "/" + STR$(APDist.NextDist)
            PRINT #RptFile, ToPrint$

         LOOP UNTIL NextDist& = 0
         PRINT #RptFile, TAB(54); "------------"
         IF Round#(DistAmt#) <> Round#(apledger.Amt) THEN
            BalMsg$ = "* Link Error"
         ELSE
            BalMsg$ = ""
         END IF
         PRINT #RptFile, TAB(30); "Total Distributed:"; TAB(54); FUsing(STR$(DistAmt#), CommaFmtT$) + BalMsg$
      END IF
      PRINT #RptFile, STRING$(78, "=")

    'END IF
  NEXT

  PRINT #RptFile,
  PRINT #RptFile, "Running Total: " + FUsing(STR$(RunTotal#), CommaFmtT$)
  PRINT #RptFile,
  PRINT #RptFile, "Dist Amt:"; FUsing(STR$(TDistAmt#), CommaFmtT$)
  'PRINT #RptFile, "Tax Amt:"; Tax#
  PRINT #RptFile, Cnt
  CLOSE

  SHELL "list APLEDGER.PRN"

END SUB

'
SUB RecalcVendorBal

'NN2:
'GOSUB 10000
'GOSUB 11000
'FOR X = 1 TO LOF(1) / 384
'GET 1, X
'  GOSUB FIGUREBALANCE
'  LSET CURBAL$ = MKD$(BALANCE#)
'  BALANCE# = 0
'  PUT 1, X
'NEXT X
'CLOSE

'FIGUREBALANCE:
'TR = CVS(FIRST$)
'F1:
'IF TR = 0 THEN RETURN
'GET 2, TR
'IF CVI(PAYCODE$) = 3 THEN TR = CVS(NDATA$): GOTO F1
'BALANCE# = BALANCE# + CVD(amount$): TR = CVS(NDATA$): GOTO F1

END SUB

SUB Repost2GL
   
  DIM apledger AS APLedger81RecType
  APLRecLen = LEN(apledger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  DIM TR2Post AS GLTransRecType
  PostFile = FREEFILE
  OPEN "TR2POST.DAT" FOR RANDOM AS PostFile LEN = LEN(TR2Post)

  SHARED Vendor AS VendorRecType
  OpenVendorFile VendorFile, NumVRecs

  RptFile = FREEFILE
  RptFileName$ = "repost.prn"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  '--Init Variables
  CommaFmt$ = "######,.##"
  RunTotal# = 0
  'TDate = Date2Num("07/18/96")
  APAcct$ = "-010-01"
  SrcStamp$ = "CK013197"

  FOR Cnt& = 1 TO NumTrans
    LOCATE 1, 1
    PRINT USING "Processing Ledger Record: #####"; Cnt&

    GET APLedgerFile, Cnt&, apledger

      '--Posting for Invoices
    IF apledger.TrCode = 3 THEN

      RunTotal# = RunTotal# + apledger.Amt
      LOCATE 1, 1
      PRINT USING "Processing Ledger Record: #####"; Cnt&

      '--Get the Vendor Name for the Description field
      GET VendorFile, apledger.VRecNum, Vendor
      VendorName$ = Vendor.VName

      '--Get the distributions for this ledger rec
      NextDist& = apledger.FrstDist 'First Distribution record number
      DistAmt# = 0
      DO
        GET APDistFile, NextDist&, APDist

        '--Make the debit
        RecordNum = RecordNum + 1
        TR2Post.AcctNum = APDist.DistAcctNum
        TR2Post.TrDate = apledger.TrDate  'APLedger.GLDistDate
        LSET TR2Post.Desc = VendorName$
        TR2Post.Ref = apledger.DOCNum
        TR2Post.DrAmt = APDist.DistAmt
        TR2Post.CrAmt = 0
        TR2Post.Src = SrcStamp$
        TR2Post.Marked = False
        PUT PostFile, RecordNum, TR2Post
        
        '--Print the debit
        ToPrint$ = SPACE$(132)
        LSET ToPrint$ = TR2Post.AcctNum
        MID$(ToPrint$, 13) = Num2Date(TR2Post.TrDate)
        MID$(ToPrint$, 25) = LEFT$(TR2Post.Desc, 13)
        MID$(ToPrint$, 38) = TR2Post.Ref
        MID$(ToPrint$, 48) = FUsing$(STR$(TR2Post.DrAmt), "########,.##")
        MID$(ToPrint$, 61) = FUsing$(STR$(TR2Post.CrAmt), "########,.##")
        MID$(ToPrint$, 74) = TR2Post.Src
        PRINT #RptFile, ToPrint$
        Debits# = Round#(Debits#) + Round#(TR2Post.DrAmt)
        Credits# = Round#(Credits#) + Round#(TR2Post.CrAmt)
        

        '--Make the credit
        RecordNum = RecordNum + 1
        TR2Post.AcctNum = LEFT$(APDist.DistAcctNum, 2) + APAcct$
        TR2Post.TrDate = apledger.TrDate 'APLedger.GLDistDate
        LSET TR2Post.Desc = VendorName$
        TR2Post.Ref = apledger.DOCNum
        TR2Post.DrAmt = 0
        TR2Post.CrAmt = APDist.DistAmt
        TR2Post.Src = SrcStamp$
        TR2Post.Marked = False
        PUT PostFile, RecordNum, TR2Post

        '--Print the credit
        ToPrint$ = SPACE$(132)
        LSET ToPrint$ = TR2Post.AcctNum
        MID$(ToPrint$, 13) = Num2Date(TR2Post.TrDate)
        MID$(ToPrint$, 25) = LEFT$(TR2Post.Desc, 13)
        MID$(ToPrint$, 38) = TR2Post.Ref
        MID$(ToPrint$, 48) = FUsing$(STR$(TR2Post.DrAmt), "########,.##")
        MID$(ToPrint$, 61) = FUsing$(STR$(TR2Post.CrAmt), "########,.##")
        MID$(ToPrint$, 74) = TR2Post.Src
        PRINT #RptFile, ToPrint$
        Debits# = Round#(Debits#) + Round#(TR2Post.DrAmt)
        Credits# = Round#(Credits#) + Round#(TR2Post.CrAmt)

        '--Get the record number of the next distribution
        NextDist& = APDist.NextDist
       LOOP UNTIL NextDist& = 0
    END IF
  NEXT

  PRINT #RptFile, "Total Debits:   "; Debits#
  PRINT #RptFile, "Total Credits:  "; Credits#
  PRINT #RptFile, "APLedger Total: "; RunTotal#

  CLOSE

  SHELL "list REPOST.PRN"

'APLedger.TrCode
'APLedger.DOCNum
'APLedger.TrDate
'APLedger.DueDate
'APLedger.GLDistDate
'APLedger.Amt
'APLedger.PAYCODE

'APDist.APLedgerRec
'APDist.DistAcctNum
'APDist.DistAmt

END SUB

SUB update81

   DIM apledger AS APLedgerRecType
   APLRecLen = LEN(apledger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

   DIM APLedger81 AS APLedger81RecType
   APL81RecLen = LEN(APLedger81)
   NewLedgerFile = FREEFILE
   OPEN "APL81.DAT" FOR RANDOM AS NewLedgerFile LEN = APL81RecLen
   
   Recs! = LOF(APLedgerFile) MOD APLRecLen
   IF Recs! = 0 THEN 'STOP
     CLS
     LOCATE 1, 1
     PRINT "Converting A/P Ledger...."
   
     FOR T = 1 TO NumTrans
        GET APLedgerFile, T, apledger
        GET NewLedgerFile, T, APLedger81

        APLedger81.VIN = apledger.VIN
        APLedger81.VendorCode = apledger.VendorCode
        APLedger81.VRecNum = apledger.VRecNum
        APLedger81.TrDate = apledger.TrDate
        APLedger81.GLDistDate = apledger.GLDistDate
        APLedger81.DueDate = apledger.DueDate
        APLedger81.TrCode = apledger.TrCode
        APLedger81.DOCNum = apledger.DOCNum
        APLedger81.PONUM = apledger.PONUM
        APLedger81.PayCode = apledger.PayCode
        APLedger81.PrintCode = apledger.PrintCode
        APLedger81.PdCheckNum = apledger.PdCheckNum
        APLedger81.PdCheckDate = apledger.PdCheckDate
        APLedger81.Comment = ""
        APLedger81.PSLFlag = "Y"
        APLedger81.Get1099 = "Y"
        APLedger81.Amt = apledger.Amt
        APLedger81.FrstDist = apledger.FrstDist
        APLedger81.LastDist = apledger.LastDist
        APLedger81.NextTrans = apledger.NextTrans
        APLedger81.Pad = ""
      
        PUT NewLedgerFile, T, APLedger81

     NEXT

     CLOSE

     NAME "APLEDGER.DAT" AS "APLEDGER.80"
     NAME "APL81.DAT" AS "APLEDGER.DAT"

     LOCATE 1, 25
     PRINT "Done"
   ELSE
     CLS
     LOCATE 1, 1
     PRINT "A/P Ledger ok"
   END IF

END SUB

SUB Working

  '--Open the A/P Ledger file
  DIM apledger AS APLedgerRecType
  APLRecLen = LEN(apledger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  FOR Cnt& = 1 TO NumTrans

    LOCATE 3, 1
    PRINT USING "Processing Transaction: #####"; Cnt&;
    PRINT USING " of ##### "; NumTrans

    GET APLedgerFile, Cnt&, apledger

    'DOCNum$ = QPTrim$(APLedger.DOCNum)
    'SELECT CASE DOCNum$
    '  CASE "13627"
    '    APLedger.DOCNum = "13624"
    '  CASE "13628"
    '    APLedger.DOCNum = "13625"
    '  CASE "13629"
    '    APLedger.DOCNum = "13626"
    '  CASE ELSE
    'END SELECT

    'PdChkNum = APLedger.PDCheckNum
    'SELECT CASE PdChkNum
    '  CASE 13627
    '   APLedger.PDCheckNum = 13624
    '  CASE 13628
    '    APLedger.PDCheckNum = 13625
    '  CASE 13629
    '    APLedger.PDCheckNum = 13626
    '  CASE ELSE
    'END SELECT
    'PUT APLedgerFile, Cnt&, APLedger

  NEXT

  CLOSE



END SUB

