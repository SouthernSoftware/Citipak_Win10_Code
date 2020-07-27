DECLARE SUB NOpenAPLedgerFile (NAPLedgerFile%, NNumTrans%, NRecLen%)
DECLARE SUB NOpenAPDistFile (NAPDistFile%, NNumDistRecs&, NRecLen%)
DECLARE SUB MergeFiles ()
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
DECLARE SUB Post2GL (FileName$, BadTrans%)

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
'MergeFiles
'Repost2GL              'DANGER!!!!!!
'EditAPLTran            'DANGER!!!!!!
'PackLedgerFile         'Check out before running!!!!!!
'PrintLedger
'Working
'fixvendor
update81
'Post2GL "TR2POST.DAT", BadTrans%

END

SUB EditAPLTran

  BEEP: BEEP: STOP
  CLS
  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
  
  'FOR Cnt = 1 TO NumTrans
  '  GET APLedgerFile, Cnt, APLedger
  '  IF APLedger.TrCode = 1 THEN
  '    'STOP
  '    APLedger.PayCode = 0
  '    APLedger.PdCheckNum = 0
  '    APLedger.PdCheckDate = 0
  '    PUT APLedgerFile, Cnt, APLedger
  '  END IF
  'NEXT
  

  GET APLedgerFile, 3365, APLedger
  'PRINT Vendor.VName
  APLedger.PayCode = -1
  PUT APLedgerFile, 3365, APLedger
  
  GET APLedgerFile, 3522, APLedger
  APLedger.PayCode = -1
  PUT APLedgerFile, 3522, APLedger
  
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

SUB MergeFiles

   STOP 'look at gosub AddLedgerRec: and record filter before running!

   CLS
   LOCATE 1, 1, 0
   PRINT "Merging A/P Ledger Databases."
   PRINT "Esc=Cancel. Any other key to continue."
   K$ = UCASE$(INPUT$(1))
   IF K$ = CHR$(27) THEN EXIT SUB

   '--Open the A/P Ledger file
   DIM APLedger AS APLedger81RecType
   APLRecLen = LEN(APLedger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

   '--Open the Ledger Distribution file
   DIM APDist AS APDistRecType
   APDRecLen = LEN(APDist)
   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
   

   '--Open the New A/P Ledger file
   DIM APLedgerA AS APLedger81RecType
   NAPLRecLen = LEN(APLedgerA)
   NOpenAPLedgerFile NAPLedgerFile, NNumTrans, NAPLRecLen

   '--Open the New Distribution file
   DIM APDistA AS APDistRecType
   NAPDRecLen = LEN(APDistA)
   NOpenAPDistFile NAPDistFile, NNumDistRecs&, NAPDRecLen
   
   FOR Cnt = 1 TO NNumTrans
     LOCATE 3, 1
     PRINT USING "Merging Transaction: #####"; Cnt&;
     PRINT USING " of ##### "; NumTrans

     GET NAPLedgerFile, Cnt, APLedgerA
     GOSUB AddLedgerRec

     NextDist& = APLedgerA.FrstDist
     'IF NextDist& > 0 THEN
       DO
         GET NAPDistFile, NextDist&, APDistA
         NextDist& = APDistA.NextDist
         GOSUB AddDistRec
       LOOP UNTIL NextDist& = 0
     'END IF
   NEXT

   CLOSE
   
   LOCATE 6, 1
   PRINT "Press any key to continue. Remember to relink."
   K$ = INPUT$(1)


EXIT SUB


AddLedgerRec:
   NumTrans = NumTrans + 1
   
   '--Version 8.1
   APLedger.VIN = APLedgerA.VIN
   APLedger.VendorCode = APLedgerA.VendorCode
   APLedger.VRecNum = APLedgerA.VRecNum
   APLedger.TrDate = APLedgerA.TrDate
   APLedger.GLDistDate = APLedgerA.GLDistDate
   APLedger.DueDate = APLedgerA.DueDate
   APLedger.TrCode = APLedgerA.TrCode
   APLedger.DOCNum = APLedgerA.DOCNum
   APLedger.PONUM = APLedgerA.PONUM
   APLedger.PayCode = APLedgerA.PayCode
   APLedger.PrintCode = APLedgerA.PrintCode
   APLedger.PdCheckNum = APLedgerA.PdCheckNum
   APLedger.PdCheckDate = APLedgerA.PdCheckDate
   APLedger.Comment = APLedgerA.Comment
   APLedger.PSLFlag = APLedgerA.PSLFlag
   APLedger.Get1099 = APLedgerA.Get1099
   APLedger.Amt = APLedgerA.Amt

   APLedger.FrstDist = 0 'apledgerA.FrstDist  'Relink these
   APLedger.LastDist = 0 'apledgerA.LastDist
   APLedger.NextTrans = 0 'apledgerA.NextTrans
   APLedger.Pad = APLedgerA.Pad

   PUT APLedgerFile, NumTrans, APLedger

   LOCATE 4, 1
   PRINT "New Transactions: "; NumNewTrans&

RETURN


AddDistRec:

   NumDistRecs& = NumDistRecs& + 1

   '--Set ledger record key to new ledger record
   APDist.APLedgerRec = NumTrans

   '--These fields stays the same
   APDist.DistAcctRec = APDistA.DistAcctRec
   APDist.DistAcctNum = APDistA.DistAcctNum
   APDist.DistAmt = APDistA.DistAmt
   APDist.NextDist = 0
   
   PUT APDistFile, NumDistRecs&, APDist
   LOCATE 5, 1
   PRINT "New Distributions: "; NumDistRecs&

RETURN


END SUB

SUB NOpenAPDistFile (NAPDistFile, NNumDistRecs&, NRecLen)

  NAPDistFile = FREEFILE
  OPEN "BADATA\APDIST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS NAPDistFile LEN = NRecLen
  NNumDistRecs& = LOF(NAPDistFile) \ NRecLen


END SUB

SUB NOpenAPLedgerFile (NAPLedgerFile, NNumTrans, NRecLen)

  NAPLedgerFile = FREEFILE
  OPEN "BADATA\APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS NAPLedgerFile LEN = NRecLen
  NNumTrans = LOF(NAPLedgerFile) \ NRecLen

END SUB

SUB PackLedgerFile
   
   STOP 'look at gosub NewLedgerRec: and record filter before running!

   CLS
   LOCATE 1, 1, 0
   PRINT "Packing A/P Ledger Databases."
   PRINT "Esc=Cancel. Any other key to continue."
   K$ = UCASE$(INPUT$(1))
   IF K$ = CHR$(27) THEN EXIT SUB
   
   'DO
   '   CLS
   '   LOCATE 1, 1
   '   PRINT "Packing A/P Ledger Databases."
   '   LOCATE 3, 1
   '   INPUT "Purge Date: ", D$
   '   PurgeDate = Date2Num%(D$)
   '   IF PurgeDate > 0 THEN
   '      PRINT "Purge ALL HISTORY before "; Num2Date(PurgeDate); "?"
   '      K$ = UCASE$(INPUT$(1))
   '      IF K$ = "Y" THEN EXIT DO
   '   END IF
   'LOOP
   
   '--Open the A/P Ledger file
   DIM APLedger AS APLedger81RecType
   APLRecLen = LEN(APLedger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
   '--Open the Ledger Distribution file
   DIM APDist AS APDistRecType
   APDRecLen = LEN(APDist)
   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
   
   '--Create a new ledger file to hold the new records
   DIM NAPLedger AS APLedger81RecType
   NAPLRecLen = LEN(APLedger)
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
      
      GET APLedgerFile, Cnt&, APLedger

      '--filter
      'IF Cnt& = 19 THEN STOP
      '--Move only transactions greater that the purge date
      '--or any open invoices to new file.
      'IF apledger.TrDate > PurgeDate OR apledger.PayCode < 3 THEN
      '--remove rec  for Troy
      'IF Cnt& < 19 OR Cnt& > 19 THEN
      '--Move only invoices
      'IF APLedger.TrCode = 1 THEN
      '--Remove Lunenburg 6x posting
      'IF Cnt& < 770 THEN
      '--Remove Richland dupe posting 11/21/96
      'IF Cnt& < 1145 OR Cnt& > 1149 THEN
      '--Remove VASS dupe posting 12/17/96

      '--Move only invoices
      'IF APLedger.TrCode = 1 THEN

      IF Cnt& < 4445 THEN
         GOSUB NewLedgerRec
         NextDist& = APLedger.FrstDist
         IF NextDist& > 0 THEN
            DO
               GET APDistFile, NextDist&, APDist
               NextDist& = APDist.NextDist
               GOSUB NewDistRec
            LOOP UNTIL NextDist& = 0
         END IF
      ELSE
         PRINT "Not updating: "; Cnt&
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
   NumNewTrans& = NumNewTrans& + 1
   
   '--Version 8.1
   NAPLedger.VIN = APLedger.VIN
   NAPLedger.VendorCode = APLedger.VendorCode
   NAPLedger.VRecNum = APLedger.VRecNum
   NAPLedger.TrDate = APLedger.TrDate
   NAPLedger.GLDistDate = APLedger.GLDistDate
   NAPLedger.DueDate = APLedger.DueDate
   NAPLedger.TrCode = APLedger.TrCode
   NAPLedger.DOCNum = APLedger.DOCNum
   NAPLedger.PONUM = APLedger.PONUM

   '*****Fix me according to what you're doing!
   '--Sets new record to original value
   NAPLedger.PayCode = APLedger.PayCode
   '--Sets invoices back to open!
   'NAPLedger.PayCode = 1
   '****set pd check fields
   
   NAPLedger.PrintCode = APLedger.PrintCode

   '****Check these also   0 for invoices
   NAPLedger.PdCheckNum = APLedger.PdCheckNum
   NAPLedger.PdCheckDate = APLedger.PdCheckDate
   '****

   NAPLedger.Comment = APLedger.Comment
   NAPLedger.PSLFlag = APLedger.PSLFlag
   NAPLedger.Get1099 = APLedger.Get1099
   NAPLedger.Amt = APLedger.Amt

   NAPLedger.FrstDist = 0 'apledger.FrstDist  'Relink these
   NAPLedger.LastDist = 0 'apledger.LastDist
   NAPLedger.NextTrans = 0 'apledger.NextTrans
   NAPLedger.Pad = APLedger.Pad
   
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

  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs

  ToPrint$ = SPACE$(80)
  RptFile = FREEFILE
  RptFileName$ = "apledger.prn"
  OPEN RptFileName$ FOR OUTPUT AS RptFile
  CommaFmt$ = "#######,.##"
  CommaFmtT$ = "###########,.##"
  RunTotal# = 0
  
  'PurgeDate = Date2Num%("08/22/96")
  BegDate = Date2Num%("02/01/97")
  EndDate = Date2Num%("02/28/97")

  FOR Cnt& = 1 TO NumTrans
    LOCATE 1, 1
    PRINT USING "Processing Ledger Record: #####"; Cnt&

    GET APLedgerFile, Cnt&, APLedger
    GET VendorFile, APLedger.VRecNum, Vendor

    'IF LEFT$(APLedger.VendorCode, 4) = "USMA" THEN
    'IF APLedger.PayCode < 3 THEN
    'IF APLedger.TrCode = 3 THEN
    'IF APLedger.TrDate <= PurgeDate THEN 'OR APLedger.PayCode < 3 THEN
    'IF APLedger.TrDate >= Date2Num("07/01/96") AND APLedger.TrDate <= Date2Num("07/31/96") THEN
    '--finds open payables
    'IF APLedger.TrCode = 3 THEN 'AND APLedger.PayCode = 1 THEN

    'IF APLedger.TrDate >= BegDate AND APLedger.TrDate <= EndDate THEN
      
      IF Cnt& = 3365 OR Cnt& = 3522 THEN
        STOP
        GOSUB FixMe
      END IF


      'Cnt = Cnt + 1
      RunTotal# = RunTotal# + APLedger.Amt
      TrAmt$ = FUsing(STR$(APLedger.Amt), CommaFmt$)
      
      SELECT CASE APLedger.TrCode
         CASE 1
            Tr$ = "INV"
         CASE 3
            Tr$ = "CHK"
         CASE -3
            Tr$ = "VCK"
         CASE 4
            Tr$ = "PO"
         CASE -4
            Tr$ = "Closed PO"
         CASE ELSE
            Tr$ = "???"
      END SELECT
      
      '--Line 1
      LSET ToPrint$ = ""
      MID$(ToPrint$, 1) = "Rec:" + STR$(Cnt&)
      MID$(ToPrint$, 12) = Vendor.VName
      MID$(ToPrint$, 52) = APLedger.VendorCode
      MID$(ToPrint$, 65) = Tr$ + TrAmt$
      'MID$(ToPrint$, 70) = APLedger.VendorCode
      PRINT #RptFile, ToPrint$

      SELECT CASE APLedger.TrCode
         CASE 1 'Invoice
            SELECT CASE APLedger.PayCode
            CASE 1
               Status$ = "Open Item"
            CASE 3
               Status$ = "Paid Item"
            END SELECT
            GOSUB PrintInvoice
         CASE -1 'Invoice
            Status$ = "Void Inv"
            GOSUB PrintInvoice
         CASE 3
            GOSUB PrintCheck
         CASE -3
            GOSUB PrintCheck
         CASE 4
            GOSUB PrintPO
         CASE -4
            'GOSUB PrintClosedPO
         CASE ELSE
      END SELECT

      '--Now print the distribution
      NextDist& = APLedger.FrstDist
      DistAmt# = 0

      IF NextDist& > 0 THEN  '--ignore checks, no distribution
         PRINT #RptFile, TAB(40); "Accounting Distribution:"
         DO
            GET APDistFile, NextDist&, APDist

            'IF QPTrim$(APDist.DistAcctNum) = "10-133-0000" THEN
            ' SCnt = SCnt + 1
            ' Tax# = Tax# + APDist.DistAmt
            'END IF

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
         IF Round#(DistAmt#) <> Round#(APLedger.Amt) THEN
            BalMsg$ = "* Link Error"
         ELSE
            BalMsg$ = ""
         END IF
         PRINT #RptFile, TAB(30); "Total Distributed:"; TAB(54); FUsing(STR$(DistAmt#), CommaFmtT$) + BalMsg$
      'END IF
      PRINT #RptFile, STRING$(78, "=")

      'APLedger.PayCode = 1
      'PUT APLedgerFile, Cnt&, APLedger
     ELSE
        PRINT #RptFile, TAB(30); "*INVALID LINK IN DISTRIBUTION*"
        PRINT #RptFile,
     END IF
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

EXIT SUB
  'APLedger.DOCNum

PrintPO:
  RETURN

PrintInvoice:
  LSET ToPrint$ = ""
  MID$(ToPrint$, 5) = "Tr Date: " + Num2Date$(APLedger.TrDate)
  MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  PRINT #RptFile, ToPrint$
  
  LSET ToPrint$ = ""
  MID$(ToPrint$, 5) = Status$
  IF Status$ = "Paid" THEN
    MID$(ToPrint$, 29) = "Paid Date: " + Num2Date$(APLedger.PdCheckDate)
    MID$(ToPrint$, 55) = "STR$(APLedger.PdCheckNum)"
  END IF
  PRINT #RptFile, ToPrint$
RETURN


PrintCheck:
  LSET ToPrint$ = ""
  MID$(ToPrint$, 5) = "Chk Date: " + Num2Date$(APLedger.TrDate)
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'MID$(ToPrint$, 70) = "Status"
  PRINT #RptFile, ToPrint$
  RETURN


FixMe:
  BEEP: STOP
  
  
  APLedger.TrCode = -1
  APLedger.PdCheckDate = 0
  APLedger.PdCheckNum = 0
  PUT APLedgerFile, Cnt&, APLedger
  
  'APLedger.PayCode
  'APLedger.PdCheckDate
  'APLedger.PdCheckNum
  'PUT APLedgerFile, Cnt&, APLedger
  'PUT VendorFile, APLedger.VRecNum, Vendor
  
  RETURN
  
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
   
  STOP
  'whoa pardner. check this out before running. variables are set as needed!!

  APAcct$ = "-201-0000"
  SrcStamp$ = "AP011597"
  
  CLS

  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)
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

  FOR Cnt& = 1 TO NumTrans
    LOCATE 1, 1
    PRINT USING "Processing Ledger Record: #####"; Cnt&

    GET APLedgerFile, Cnt&, APLedger

    '--Posting for Invoices Trcode=1
    '--Posting for PO's Trcode=1
    IF APLedger.TrCode = 1 THEN

      RunTotal# = RunTotal# + APLedger.Amt
      LOCATE 1, 1
      PRINT USING "Processing Ledger Record: #####"; Cnt&

      '--Get the Vendor Name for the Description field
      GET VendorFile, APLedger.VRecNum, Vendor
      VendorName$ = Vendor.VName

      '--Get the distributions for this ledger rec
      NextDist& = APLedger.FrstDist 'First Distribution record number
      DistAmt# = 0
      DO
        GET APDistFile, NextDist&, APDist

        '--Make the debit
        RecordNum = RecordNum + 1
        TR2Post.AcctNum = APDist.DistAcctNum
        'TR2Post.TrDate = APLedger.GLDistDate
        TR2Post.TrDate = APLedger.TrDate
        LSET TR2Post.Desc = VendorName$
        TR2Post.Ref = APLedger.DOCNum
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
        'TR2Post.TrDate = APLedger.GLDistDate
        TR2Post.TrDate = APLedger.TrDate
        LSET TR2Post.Desc = VendorName$
        TR2Post.Ref = APLedger.DOCNum
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

'--Ledger fields
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

   DIM APLedger AS APLedgerRecType
   APLRecLen = LEN(APLedger)
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
        GET APLedgerFile, T, APLedger
        GET NewLedgerFile, T, APLedger81

        APLedger81.VIN = APLedger.VIN
        APLedger81.VendorCode = APLedger.VendorCode
        APLedger81.VRecNum = APLedger.VRecNum
        APLedger81.TrDate = APLedger.TrDate
        APLedger81.GLDistDate = APLedger.GLDistDate
        APLedger81.DueDate = APLedger.DueDate
        APLedger81.TrCode = APLedger.TrCode
        APLedger81.DOCNum = APLedger.DOCNum
        APLedger81.PONUM = APLedger.PONUM
        APLedger81.PayCode = APLedger.PayCode
        APLedger81.PrintCode = APLedger.PrintCode
        APLedger81.PdCheckNum = APLedger.PdCheckNum
        APLedger81.PdCheckDate = APLedger.PdCheckDate
        APLedger81.Comment = ""
        APLedger81.PSLFlag = "Y"
        APLedger81.Get1099 = "Y"
        APLedger81.Amt = APLedger.Amt
        APLedger81.FrstDist = APLedger.FrstDist
        APLedger81.LastDist = APLedger.LastDist
        APLedger81.NextTrans = APLedger.NextTrans
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
  DIM APLedger AS APLedgerRecType
  APLRecLen = LEN(APLedger)
  OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen

  FOR Cnt& = 1 TO NumTrans

    LOCATE 3, 1
    PRINT USING "Processing Transaction: #####"; Cnt&;
    PRINT USING " of ##### "; NumTrans

    GET APLedgerFile, Cnt&, APLedger

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

