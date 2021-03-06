DECLARE SUB GetFundList (FundList$(), NumFunds%)
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
DECLARE SUB PackLedgerFile ()
DECLARE SUB RelinkLedger2Vendor (Orphans)
DECLARE SUB RelinkDist2Trans ()
DECLARE SUB PrintLedger ()

CONST False = 0, True = NOT False

'$INCLUDE: 'ap.bi'
'$INCLUDE: 'gl.bi'

DIM SHARED Vendor AS VendorRecType

'PackLedgerFile         'Check out before running!!!!!!
'PrintLedger
'END

  CLS
  PRINT "A/P Utility"
  PRINT "L = ReLink Transactions."
  PRINT "P = Print Ledger."
  PRINT "C = Purge Old History."

  K$ = UCASE$(INPUT$(1))

  SELECT CASE K$
    CASE "L"
      RelinkLedger2Vendor Orphans
      RelinkDist2Trans
      IF Orphans > 0 THEN
        PRINT
        PRINT "Ledger/Vendor Link encountered orphans!"
        PRINT "Number of Orphaned Transactions: "; Orphans
        PRINT "Contact Customer Support. 1-800-842-8190"
        PRINT "Press any key to exit."
        K$ = INPUT$(1)
      ELSE
        PRINT
        PRINT "Linking Operation complete.  Press any key to exit."
        K$ = INPUT$(1)
      END IF
    CASE "P"
        PrintLedger
    CASE "C"
        PackLedgerFile
    CASE ELSE
        PRINT "Link operation cancelled."
  END SELECT

RUN "glutil"

SUB PackLedgerFile

   
   'look at gosub NewLedgerRec:

   'CLS
   'LOCATE 1, 1, 0
   'PRINT "Packing A/P Ledger Databases."
   'PRINT "Esc=Cancel. Any other key to continue."
   'K$ = UCASE$(INPUT$(1))
   'IF K$ = CHR$(27) THEN EXIT SUB
   
   DO
      CLS
      LOCATE 1, 1
      PRINT "Close A/P History"
      LOCATE 3, 1
      INPUT "Enter the Closing Date: ", D$
      IF LEN(D$) = 0 THEN EXIT SUB
      PurgeDate = Date2Num%(D$)
      IF PurgeDate > 0 THEN
         PRINT "Purge ALL HISTORY through "; Num2Date(PurgeDate); "?"
         PRINT "Esc=Cancel"
         K$ = UCASE$(INPUT$(1))
         SELECT CASE K$
           CASE "Y"
              Cancel = False
              ExitFlag = True
           CASE CHR$(27)
              Cancel = True
              ExitFlag = True
         END SELECT
      END IF
   LOOP UNTIL ExitFlag

   IF Cancel THEN
     CLS
     EXIT SUB
   END IF

   
   CLS

   '--Open the A/P Ledger file
   DIM ApLedger AS APLedger81RecType
   APLRecLen = LEN(ApLedger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
   '--Open the Ledger Distribution file
   DIM APDist AS APDistRecType
   APDRecLen = LEN(APDist)
   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
   
   '--Create a new ledger file to hold the new records
   DIM NAPLedger AS APLedger81RecType
   NAPLRecLen = LEN(ApLedger)
   NAPLedgerFile = FREEFILE
   OPEN "APLEDGER.NEW" FOR RANDOM AS NAPLedgerFile LEN = NAPLRecLen
   NumNewTrans& = LOF(NAPLedgerFile) \ NAPLRecLen

   '--Create a new ledger distribution file
   DIM NAPDist AS APDistRecType
   NAPDRecLen = LEN(NAPDist)
   NAPDistFile = FREEFILE
   OPEN "APDIST.NEW" FOR RANDOM AS NAPDistFile LEN = NAPDRecLen
   NumNewDistRecs& = LOF(NAPDistFile) \ NAPDRecLen

   FOR cnt& = 1 TO NumTrans

      LOCATE 3, 1
      PRINT USING "Processing Transaction: #####"; cnt&;
      PRINT USING " of ##### "; NumTrans
      
      GET APLedgerFile, cnt&, ApLedger

      '--Move only transactions greater that the purge date
      '--or any open invoices to new file.
      IF ApLedger.TrDate > PurgeDate OR ApLedger.PayCode < 3 THEN
         'IF cnt& < 626 THEN
         GOSUB NewLedgerRec
         NextDist& = ApLedger.FrstDist
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

   RelinkDist2Trans
   RelinkLedger2Vendor (Orphans)
   
EXIT SUB


NewLedgerRec:
   'STOP 'see comment below

   NumNewTrans& = NumNewTrans& + 1

   '--Version 8
   'NAPLedger.VIN = apledger.VIN
   'NAPLedger.VendorCode = apledger.VendorCode
   'NAPLedger.VRecNum = apledger.VRecNum
   'NAPLedger.TrDate = apledger.TrDate
   'NAPLedger.GLDistDate = apledger.GLDistDate
   'NAPLedger.DUEDATE = apledger.DUEDATE
   'NAPLedger.TrCode = apledger.TrCode
   'NAPLedger.DOCNum = apledger.DOCNum
   'NAPLedger.PONUM = apledger.PONUM
   'NAPLedger.PayCode = apledger.PayCode
   'NAPLedger.PrintCode = apledger.PrintCode
   'NAPLedger.PDCheckNum = apledger.PDCheckNum
   'NAPLedger.PDCheckDate = apledger.PDCheckDate
   'NAPLedger.MiscCode = apledger.MiscCode
   'NAPLedger.Amt = apledger.Amt
   ''--These fields will have to be relinked
   'NAPLedger.FrstDist = 0
   'NAPLedger.LastDist = 0
   'NAPLedger.NextTrans = 0

   '--Version 8.1
   NAPLedger.VIN = ApLedger.VIN
   NAPLedger.VendorCode = ApLedger.VendorCode
   NAPLedger.VRecNum = ApLedger.VRecNum
   NAPLedger.TrDate = ApLedger.TrDate
   NAPLedger.GLDistDate = ApLedger.GLDistDate
   NAPLedger.DueDate = ApLedger.DueDate
   NAPLedger.TrCode = ApLedger.TrCode
   NAPLedger.DOCNum = ApLedger.DOCNum
   NAPLedger.PONUM = ApLedger.PONUM

   '--Fix me according to what you're doing!
   NAPLedger.PayCode = ApLedger.PayCode
   'NAPLedger.PayCode = 1 '--Sets invoices back to open!

   NAPLedger.PrintCode = ApLedger.PrintCode
   NAPLedger.PdCheckNum = ApLedger.PdCheckNum
   NAPLedger.PdCheckDate = ApLedger.PdCheckDate
   NAPLedger.Comment = ApLedger.Comment
   NAPLedger.PSLFlag = ApLedger.PSLFlag
   NAPLedger.Get1099 = ApLedger.Get1099
   NAPLedger.Amt = ApLedger.Amt
   NAPLedger.FrstDist = 0 'apledger.FrstDist  'Relink these
   NAPLedger.LastDist = 0 'apledger.LastDist
   NAPLedger.NextTrans = 0 'apledger.NextTrans
   NAPLedger.Pad = ApLedger.Pad
   
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

   '--Relinking on the fly. DO NOT USE..UNTESTED
   'SELECT CASE NextDist&
   '   CASE 0
   '      '--No more distributions
   '      NAPDist.NextDist = 0
   '   CASE ELSE
   '      '--if There is another distribution
   '      '--It will be the next record number
   '      NAPDist.NextDist = NumNewDistRecs + 1
   'END SELECT

   PUT NAPDistFile, NumNewDistRecs&, NAPDist
   LOCATE 5, 1
   PRINT "New Distributions: "; NumNewDistRecs&

RETURN


END SUB

SUB PrintLedger

  CLS

  DIM ApLedger AS APLedger81RecType
  APLRecLen = LEN(ApLedger)
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

  FOR cnt& = 1 TO NumTrans
    LOCATE 1, 1
    PRINT USING "Processing Ledger Record: #####"; cnt&

    GET APLedgerFile, cnt&, ApLedger
    IF ApLedger.VRecNum > 0 THEN
      GET VendorFile, ApLedger.VRecNum, Vendor
      VendorName$ = Vendor.VName
    ELSE
      VendorName$ = "Orphaned Transaction"
    END IF
      
    cnt = cnt + 1
    RunTotal# = RunTotal# + ApLedger.Amt

    ToPrint$ = SPACE$(80)
    MID$(ToPrint$, 2) = "Trans: " + STR$(cnt&)
    MID$(ToPrint$, 15) = ApLedger.VendorCode
    MID$(ToPrint$, 27) = VendorName$
    PRINT #RptFile, ToPrint$

    LSET ToPrint$ = ""

    SELECT CASE ApLedger.TrCode
       CASE 1
          'MID$(ToPrint$, 5) = "Invoice " + APLedger.DOCNum
          GOSUB PrintInv
       CASE 3
          'MID$(ToPrint$, 5) = "Check " + APLedger.DOCNum
          GOSUB PrintChk
       CASE 4
          'MID$(ToPrint$, 5) = "Purchase Order  " + APLedger.DOCNum
          GOSUB PrintPO
       CASE ELSE
          GOSUB PrintOther
          'APLedger.TrCode = 4
          'PUT APLedgerFile, Cnt&, APLedger
    END SELECT

    '--Now print the distribution
    NextDist& = ApLedger.FrstDist
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
          MID$(ToPrint$, 40) = APDist.DistAcctNum
          MID$(ToPrint$, 54) = FUsing(STR$(APDist.DistAmt), CommaFmt$)
          MID$(ToPrint$, 65) = STR$(APDist.APLedgerRec) + "/" + STR$(ThisRec&) + "/" + STR$(APDist.NextDist)
          PRINT #RptFile, ToPrint$

       LOOP UNTIL NextDist& = 0
       PRINT #RptFile, TAB(54); "------------"
       IF Round#(DistAmt#) <> Round#(ApLedger.Amt) THEN
          BalMsg$ = "* Link Error"
       ELSE
          BalMsg$ = ""
       END IF
       PRINT #RptFile, TAB(30); "Total Distributed:"; TAB(50); FUsing(STR$(DistAmt#), CommaFmtT$) + BalMsg$
    END IF
    PRINT #RptFile, STRING$(78, "=")

  NEXT

  PRINT #RptFile,
  PRINT #RptFile, "Running Total: " + FUsing(STR$(RunTotal#), CommaFmtT$)
  PRINT #RptFile,
  PRINT #RptFile, "Dist Amt:"; FUsing(STR$(TDistAmt#), CommaFmtT$)
  'PRINT #RptFile, "Tax Amt:"; Tax#
  PRINT #RptFile, cnt
  CLOSE

  SHELL "list APLEDGER.PRN"

EXIT SUB




PrintInv:
  MID$(ToPrint$, 5) = "Invoice " + ApLedger.DOCNum
  MID$(ToPrint$, 54) = "Total Amt: " + FUsing(STR$(ApLedger.Amt), CommaFmt$)
  PRINT #RptFile, ToPrint$

  LSET ToPrint$ = ""
  MID$(ToPrint$, 5) = "Tr Date: " + Num2Date$(ApLedger.TrDate)
  MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(ApLedger.DueDate)
  MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(ApLedger.GLDistDate)
  PRINT #RptFile, ToPrint$

  LSET ToPrint$ = ""
  SELECT CASE ApLedger.PayCode
    CASE 1
       Status$ = "Status: Open"
       PdChkDate$ = " "
       PdChkNum$ = " "
    CASE 3
       Status$ = "Status: Paid"
       PdChkDate$ = "Check Date: " + Num2Date(ApLedger.PdCheckDate)
       PdChkNum$ = "Check Num: " + STR$(ApLedger.PdCheckNum)
    CASE ELSE
       Status$ = "Status: Invalid Pay Code"
       PdChkDate$ = Num2Date(ApLedger.PdCheckDate)
       PdChkNum$ = STR$(ApLedger.PdCheckNum)
  END SELECT
  MID$(ToPrint$, 5) = Status$
  MID$(ToPrint$, 27) = PdChkDate$ 'Num2Date(APLedger.PdCheckDate)
  MID$(ToPrint$, 54) = PdChkNum$
  PRINT #RptFile, ToPrint$

RETURN


PrintChk:
  LSET ToPrint$ = ""
  MID$(ToPrint$, 5) = "Check " + ApLedger.DOCNum
  MID$(ToPrint$, 29) = "Dated: " + Num2Date$(ApLedger.TrDate)
  MID$(ToPrint$, 54) = "Check Amt: " + FUsing(STR$(ApLedger.Amt), CommaFmt$)
  PRINT #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'PRINT #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'SELECT CASE APLedger.PayCode
  '  CASE 1
  '     Status$ = "Status: Open"
  '     PdChkDate$ = " "
  '     PdChkNum$ = " "
  '  CASE 3
  '     Status$ = "Status: Paid"
  '     PdChkDate$ = "Check Date: " + Num2Date(APLedger.PdCheckDate)
  '     PdChkNum$ = "Check Num: " + STR$(APLedger.PdCheckNum)
  '  CASE ELSE
  '     Status$ = "Status: Invalid Pay Code"
  '     PdChkDate$ = Num2Date(APLedger.PdCheckDate)
  '     PdChkNum$ = STR$(APLedger.PdCheckNum)
  'END SELECT
  'MID$(ToPrint$, 5) = Status$
  'MID$(ToPrint$, 27) = PdChkDate$ 'Num2Date(APLedger.PdCheckDate)
  'MID$(ToPrint$, 54) = PdChkNum$
  'PRINT #RptFile, ToPrint$
  
RETURN


PrintPO:
  LSET ToPrint$ = ""
  MID$(ToPrint$, 5) = "Purchase Order  " + ApLedger.DOCNum
  MID$(ToPrint$, 29) = "PO Date: " + Num2Date$(ApLedger.TrDate)
  MID$(ToPrint$, 54) = "Total Amt: " + FUsing(STR$(ApLedger.Amt), CommaFmt$)
  PRINT #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'MID$(ToPrint$, 5) = "Tr Date: " + Num2Date$(APLedger.TrDate)
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'PRINT #RptFile, ToPrint$
  '
  'LSET ToPrint$ = ""
  SELECT CASE ApLedger.PayCode
    CASE 4
       Status$ = "Status: Open"
       PdChkDate$ = " "
       PdChkNum$ = " "
    CASE -4
       Status$ = "Status: Paid"
       PdChkDate$ = "Check Date: " + Num2Date(ApLedger.PdCheckDate)
       PdChkNum$ = "Check Num: " + STR$(ApLedger.PdCheckNum)
    CASE ELSE
       Status$ = "Status: Invalid Pay Code"
       PdChkDate$ = Num2Date(ApLedger.PdCheckDate)
       PdChkNum$ = STR$(ApLedger.PdCheckNum)
  END SELECT
  MID$(ToPrint$, 5) = Status$
  MID$(ToPrint$, 27) = PdChkDate$ 'Num2Date(APLedger.PdCheckDate)
  MID$(ToPrint$, 54) = PdChkNum$
  PRINT #RptFile, ToPrint$

RETURN


PrintOther:
 LSET ToPrint$ = ""
 MID$(ToPrint$, 5) = "Unknown Tr Code" + STR$(ApLedger.TrCode) + ApLedger.DOCNum
 PRINT #RptFile, ToPrint$
RETURN

  
END SUB

SUB RebuildHistory

  CLS
  PRINT "Take a break.. this could take a while. DO NOT DISTURB!"

  DIM ApLedger AS APLedger81RecType
  APLRecLen = LEN(ApLedger)
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

  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  
  FOR cnt& = 1 TO NumTrans
    LOCATE 1, 1
    cnt = cnt + 1
    PRINT USING "Processing Ledger Record: #####"; cnt&

    GET APLedgerFile, cnt&, ApLedger


    '--Now print the distribution
    NextDist& = ApLedger.FrstDist
    DistAmt# = 0
    IF NextDist& > 0 THEN  '--ignore checks, no distribution
      DO
        GET APDistFile, NextDist&, APDist
        DistAmt# = DistAmt# + APDist.DistAmt
        TDistAmt# = TDistAmt# + APDist.DistAmt
        ThisRec& = NextDist&
        NextDist& = APDist.NextDist
      LOOP UNTIL NextDist& = 0
      IF Round#(DistAmt#) <> Round#(ApLedger.Amt) AND ApLedger.TrCode = 3 THEN
        GOSUB RebuildLink
      END IF

    ELSE
      '--No distribution attached
      IF ApLedger.TrCode = 3 THEN
        GOSUB RebuildLink
      END IF

    END IF
  NEXT

  CLOSE

EXIT SUB

RebuildLink:
  CheckNumber& = VAL(ApLedger.DOCNum)
  FOR L = 1 TO NumTrans
    GET APLedgerFile, L, ApLedger
    IF ApLedger.TrCode = 1 THEN
      IF APLeger.PdChkNum = CheckNumber THEN
        STOP
        NextDist& = ApLedger.FrstDist
        DistAmt# = 0
        IF NextDist& > 0 THEN  '--ignore checks, no distribution
          DO
            GET APDistFile, NextDist&, APDist
            DistAmt# = DistAmt# + APDist.DistAmt
            TDistAmt# = TDistAmt# + APDist.DistAmt
            ThisRec& = NextDist&
            NextDist& = APDist.NextDist
          LOOP UNTIL NextDist& = 0
        ELSE
          STOP
        END IF

      
      END IF
    END IF
  NEXT
  
  RETURN

END SUB

SUB RelinkDist2Trans

   'COLOR 7
   'CLS
   LOCATE 7, 1, 0
   'PRINT "Re-Linking Ledger Distribution Database."
   PRINT "Linking Distributions to Ledger."

   DIM ApLedger AS APLedger81RecType
   APLRecLen = LEN(ApLedger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
   DIM APLdgrDist AS APDistRecType
   APLDRecLen = LEN(APLdgrDist)
   OpenAPDistFile APDistFile, NumDistRecs&, APLDRecLen

   'PRINT "Initializing Ledger Database."
   FOR cnt& = 1 TO NumTrans
      GET APLedgerFile, cnt&, ApLedger
      ApLedger.FrstDist = 0
      ApLedger.LastDist = 0
      PUT APLedgerFile, cnt&, ApLedger
   NEXT

   FOR cnt& = 1 TO NumDistRecs&
      
      Done! = (cnt& / NumDistRecs&) * 100
      LOCATE 8, 1
      PRINT USING "###% Complete."; Done!

      '--Assume no one else will follow.
      GET APDistFile, cnt&, APLdgrDist
      APLdgrDist.NextDist = 0
      PUT APDistFile, cnt&, APLdgrDist

      '--Get the parent record
      

      GET APLedgerFile, APLdgrDist.APLedgerRec, ApLedger

      IF ApLedger.FrstDist > 0 THEN
         '--We're not the first one here, so let us not forget those who have
         '--come before us
         Prev& = ApLedger.LastDist

         '--This is now the new last distribution
         '--Update Last Dist pointer in apledger to this rec
         ApLedger.LastDist = cnt&
         PUT APLedgerFile, APLdgrDist.APLedgerRec, ApLedger

         '--Get the former last distribution
         '--and tell it that this rec is the next one
         GET APDistFile, Prev&, APLdgrDist
         APLdgrDist.NextDist = cnt&
         PUT APDistFile, Prev&, APLdgrDist

      ELSE
         '--Virgin territory. we're now first and last
         ApLedger.FrstDist = cnt&
         ApLedger.LastDist = cnt&
         PUT APLedgerFile, APLdgrDist.APLedgerRec, ApLedger
      END IF

   NEXT

   CLOSE

   'PRINT "Press any key to continue."
   'K$ = INPUT$(1)
   
END SUB

SUB RelinkLedger2Vendor (Orphans)

   'COLOR 7

   CLS
   LOCATE 1, 1, 0
   PRINT "Linking A/P Databases."
   
   LOCATE 3, 1, 0
   PRINT "Validating Ledger/Vendor Link."

   'DIM Vendor AS VendorRecType
   OpenVendorFile VendorFile, NumVRecs

   DIM ApLedger AS APLedger81RecType
   APLRecLen = LEN(ApLedger)
   OpenAPLedgerFile APLedgerFile, NumTrans, APLRecLen
   
   LogFile = FREEFILE
   LogFileName$ = "APLINK.LOG"
   OPEN LogFileName$ FOR OUTPUT AS LogFile

   ToPrint$ = SPACE$(80)
   ToPrint$ = "Linking operations began on " + DATE$ + " " + TIME$
   PRINT #LogFile, ToPrint$

   'CommaFmt$ = "#######,.##"

   '--Reset the vendor trans pointers to 0.
   'PRINT "Initializing Vendor File."
   FOR cnt& = 1 TO NumVRecs
      GET VendorFile, cnt&, Vendor
         Vendor.FrstTran = 0
         Vendor.LastTran = 0
      PUT VendorFile, cnt&, Vendor
   NEXT
   
   PRINT "APLedger Records:"; NumTrans

   '--Relink Transactions to Vendor
   FOR cnt& = 1 TO NumTrans
      GET APLedgerFile, cnt&, ApLedger

      '--reset next transaction pointer to 0
      ApLedger.NextTrans = 0
      PUT APLedgerFile, cnt&, ApLedger

      VRecNum = ApLedger.VRecNum   'GetVendorRec(apledger.VendorCode)
      'VRECNUM = GetVendorRec(APLedger.VendorCode)
      
      IF VRecNum > 0 AND VRecNum <= NumVRecs THEN
         
         GET VendorFile, VRecNum, Vendor

         IF Vendor.FrstTran > 0 THEN
            '--Vendor has previous transactions..
            '--Remember the last transaction for this vendor
            Prev& = Vendor.LastTran

            '--In the vendor record...
            '--Set the Last Trans pointer to this record
            Vendor.LastTran = cnt&
            PUT VendorFile, VRecNum, Vendor

            '--In the apledger record...
            '--Set the Last trans pointer in the prev trans
            '--to point to this record
            GET APLedgerFile, Prev&, ApLedger
            ApLedger.NextTrans = cnt&
            PUT APLedgerFile, Prev&, ApLedger
         ELSE
            '--First Trans for this vendor
            '--set both pointers to this ledger record
            Vendor.FrstTran = cnt&
            Vendor.LastTran = cnt&
            PUT VendorFile, VRecNum, Vendor
        END IF
      ELSE
        GOSUB LogAPLOrphan
      END IF

      LOCATE 5, 1
      PRINT "Processed Record: "; cnt&
      
   NEXT

   ToPrint$ = SPACE$(80)
   ToPrint$ = "Linking operations completed on " + DATE$ + " " + TIME$
   PRINT #LogFile, ToPrint$
   IF Ophan > 1 THEN
     ToPrint$ = SPACE$(80)
     ToPrint$ = "Orphaned Transactions: " + STR$(Ophan)
   ELSE
     ToPrint$ = SPACE$(80)
     ToPrint$ = "No Orphaned Transactions. "
   END IF
   PRINT #LogFile, ToPrint$
   CLOSE

EXIT SUB
   
LogAPLOrphan:
  Ophan = Oprhan + 1
  PRINT #LogFile, "Rec: " + STR$(cnt&) + " Orphan: " + STR$(ApLedger.VRecNum) + " " + ApLedger.VendorCode
RETURN


END SUB

