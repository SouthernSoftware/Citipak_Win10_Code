'-for demo only screen
         CONST DemoOnly = "DEMOONLY"
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'    CONST GLAcctDataFile = "BAACCT.DAT"
     CONST GLAcctIdxFile = "BAACCTDX.DAT"
    CONST JGLAcctIdxFile = "GLACCT.IDX"

         CONST CalcQLib = "PRQL\PRCALC"
         CONST CtrlQLib = "PRQL\PRCTRL"
     CONST EmployeeQLib = "PRQL\PREMPLOY"
'      CONST ELookUpQLib = "PRQL\PREMPLUP"
         CONST MiscQLib = "PRQL\PRMISCRN"
'          CONST TaxQLib = "PRQL\PRSFTAX"
        CONST AccruQLib = "PRQL\PRACCRU"
        CONST EMPrnQLib = "PRQL\PREMPRNT"
'        CONST CheckQLib = "PRQL\PRCHECK"
           CONST W2QLib = "PRQL\PRW2"

      'CONST MiscLibName = "PRQL\PRMISCRN.QSL"

      CONST PRActiveFile = "PRDATA\PRACTIVE.FLG"    '*
          'CONST PRDDPath = "PRDD\"     '*

      CONST EmpData1Name = "PRDATA\PREMP1.DAT"    '*
      CONST EmpData2Name = "PRDATA\PREMP2.DAT"    '*
      CONST EmpData3Name = "PRDATA\PREMP3.DAT"    '*
       CONST EmpIdxLName = "PRDATA\PREMPL.IDX"    '*
       CONST EmpIdxNName = "PRDATA\PREMPN.IDX"    '*
    CONST EMPNumFileName = "PRDATA\PREMPUNO.DAT"  '*
    CONST EMPPinFileName = "PRDATA\PREMPPIN.DAT"  '*
    CONST AccrueFileName = "PRDATA\PRACCRUE.DAT"  '*

    CONST ChecksFileName = "PRDATA\PRCHECKS.DAT"  '*
   CONST VoucherFileName = "PRDATA\PRVOUCHR.DAT"  '*
  CONST CHistoryFileName = "PRDATA\PRCKHIST.DAT"

 CONST PPDefaultFileName = "PRDATA\PRPPDEF.DAT"   '*
 CONST PPHistoryFileName = "PRDATA\PRPPHIST.DAT"

    CONST ReqFldFileName = "PRDATA\PRRFLD.DAT"

 CONST PayrollStatusFile = "PRDATA\PRSTATUS.DAT"

    CONST RetireFileName = "PRDATA\PRRETIRE.DAT"  '*
   CONST DedCodeFileName = "PRDATA\PRDEDCOD.DAT"  '*
   CONST MatCodeFileName = "PRDATA\PRMATCOD.DAT"  '*


'new: 07-14-95
   CONST PrinterDefFile = "PRDATA\PRPRNDEF.DAT"  '*
'new: 07-17-95
   CONST PrinterSetUPFile = "PRDATA\PRPRNSET.DAT"  '*
'new: 05-21-95
   CONST StateTaxDefFile = "PRDATA\PRSTADEF.DAT"  '*

'new: 11-24-97
      CONST DraftInfoError = "DDINFERR"
 CONST PPDraftInfoFileName = "PRDATA\PPDFINFO.DAT"  '*
   CONST DraftInfoFileName = "PRDATA\PRDRAFTI.DAT"    '*
      CONST Draft1FileName = "PRDATA\PRDRAFT1.DAT"
      CONST Draft5FileName = "PRDATA\PRDRAFT5.DAT"
      CONST Draft6FileName = "PRDATA\PRDRAFT6.DAT"
      CONST Draft8FileName = "PRDATA\PRDRAFT8.DAT"
      CONST Draft9FileName = "PRDATA\PRDRAFT9.DAT"


  CONST StateTaxFileName = "PRDATA\PRSTATAX.DAT"  '*
CONST FederalTaxFileName = "PRDATA\PRFEDTAX.DAT"  '*
   CONST ErnCodeFileName = "PRDATA\PRERNCOD.DAT"  '*
     CONST LeaveFileName = "PRDATA\PRLEAVE.DAT"   '*
       CONST EICFileName = "PRDATA\PREICTBL.DAT"  '*
      CONST UnitFileName = "PRDATA\PRUNIT.DAT"    '*
       CONST SysFileName = "PRDATA\PRSYS.DAT"     '*
 CONST TransWorkFileName = "PRDATA\PRTRANST.DAT"  '*
 CONST TransHistFileName = "PRDATA\PRTRANSH.DAT"  '*
    CONST CheckPrintFile = "PRDATA\CHKPRNT.DAT"

   CONST W2SetupFile = "PRDATA\PRW2SETU.DAT"
    CONST W2InfoFile = "PRDATA\PRW2INFO.DAT"
    CONST W2PrintFile = "PRDATA\W2PRNT.DAT"

'  CONST PRReportExe = "PRREPORT"

  CONST ReportListFileName = "PRDATA\PRRPTLST.DAT"

  CONST ManRegisterRptName = "PRRPTS\MANREGIS.RPT"
  CONST PayRegisterRptName = "PRRPTS\REGISTER.RPT"
 CONST DistributionRptName = "PRRPTS\DISTRIBU.RPT"
CONST CheckRegisterRptName = "PRRPTS\CHECKREG.RPT"
       CONST EmpHistoryRpt = "PRRPTS\EMPHist.RPT"
      CONST AccrualRptFile = "PRRPTS\ACCRUAL.RPT"
'--------------------------------

     CONST EmpDataFileMask = "PRRPTS\PREMPRPT.DPM"
      CONST CheckAlignMask = "PRRPTS\PRCHKMSK.DPM"
        CONST W2PrintMask1 = "PRRPTS\PRW2MSK1.DPM"
        CONST W2PrintMask2 = "PRRPTS\PRW2MSK2.DPM"

'--------------------------------
      CONST GLErrorScrn = "GLERROR"
      CONST NegativePay = "NEGNET"
     CONST UpdatingDisk = "UPDATDSK"
        CONST UpdatedOK = "UPDATEOK"
       CONST DelFromPay = "DELFPAY"
   CONST EMPBlankLookUp = "BLANKUP"
      CONST EMPNotFound = "NOTFOUND"
       CONST EMPDupeNum = "DUPENUMB"
        CONST EMPNoRecs = "NORECORD"
          CONST NoBlank = "NOBLANK"
  CONST CreatingPayRoll = "CREATPAY"
      CONST MountChecks = "MOUNTCHK"
CONST PrintingAlignment = "PRNTALGN"
   CONST PrintTestAgain = "PRNTAGIN"
      CONST ReadyToPost = "RDY2POST"
          CONST Posting = "POSTING"
         CONST Printing = "PRINTING"
   CONST PrintingChecks = "NOWPRNT"
   CONST ProcessingChecks = "PROCHEKS"
   CONST PayrollInProg = "INPROGRS"

 CONST ReadyPrintChecks = "CHKPRNT"
   CONST UpdatingChecks = "UPCHECK"
  CONST InvalidCheckNum = "INVALNUM"

       CONST CheckReady = "CHKREADY"
        CONST AskCancel = "ASKCANCL"
         CONST NotReady = "PRNERR"
         CONST Canceled = "CANCELED"
        CONST PrintDone = "PRNTDONE"

       CONST EmpHistory = "EMPHIST"
   CONST ReportParmError = "RPTPRMER"
  CONST NoPeriodDefaults = "NODEFALT"

      CONST NoManualNow = "NOMANUAL"
      CONST NoNormalNow = "NONORMAL"

     CONST DistributErr = "DISTERR"
       CONST AccountErr = "ACCTERR"

   CONST ProcessReport = "PRORPT"

   CONST FileViewWindow = "VIEWIND2"
   CONST NoTransHist = "NOTRANSH"

   CONST ZeroEmpYTD = "ZEROYTD"

   CONST W2SetUp = "W2SETUP"
   CONST W2Info = "W2FORM"

   CONST LockError = "LOCKERR"

