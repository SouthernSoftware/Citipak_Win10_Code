DEFINT A-Z
DECLARE SUB CheckSetupFile ()
DECLARE SUB CheckRateFile ()
DECLARE SUB CheckDraftFile ()
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'ubdraft.BI'
'$INCLUDE: 'ubrate.BI'
'$INCLUDE: 'ubsetup.BI'

  STACK 8000

  CLS
  PRINT
  PRINT "UB File Checking/Conversion Utility v1.1"
  PRINT "  (C)1997-98 Southern Software Inc."

  CheckDraftFile
  CheckRateFile
  CheckSetupFile
  
  PRINT
  PRINT
  PRINT " Processing Completed."
  PRINT
  PRINT " Press any key to continue.";
  LOCATE , , 1
  a$ = INPUT$(1)
  
  KillFile "UBRUN1ST.EXE"
  RUN "UBMENU.EXE"

SUB CheckDraftFile

  PRINT
  PRINT
  PRINT "  UB System Draft file: ";
  DIM ODraft AS oUBDraftRecType
  DIM NDraft AS UBDraftRecType

  OLen = LEN(ODraft)
  NLen = LEN(NDraft)
  
  DraftFile = FREEFILE
  OPEN "UBSDRAFT.DAT" FOR BINARY AS DraftFile
  DraftSize = LOF(DraftFile)
  CLOSE

  IF DraftSize > 0 THEN
    IF DraftSize < NLen THEN
      DraftFile = FREEFILE
      OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS DraftFile LEN = OLen
      GET DraftFile, 1, ODraft
      CLOSE
      NDraft.BANKDEST = ODraft.BANKDEST
      NDraft.BANKORIG = ODraft.BANKORIG
      NDraft.BANKNAME = ODraft.BANKNAME
      NDraft.BANKLOC = ODraft.BANKLOC
      LSET NDraft.COMPACCT = ""
      NDraft.FEDID = ODraft.FEDID
      NDraft.FEDPREFX = ODraft.FEDPREFX
      DraftFile = FREEFILE
      OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS DraftFile LEN = NLen
      PUT DraftFile, 1, NDraft
      CLOSE
      PRINT "Converted."
    ELSE
      PRINT "OK."
    END IF
  ELSE
    PRINT "FILE SIZE 0 bytes"
  END IF

END SUB

SUB CheckRateFile

  ChkByte$ = CHR$(1)
  PRINT
  PRINT "   UB System Rate file: ";
  DIM ORate AS oUBRateTblRecType
  DIM NRate AS UBRateTblRecType

  OLen = LEN(ORate)
  NLen = LEN(NRate)

  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM AS RateFile LEN = OLen
  RateFileSize& = LOF(RateFile)
  IF RateFileSize& > 0 THEN
    GET RateFile, 1, ORate
  ELSE
    PRINT "FILE SIZE 0 bytes"
    CLOSE
    EXIT SUB
  END IF
  CLOSE

  IF RIGHT$(ORate.RATEDESC, 1) <> ChkByte$ THEN
    KillFile "oUBRate.DAT"
    NAME "UBRate.DAT" AS "oUBRate.DAT"
    oRateFile = FREEFILE
    OPEN "oUBRate.DAT" FOR RANDOM SHARED AS oRateFile LEN = OLen
    NumOfRate = LOF(oRateFile) / OLen

    nRateFile = FREEFILE
    OPEN "UBRate.DAT" FOR RANDOM SHARED AS nRateFile LEN = NLen
    
    FOR Cnt = 1 TO NumOfRate
      GET oRateFile, Cnt, ORate
      NRate.RATECODE = ORate.RATECODE
      NRate.RATEDESC = ORate.RATEDESC
      NRate.ChkByte = ChkByte$
      NRate.MINAMT = ORate.MINAMT
      NRate.MINUNITS = ORate.MINUNITS
      NRate.MaxAmt = 0
      FOR TblCnt = 1 TO 10
        LSET NRate.TblBreaks(TblCnt) = ORate.TblBreaks(TblCnt)
      NEXT
      PUT nRateFile, Cnt, NRate
    NEXT
    PRINT "Converted."
  ELSE
    PRINT "OK."
  END IF
  CLOSE

END SUB

SUB CheckSetupFile

  PRINT
  PRINT "  UB System Setup file: ";
  DIM OSetup AS oUBSetupRecType
  DIM NSetup AS UBSetupRecType

  OLen = LEN(OSetup)
  NLen = LEN(NSetup)

  SetupFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR BINARY AS SetupFile
  SetupSize = LOF(SetupFile)
  CLOSE

  IF SetupSize > 0 THEN
    IF SetupSize < NLen THEN
      SetupFile = FREEFILE
      OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS SetupFile LEN = OLen
      GET SetupFile, 1, OSetup
      CLOSE

      NSetup.UTILNAME = OSetup.UTILNAME
      NSetup.DEFCITY = OSetup.DEFCITY
      NSetup.DEFSTATE = OSetup.DEFSTATE
      NSetup.ZIPCODE = OSetup.ZIPCODE
      NSetup.PreByBook = OSetup.PreByBook
      NSetup.RecpPort = OSetup.RecpPort
      NSetup.RECPDEFT = OSetup.RECPDEFT
      NSetup.ESTREAD = OSetup.ESTREAD
      NSetup.BANKDFT = OSetup.BANKDFT
      NSetup.UseSeq = OSetup.UseSeq
      NSetup.BILLCYCL = OSetup.BILLCYCL
      NSetup.DefLook = OSetup.DefLook
      NSetup.MethAcct = OSetup.MethAcct
      NSetup.SkipInactive = OSetup.SkipInactive
      NSetup.SkipSeparator = OSetup.SkipSeparator
      NSetup.Make99File = OSetup.Make99File
      NSetup.LowRead = OSetup.LowRead
      NSetup.HighRead = OSetup.HighRead
      NSetup.HHDEVICE = OSetup.HHDEVICE
      FOR zz = 1 TO 15
        LSET NSetup.Revenues(zz) = OSetup.Revenues(zz)
        NSetup.Revenues(zz).ProRate = "N"
        LSET NSetup.BillAcct(zz) = OSetup.BillAcct(zz)
        LSET NSetup.PayAcct(zz) = OSetup.PayAcct(zz)
        LSET NSetup.DepAcct(zz) = OSetup.DepAcct(zz)
      NEXT
      
      OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS SetupFile LEN = NLen
      PUT SetupFile, 1, NSetup
      CLOSE
      PRINT "Converted."
    ELSE
      PRINT "OK."
    END IF
  ELSE
    PRINT "FILE SIZE 0 bytes"
  END IF

END SUB

