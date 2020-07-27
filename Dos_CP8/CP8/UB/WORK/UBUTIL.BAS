DEFINT A-Z
DECLARE SUB UNPreNote ()
DECLARE SUB FixRunBalance ()
DECLARE SUB SortTRHistory ()
DECLARE SUB ReadSeqDupes ()
DECLARE SUB DeleteCust (AcctNo&)
DECLARE SUB Archive (ThruYear%)
DECLARE SUB CheckRateTbl ()
DECLARE SUB UNDelete (AcctNo&)
DECLARE SUB SetReadFlag (ReadFlag%)
DECLARE FUNCTION GetProceed$ ()
DECLARE SUB ShowProcessing (Cnt&, NumOfRecs&)
DECLARE SUB FixAverageUse (AvgCnt)
DECLARE SUB ShowWarning ()
DECLARE SUB ReseqLoca (SeqStart&)
DECLARE SUB ReseqRead (SeqStart&)
DECLARE SUB TransDate (FromDate%, ToDate%)
DECLARE FUNCTION MakeSeqNumb$ (WhatSeq&)
DECLARE SUB SetBStatus (AcctNo&)
DECLARE FUNCTION GetCustCnt& ()
DECLARE SUB ShowHelp (DOITFLag)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION UTILFileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Num$)
DECLARE FUNCTION QPValL& (Num$)
DECLARE FUNCTION InCount% (Source$, SEARCH$)
DECLARE SUB DiskInfo (Drive$, Bytes%, Sectors%, FreeClust&, TotClust&)
DECLARE FUNCTION DiskRoom& (Drive$)
DECLARE SUB Parse (Work$, Array$(), ParmCnt)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB KillFile (FileName$)
DECLARE SUB UTLog (Text$)
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION FUsing$ (Numb$, Fmt$)

  CONST False = 0, True = NOT False
  
  STACK 5000
  
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubRate.BI'

TYPE ArcTranType
  RecNo      AS LONG
  TRType     AS INTEGER
  RunBalance AS DOUBLE
END TYPE


  REDIM Parms$(1)
  
  NumOfCust& = GetCustCnt&

  'GOTO DoUnPreNote

  CLS

  LOCATE , , 1
  PRINT
  PRINT TAB(17); "UBUTIL v5.01  (c)1997-2000 Southern Software Inc."
  PRINT
  
  Cmd$ = COMMAND$
  
  IF INSTR(Cmd$, "HELP") OR LEN(Cmd$) = 0 THEN
    ShowHelp False
    GOTO NormalExit
  END IF
  
  Parse Cmd$, Parms$(), ParmCnt
  
  SELECT CASE Parms$(1)      'Check the first parm

  CASE "UNPRENOTE"
    UNPreNote
  CASE "TRANSDATE"
    IF ParmCnt < 3 THEN
      ErrParm = ParmCnt + 1
    ELSE
      FromDate = Date2Num(Parms$(2))
      ToDate = Date2Num(Parms$(3))
      IF (FromDate < 0) OR (ToDate < 0) THEN
        IF FromDate < 0 THEN
          ErrParm = 2
        ELSE
          ErrParm = 3
        END IF
      ELSE
        TransDate FromDate, ToDate
      END IF
    END IF
  CASE "RESEQUENCE"
    IF ParmCnt < 3 THEN
      ErrParm = ParmCnt + 1
    ELSE
      SeqStart& = QPValL(Parms$(3))
      SELECT CASE Parms$(2)
      CASE "LOCATION"
        ReseqLoca SeqStart&
      CASE "READ"
        ReseqRead SeqStart&
      CASE ELSE
        ErrParm = 2
      END SELECT
    END IF
  CASE "READFLAG"
    IF ParmCnt < 2 THEN
      ErrParm = ParmCnt + 1
    ELSE
      SELECT CASE Parms$(2)
      CASE "SET"
        ReadFlag = True
      CASE "CLEAR"
        ReadFlag = False
      CASE ELSE
        ErrParm = 2
      END SELECT
      IF ErrParm = 0 THEN
        SetReadFlag ReadFlag
      END IF
    END IF
  CASE "UNDELETE"
    IF ParmCnt < 2 THEN      'don't know what account to change
      ErrParm = 2            'error in parm 2
    ELSE
      AcctNo& = QPValL(Parms$(2))
      SELECT CASE AcctNo&
      CASE IS < 1, IS > NumOfCust&
        ErrParm = 2
      CASE ELSE
        UNDelete AcctNo&
      END SELECT
    END IF
  CASE "DELETE"
    IF ParmCnt < 2 THEN      'don't know what account to change
      ErrParm = 2            'error in parm 2
    ELSE
      AcctNo& = QPValL(Parms$(2))
      SELECT CASE AcctNo&
      CASE IS < 1, IS > NumOfCust&
        ErrParm = 2
      CASE ELSE
        DeleteCust AcctNo&
      END SELECT
    END IF

  CASE "AVERAGEUSE"
    IF ParmCnt < 2 THEN      'don't know what account to change
      ErrParm = 2            'error in parm 2
    ELSE
      AvgCnt = QPValL(Parms$(2))
      SELECT CASE AvgCnt
      CASE IS < 1, IS > 24
        ErrParm = 2
      CASE ELSE
        FixAverageUse AvgCnt
      END SELECT
    END IF

'  CASE "READSEQDUPE"
'    ReadSeqDupes
'  CASE "ARCHIVE"
'    IF ParmCnt < 2 THEN      'don't know what to archive
'      ErrParm = 2            'error in parm 2
'    ELSE
'      ThisYear = QPValI(RIGHT$(DATE$, 4))
'      ThruYear = QPValI(Parms$(2))
'      IF ThruYear >= ThisYear THEN
'        ErrParm = 2            'error in parm 2
'      ELSE
'        Archive ThruYear
'      END IF
'    END IF
'
'  CASE "FIXRUNBALANCE"
'    FixRunBalance

  CASE "HISTORYSORT"
    SortTRHistory
  CASE ELSE
    ErrParm = 1
  END SELECT
  
  IF ErrParm > 0 THEN
    GOSUB ShowError
    ShowHelp True
    GOTO AbortExit
  END IF
  
NormalExit:
  PRINT
  PRINT
  PRINT "   UBUTIL Program completed."

AbortExit:
  CLOSE
  END
  
ShowError:
  ErrText$ = " COMMAND LINE: "
  PRINT ErrText$;
  FOR zz = 1 TO ParmCnt
    PRINT Parms$(zz); " ";
  NEXT
  FOR zz = 1 TO ErrParm - 1
    TLen = TLen + LEN(Parms$(zz)) + 1
  NEXT
  TLen = TLen + LEN(ErrText$) + 1
  PRINT
  PRINT "ERROR POSTION:";
  LOCATE , TLen
  IF UBOUND(Parms$) < ErrParm THEN
    ParmLen = 1
  ELSE
    ParmLen = LEN(Parms$(ErrParm))
  END IF
  PRINT STRING$(ParmLen, 24)
RETURN

DoUnPreNote:
  REDIM UBCust(1)   AS NewUBCustRecType
  UBCustLen = LEN(UBCust(1))


  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS #CustFile LEN = UBCustLen
  FOR Cnt = 1 TO NumOfCust&
    GET CustFile, Cnt, UBCust(1)
    UBCust(1).PreNoteFlag = 0
    PUT CustFile, Cnt, UBCust(1)

  NEXT
  CLOSE

SUB DeleteCust (AcctNo&)

  REDIM UBCust(1)   AS NewUBCustRecType
  UBCustLen = LEN(UBCust(1))

  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS #CustFile LEN = UBCustLen
  GET #CustFile, AcctNo&, UBCust(1)
  CustBal# = uRound#(UBCust(1).CurrBalance + UBCust(1).PrevBalance)

  PRINT "   Delete An Account"
  PRINT
  PRINT "   Account: "; AcctNo&; "  "; QPTrim$(UBCust(1).CustName)
  PRINT "   Balance: "; FUsing$(STR$(CustBal#), "######.##")
  PRINT "   Deposit: "; FUsing$(STR$(UBCust(1).DepositAmt), "######.##")
  PRINT
  PRINT "   ARE YOU SURE?  "
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT "Y"
    UBCust(1).DelFlag = True
    PUT #CustFile, AcctNo&, UBCust(1)
    PRINT
    PRINT "   Account Has Been Deleted."
    UTLog "DELETED ACCOUNT:" + STR$(AcctNo&) + "  " + QPTrim$(UBCust(1).CustName)
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   ACCOUNT NOT DELETED."
  END SELECT
  CLOSE CustFile

DeleteExit:


END SUB

SUB FixAverageUse (AvgCnt)

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  ShowWarning
  PRINT "   Recalc Average Usage, using last"; AvgCnt; "periods."
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB GOFixAverage
    PRINT
    PRINT
    PRINT "   Adjust Average Completed."
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   NO AVERAGES WERE ADJUSTED"
  END SELECT
EXIT SUB

GOFixAverage:
  PRINT

  UBTran = FREEFILE
  OPEN "UBTrans.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    ShowProcessing Cnt&, NumOfRecs&
    GET UBFile, Cnt&, UBCustRec(1)
    REDIM TotalUse(1 TO 7) AS LONG
    REDIM UseCnt(1 TO 7) AS INTEGER
    DidCnt = 0
    LastTran& = UBCustRec(1).LastTrans
    DO WHILE LastTran& > 0
      GET UBTran, LastTran&, UBTranRec(1)
      IF UBTranRec(1).TransType = TranUtilityBill THEN
        FOR MCnt = 1 TO 7
          IF UBTranRec(1).CurRead(MCnt) > 0 THEN
            TotalUse(MCnt) = TotalUse(MCnt) + (UBTranRec(1).CurRead(MCnt) - UBTranRec(1).PrevRead(MCnt))
            UseCnt(MCnt) = UseCnt(MCnt) + 1
          END IF
        NEXT
        DidCnt = DidCnt + 1
        IF DidCnt >= AvgCnt THEN
          EXIT DO
        END IF
      END IF
      LastTran& = UBTranRec(1).PrevTrans
    LOOP

    FOR MCnt = 1 TO 7
      IF TotalUse(MCnt) > 0 THEN
        UBCustRec(1).LocMeters(MCnt).AvgUse = TotalUse(MCnt) / UseCnt(MCnt)
        UBCustRec(1).LocMeters(MCnt).UseCnt = UseCnt(MCnt)
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)

  NEXT
  CLOSE

RETURN

END SUB

FUNCTION GetCustCnt&
  REDIM UBCust(1)   AS NewUBCustRecType
  UBCustLen = LEN(UBCust(1))
  GetCustCnt& = UTILFileSize&("UBCUST.DAT") / UBCustLen
  CLOSE
  ERASE UBCust
END FUNCTION

FUNCTION GetProceed$
  PRINT
  PRINT "   Proceed (Y/N)? ";
  GetProceed$ = UCASE$(INPUT$(1))
END FUNCTION

FUNCTION MakeSeqNumb$ (WhatSeq&) STATIC
  SEQ$ = "000000" + QPTrim$(STR$(WhatSeq&))
  MakeSeqNumb$ = RIGHT$(SEQ$, 6)
END FUNCTION

SUB Parse (Work$, Array$(), ParmCnt)
  Delim$ = " "
  BeginPtr = 1
  Element = 1
  FOR X = 1 TO LEN(Delim$)
    Counter = Counter + InCount(Work$, MID$(Delim$, X, 1))
  NEXT
  REDIM Array$(Counter + 1)
  FOR EndPtr = 1 TO LEN(Work$)
    IF INSTR(Delim$, MID$(Work$, EndPtr, 1)) THEN
      Array$(Element) = MID$(Work$, BeginPtr, EndPtr - BeginPtr)
      Element = Element + 1
      BeginPtr = EndPtr + 1
    END IF
  NEXT
  Array$(Element) = MID$(Work$, BeginPtr)
  FOR Cnt = 1 TO Element
    Tmp$ = QPTrim$(Array$(Cnt))
    IF LEN(Tmp$) > 0 THEN
      ParmCnt = ParmCnt + 1
      REDIM PRESERVE TArray$(1 TO ParmCnt)
      TArray$(ParmCnt) = Tmp$
    END IF
  NEXT
  REDIM Array$(1 TO ParmCnt)
  FOR Cnt = 1 TO ParmCnt
    Array$(Cnt) = TArray$(Cnt)
  NEXT
  
END SUB

SUB ReadSeqDupes

  PRINT "   Search for Read Sequence Duplicates?"
  PRINT
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB SearchReadSeq
    PRINT
    PRINT "   Resequencing Completed."
    PRINT
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   NO READING SEQUENCES WERE CHANGED"
  END SELECT

EXIT SUB

SearchReadSeq:
RETURN


END SUB

SUB ReseqLoca (SeqStart&)

  ShowWarning
  PRINT "   Resequence Locations"
  PRINT
  PRINT "   Increment By: "; SeqStart&
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB GOReseqLocs
    PRINT
    PRINT
    PRINT "   Resequencing Completed."
    PRINT
    PRINT "   YOU MUST REINDEX THE UTILITY FILES"
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   NO LOCATIONS WERE RESEQUENCED"
  END SELECT
  
EXIT SUB
  
GOReseqLocs:
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = UTILFileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  CurrBook$ = ""
  WhatSeq& = 0
  
  FOR Cnt& = 1 TO IdxNumOfRecs
    ShowProcessing Cnt&, CLNG(IdxNumOfRecs)
    ThisCust& = IdxBuff(Cnt&).RecNum
    GET UBFile, ThisCust&, UBCustRec(1)
    IF UBCustRec(1).DelFlag THEN
      UBCustRec(1).Book = ""
      UBCustRec(1).SEQNUMB = ""
      PUT UBFile, ThisCust&, UBCustRec(1)
    ELSE
      ThisBook$ = QPTrim$(UBCustRec(1).Book)
      IF LEN(ThisBook$) > 0 THEN
        IF CurrBook$ <> ThisBook$ THEN
          CurrBook$ = ThisBook$
          WhatSeq& = SeqStart&
          UBCustRec(1).SEQNUMB = MakeSeqNumb$(WhatSeq&)
        ELSEIF CurrBook$ = ThisBook$ THEN
          WhatSeq& = WhatSeq& + SeqStart&
          UBCustRec(1).SEQNUMB = MakeSeqNumb$(WhatSeq&)
        END IF
        PUT UBFile, ThisCust&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE
  
  ERASE UBCustRec, IdxBuff
RETURN

END SUB

SUB ReseqRead (SeqStart&)
  
  ShowWarning
  PRINT "   Resequence Meter Read Sequence"
  PRINT
  PRINT "   Increment By: "; SeqStart&
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB GOReseqRead
    PRINT
    PRINT
    PRINT "   Resequencing Completed."
    PRINT
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   NO READING SEQUENCES WERE CHANGED"
  END SELECT
EXIT SUB
  
  
GOReseqRead:
  PRINT
  REDIM UBCust(1) AS NewUBCustRecType
  CustLen = LEN(UBCust(1))
  
  IdxRecLen = 4 'we are using a long integer
  
  OPEN "UBCUSTBK.IDX" FOR RANDOM AS #1 LEN = 4
  NumOfRecs& = LOF(1) \ 4
  REDIM IdxBuff(1 TO NumOfRecs&) AS LONG
  FOR Cnt& = 1 TO NumOfRecs&
    GET #1, Cnt&, IdxBuff(Cnt&)
  NEXT
  CLOSE
  
  OPEN "ubcust.dat" FOR RANDOM AS #1 LEN = CustLen
  NumOfRecs& = LOF(1) \ CustLen
  FOR Cnt& = 1 TO NumOfRecs&
    ShowProcessing Cnt&, NumOfRecs&
    ThisCust& = IdxBuff(Cnt&)
    GET #1, ThisCust&, UBCust(1)
    IF UBCust(1).DelFlag = 0 THEN
      SeqCnt& = SeqCnt& + 1
      SEQ& = SeqCnt& * 50
      UBCust(1).SEQ = SEQ&
      PUT #1, ThisCust&, UBCust(1)
    END IF
  NEXT
  CLOSE
  
  ERASE IdxBuff, UBCust
RETURN
  
END SUB

SUB SetReadFlag (ReadFlag)

  ShowWarning
  PRINT "   Meter Reading Flags"
  PRINT
  SELECT CASE ReadFlag
  CASE True
    PRINT "   Set Meters as Read?"
  CASE ELSE
    PRINT "   Set Meters as UNRead?"
  END SELECT
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB GOSetReadFlags
    PRINT
    PRINT
    SELECT CASE ReadFlag
    CASE True
      PRINT "   Meter Read Flags Set."
    CASE ELSE
      PRINT "   Meter Read Flags Cleared."
    END SELECT
    PRINT
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   NO READ FLAGS WERE CHANGED."
  END SELECT

EXIT SUB

GOSetReadFlags:
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    ShowProcessing Cnt&, NumOfRecs&
    GET UBFile, Cnt&, UBCustRec(1)
    FOR MCnt = 1 TO 7
      IF ReadFlag THEN
        UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y"
      ELSE
        UBCustRec(1).LocMeters(MCnt).ReadFlag = ""
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
RETURN
END SUB

SUB ShowHelp (DOITFLag)
  PRINT
  PRINT TAB(8); "Syntax:  UBUTIL COMMAND <arg1> <arg2> <arg3...>"
  PRINT
  PRINT TAB(6); "COMMANDS:";
  PRINT TAB(17); "    TRANSDATE  <FromDate> <ToDate>"
  PRINT TAB(17); "   RESEQUENCE  <Location | Read> <STEP>"
  PRINT TAB(17); "     READFLAG  <Set | Clear>"
  PRINT TAB(17); "     UNDELETE  <AccountNumber>"
  PRINT TAB(17); "       DELETE  <AccountNumber>"
  PRINT TAB(17); "   AVERAGEUSE  <Period Count>"
  PRINT TAB(17); "    UNPRENOTE  <none>"

'  PRINT TAB(17); "FIXRUNBALANCE  <none>"
'  PRINT TAB(17); "READSEQDUPE  <none>"
'  PRINT TAB(17); "  HISTORYSORT  <none>"
'  PRINT TAB(17); "   ARCHIVE  <Thru Year>"
  PRINT
  IF DOITFLag THEN
    PRINT TAB(7); "WARNING: THIS PROGRAM SHOULD ONLY BE USED WHEN"
    PRINT TAB(7); "         INSTRUCTED TO DO SO BY SOFTWARE SUPPORT!"
    PRINT
    PRINT TAB(16); "Contact Southern Software for Assistance."
  END IF
END SUB

SUB ShowProcessing (Cnt&, NumOfRecs&) STATIC
  LOCATE , 1, 0
  PRINT "   Processing:"; Cnt&; " of"; NumOfRecs&;
END SUB

SUB ShowWarning
PRINT
PRINT TAB(4); "WARNING: ALL OTHER UTILITY OPERATIONS MUST STOP UNTIL"
PRINT TAB(4); "         THIS PROGRAM HAS COMPLETED PROCESSING. IF NOT"
PRINT TAB(4); "         YOUR DATA WILL PROBABLY BE DESTORYED!!!!!!!!"
PRINT
END SUB

SUB SortTRHistory

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  ShowWarning
  PRINT "   Sort Transaction History"
  PRINT
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB GOSortTransactions
    PRINT
    PRINT
    PRINT "   Transaction Sort Completed."
    PRINT
    PRINT "   YOU MUST REINDEX THE UTILITY FILES"
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   HISTORY SORT ABORTED!"
  END SELECT

  EXIT SUB

GOSortTransactions:
  PRINT
  IF Exist("UBTRANS.$$$") THEN
    KillFile "UBTRANS.$$$"
  END IF
  NAME "UBTRANS.DAT" AS "UBTRANS.$$$"

  UBoTran = FREEFILE
  OPEN "UBTrans.$$$" FOR RANDOM SHARED AS UBoTran LEN = UBTranRecLen

  UBTran = FREEFILE
  OPEN "UBTrans.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    ShowProcessing Cnt&, NumOfRecs&
    GET UBFile, Cnt&, UBCustRec(1)
    REDIM TRList(1 TO 1) AS LONG
    TRCnt = 0
    LastTran& = UBCustRec(1).LastTrans
    DO WHILE LastTran& > 0
      GET UBoTran, LastTran&, UBTranRec(1)
      TRCnt = TRCnt + 1
      REDIM PRESERVE TRList(1 TO TRCnt) AS LONG
      TRList(TRCnt) = LastTran&
      'PUT UBTran, , UBTranRec(1)
      LastTran& = UBTranRec(1).PrevTrans
    LOOP
    IF TRCnt > 0 THEN
      FOR TR = TRCnt TO 1 STEP -1
        GET UBoTran, TRList(TR), UBTranRec(1)
        PUT UBTran, , UBTranRec(1)
      NEXT
    END IF
  NEXT
  CLOSE
RETURN

END SUB

SUB TransDate (FromDate, ToDate)

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  ShowWarning
  PRINT "   Change Transaction Dates"
  PRINT
  PRINT "   From: "; Num2Date$(FromDate); " to "; Num2Date$(ToDate)
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT "Y"
    PRINT
    UBFile = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
    TNumOfRecs& = LOF(UBFile) / UBTranRecLen
    FOR Cnt& = 1 TO TNumOfRecs&
      ShowProcessing Cnt&, TNumOfRecs&
      GET UBFile, Cnt&, UBTranRec(1)
      'IF UBTranRec(1).TransType = TranUtilityBill THEN
        IF UBTranRec(1).TransDate = FromDate THEN
          TTCnt = TTCnt + 1
          UBTranRec(1).TransDate = ToDate
          PUT UBFile, Cnt&, UBTranRec(1)
        END IF
      'END IF
    NEXT
    CLOSE
    PRINT
    PRINT
    PRINT "      Changed: "; TTCnt
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT
    PRINT "   NO TRANSACTION DATES CHANGED"
  END SELECT
  ERASE UBTranRec

END SUB

SUB UNDelete (AcctNo&)

  REDIM UBCust(1)   AS NewUBCustRecType
  UBCustLen = LEN(UBCust(1))

  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS #CustFile LEN = UBCustLen
  GET #CustFile, AcctNo&, UBCust(1)
  'CLOSE
  PRINT "   Undelete An Account"
  PRINT
  PRINT "   Account: "; AcctNo&; "   "; UBCust(1).CustName
  PRINT
  IF UBCust(1).DelFlag = 0 THEN
    CLOSE
    PRINT "   CUSTOMER ACCOUNT IS NOT DELETED!!"
    GOTO NoDeleteExit
  ELSE
    PRINT "   Undelete This Account?"
  END IF
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT "Y"
    UBCust(1).DelFlag = 0
    PUT #CustFile, AcctNo&, UBCust(1)
    PRINT
    PRINT "   Account Has Been Undeleted."
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   ACCOUNT STILL DELETED."
  END SELECT
  CLOSE CustFile

NoDeleteExit:


END SUB

SUB UNPreNote

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  ShowWarning
  PRINT "   Clear prenote Flags"
  PRINT
  Ok$ = GetProceed$
  SELECT CASE Ok$
  CASE "Y"
    PRINT "Y"
    PRINT
    UBCFile = FREEFILE
    OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
    CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
    FOR Cnt& = 1 TO CNumOfRecs&
      ShowProcessing Cnt&, CNumOfRecs&
      GET UBCFile, Cnt&, UBCustRec(1)
      IF UBCustRec(1).PreNoteFlag <> 0 THEN
        UBCustRec(1).PreNoteFlag = 0
        TCnt = TCnt + 1
        PUT UBCFile, Cnt&, UBCustRec(1)
      END IF
    NEXT
    CLOSE
    PRINT
    PRINT
    PRINT "      Changed: "; TCnt
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT
    PRINT "   NO CUSTOMERS CHANGED"
  END SELECT

END SUB

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

'This function is a replacement for the QuickPak FileSize function.
'Due to the way Windows NT updates a file's size in the directory, an
'error can occur using DOS Function 4Eh (Find first file service) to
'read a file's size from the Directory. You can force Windows NT to
'commit the directory info by just opening the file again.
FUNCTION UTILFileSize& (FileName$)
  IF Exist(FileName$) THEN
    Handle = FREEFILE
    OPEN FileName$ FOR BINARY AS #Handle
    UTILFileSize& = LOF(Handle)
    CLOSE Handle
  ELSE
    UTILFileSize& = 0
  END IF
END FUNCTION

SUB UTLog (Text$) STATIC

  IF NOT BeenDone THEN
    BeenDone = True
    Today$ = DATE$
    Today$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
  END IF

  TheTime$ = TIME$
  IF LEFT$(TheTime$, 1) = "0" THEN
    Hour = VAL(MID$(TheTime$, 2, 1))
  ELSE
    Hour = VAL(MID$(TheTime$, 1, 2))
  END IF

  SELECT CASE Hour
  CASE IS > 11
    Hour = Hour - 12
    IF Hour = 0 THEN Hour = 12
    AmPm$ = "pm"
  CASE 1 TO 12
    AmPm$ = "am"
  CASE 0
    Hour = 12
    AmPm$ = "am"
  END SELECT
  SELECT CASE Hour
    CASE 1 TO 9
      Hour$ = "0" + QPTrim$(STR$(Hour))
    CASE ELSE
      Hour$ = QPTrim$(STR$(Hour))
  END SELECT
  TheTime$ = Hour$ + ":" + MID$(TheTime$, 4) + AmPm$
  LogFile = FREEFILE
  OPEN "UBLOG.DAT" FOR APPEND SHARED AS LogFile LEN = 255
  PRINT #LogFile, "UBUTIL: "; Today$; " @"; TheTime$
  PRINT #LogFile, "        "; Text$
  CLOSE #LogFile
  Text$ = "": TheTime$ = ""
END SUB

