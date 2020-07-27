DEFINT A-Z

DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION QPTrim$ (Text$)

DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB KillFile (FileName$)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubowner.bi'

  CLS

  UsingBook = True

ExportCustInfo:
  PRINT
  PRINT
  PRINT "Exporting utility data. . ."

  q$ = CHR$(34)
  qc$ = q$ + ","
  qcq$ = q$ + "," + q$

  IndexName$ = BookIndexFile
  UsingBook = True
  OKFlag = True
  BuckFmt$ = "######.##"

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize(IndexName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "PINEXPRT.ASC" FOR OUTPUT AS UBRpt

  FOR Cnt = 1 TO NumOfRecs
    LOCATE 8, 1: PRINT "Processing:"; Cnt; "of"; NumOfRecs;

    AcctNumber = IdxBuff(Cnt).RecNum
    GET UBCust, AcctNumber, UBCustRec(1)
    GET UBFile, AcctNumber, UBOwnerRec(1)

    '*************************************
    '   Main body of Printing goes here
    IF UBCustRec(1).DelFlag <> -1 AND UBCustRec(1).Status = "A" THEN
      ThisBook$ = UBCustRec(1).Book
      IF LEFT$(ThisBook$, 1) = "0" THEN
        WhatBook = VAL(RIGHT$(ThisBook$, 1))
      ELSE
        WhatBook = VAL(ThisBook$)
      END IF
      SELECT CASE WhatBook
      CASE 7 TO 22, 26, 30, 99
        Export& = Export& + 1
        GOSUB ExportThisAccount
      CASE ELSE
      END SELECT
    END IF
  NEXT

  CLOSE

  LOCATE 12, 1
  PRINT "Exported:"; Export&; "Accounts."
  PRINT " to FILE: PINEXPRT.ASC"

  LOCATE 16, 1
  PRINT "Done. . ."


'   WaitForAction
'
ExitMastCustListing:
  'COLOR 7, 0
END

ExportThisAccount:

  LocationNumber$ = QPTrim$(UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb)
  Zip$ = QPTrim$(UBCustRec(1).ZIPCODE)
  IF LEN(Zip$) > 5 THEN
    Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  END IF

  PRINT #UBRpt, q$; QPTrim$(STR$(AcctNumber));
  PRINT #UBRpt, qcq$; LocationNumber$;
  PRINT #UBRpt, qcq$; UBCustRec(1).Status;
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).CustName);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).ADDR1);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).ADDR2);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).ServAddr);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).CITY);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).STATE);
  PRINT #UBRpt, qcq$; Zip$;
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).HPHONE);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).WPHONE);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).SOSEC);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).DRVLIC);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUSTTYPE);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).Addr911);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).BillTo);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).POSTRTE);
  'BILLCYCL      AS INTEGER
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).ZONE);
  'SEQ           AS LONG
  'Page 2
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).CASHONLY);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).LATEFEE);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).CUTOFFYN);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).TAXEXPT);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).SRCIT);
  'EPPFlag

  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).USEDRAFT);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).AcctType);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).BANKNAME);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).BANKLOC);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).TRANSIT);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).BANKACCT);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).BILLCMNT);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).PAYCMNT);

  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).PUMPCODE);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE1);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).USERCODE2);
  IF UBCustRec(1).ProRatePCT > 0 THEN
    PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).ProRatePCT));
  ELSE
    PRINT #UBRpt, qcq$; "100";
  END IF

  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG1);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG2);
  PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).HHMSG3);

'Revs they only use 11
  FOR RCnt = 1 TO 11
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).Serv(RCnt).RateCode);
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).Serv(RCnt).RMtrType);
  NEXT

'flatrates
  FOR FCnt = 1 TO 4
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRDESC);
    IF UBCustRec(1).FlatRates(FCnt).FRAmt > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).FlatRates(FCnt).FRAmt));
    ELSE
      PRINT #UBRpt, qcq$; "0.00";
    END IF
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).FlatRates(FCnt).FRFreq);

    IF UBCustRec(1).FlatRates(FCnt).RevSrc > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).FlatRates(FCnt).RevSrc));
    ELSE
      PRINT #UBRpt, qcq$; "0";
    END IF

    IF UBCustRec(1).FlatRates(FCnt).NumMin > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).FlatRates(FCnt).NumMin));
    ELSE
      PRINT #UBRpt, qcq$; "0";
    END IF
  NEXT

'meters
  FOR MCnt = 1 TO 7
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRNUM);
    IF UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).LocMeters(MCnt).MTRMulti));
    ELSE
      PRINT #UBRpt, qcq$; "0";
    END IF
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType);
    PRINT #UBRpt, qcq$; QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRUnit);

    IF UBCustRec(1).LocMeters(MCnt).NumUser > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).LocMeters(MCnt).NumUser));
    ELSE
      PRINT #UBRpt, qcq$; "0";
    END IF
    IF UBCustRec(1).LocMeters(MCnt).InsDate > 0 THEN
      PRINT #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).InsDate);
    ELSE
      PRINT #UBRpt, qcq$; "??/??/????";
    END IF
    IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).LocMeters(MCnt).CurRead));
    ELSE
      PRINT #UBRpt, qcq$; "0";
    END IF

    IF UBCustRec(1).LocMeters(MCnt).PrevRead > 0 THEN
      PRINT #UBRpt, qcq$; QPTrim$(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead));
    ELSE
      PRINT #UBRpt, qcq$; "0";
    END IF

    IF UBCustRec(1).LocMeters(MCnt).CurDate > 0 THEN
      PRINT #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).CurDate);
    ELSE
      PRINT #UBRpt, qcq$; "??/??/????";
    END IF
    IF UBCustRec(1).LocMeters(MCnt).PastDate > 0 THEN
      PRINT #UBRpt, qcq$; Num2Date$(UBCustRec(1).LocMeters(MCnt).PastDate);
    ELSE
      PRINT #UBRpt, qcq$; "??/??/????";
    END IF
'    ReadFlag  AS STRING * 1    'hidden & protected
'    AvgUse    AS LONG          'hidden & protected
'    UseCnt    AS INTEGER       'hidden & protected
  NEXT

  PRINT #UBRpt, qcq$; USING BuckFmt$; UBCustRec(1).CurrBalance;
  PRINT #UBRpt, qcq$; USING BuckFmt$; UBCustRec(1).PrevBalance;

  FOR RCnt = 1 TO 11
    PRINT #UBRpt, qcq$; USING BuckFmt$; UBCustRec(1).CurrRevAmts(RCnt);
  NEXT

  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnLName);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).OwnFName);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ADDR1);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ADDR2);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).CITY);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).STATE);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).ZIPCODE);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).HPHONE);
  PRINT #UBRpt, qcq$; QPTrim$(UBOwnerRec(1).WPHONE);
  
  PRINT #UBRpt, q$

RETURN

