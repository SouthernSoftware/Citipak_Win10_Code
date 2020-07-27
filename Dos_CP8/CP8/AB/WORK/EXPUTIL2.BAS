DEFINT A-T
DEFINT V-Z
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION FUsing$ (Bucks$, Format$)

  '$INCLUDE: 'newcust.bi'
  
  CONST False = 0, True = NOT False

  CLS
  PRINT "Exporting Utility Customer Information"

  q$ = CHR$(34)
  c$ = CHR$(44)

  DIM U AS NewUBCustRecType
  UBCustRecLen = LEN(U)          'Length of Cust Record Structure

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  EXFile = FREEFILE
  OPEN "FROMUTIL.ASC" FOR OUTPUT AS EXFile

  PRINT #EXFile, q$; "ACCTNO"; q$; c$; q$; "BOOK"; q$; c$; q$; "SEQNUMB"; q$; c$; q$; "Status"; q$; c$; q$; "OPENDATE"; q$; c$; q$; "SEARCH"; q$; c$; q$; "CUSTNAME"; q$; c$; q$; "ADDR1"; q$; c$; q$; "ADDR2"; q$; c$; q$; "SERVADDR"; q$; c$; q$;  _
"CITY"; q$; c$; q$; "STATE"; q$; c$; q$; "ZIPCODE"; q$; c$; q$; "HPHONE"; q$; c$; q$; "WPHONE"; q$; c$; q$; "SOSEC"; q$; c$; q$; "DRVLIC"; q$; c$; q$; "CUSTTYPE"; q$; c$; q$; "Addr911"; q$; c$; q$; "BillTo"; q$; c$; q$; "BILLCOPY"; q$; c$; q$;  _
"POSTRTE"; q$; c$; q$; "BILLCYCL"; q$; c$; q$; "ZONE"; q$; c$;
  PRINT #EXFile, q$; "SEQ"; q$; c$; q$; "CASHONLY"; q$; c$; q$; "LATEFEE"; q$; c$; q$; "CUTOFFYN"; q$; c$; q$; "TAXEXPT"; q$; c$; q$; "SRCIT"; q$; c$; q$; "EPPFlag"; q$; c$; q$; "USEDRAFT"; q$; c$; q$; "AcctType"; q$; c$; q$; "BANKNAME"; q$; c$; q$; _
 "BANKLOC"; q$; c$; q$; "TRANSIT"; q$; c$; q$; "BANKACCT"; q$; c$; q$; "BILLCMNT"; q$; c$; q$; "PAYCMNT"; q$; c$; q$; "PUMPCODE"; q$; c$; q$; "USERCODE1"; q$; c$; q$; "USERCODE2"; q$; c$; q$; "ProRatePCT"; q$; c$; q$; "HHMSG1"; q$; c$; q$; "HHMSG2"; _
 q$; c$; q$; "HHMSG3"; q$; c$;
  PRINT #EXFile, q$; "RateCode1"; q$; c$; q$; "RMtrType1"; q$; c$; q$; "RateCode2"; q$; c$; q$; "RMtrType2"; q$; c$; q$; "RateCode3"; q$; c$; q$; "RMtrType3"; q$; c$; q$; "RateCode4"; q$; c$; q$; "RMtrType4"; q$; c$; q$; "RateCode5"; q$; c$; q$;  _
"RMtrType5"; q$; c$; q$; "RateCode6"; q$; c$; q$; "RMtrType6"; q$; c$; q$; "RateCode7"; q$; c$; q$; "RMtrType7"; q$; c$; q$; "RateCode8"; q$; c$; q$; "RMtrType8"; q$; c$; q$; "RateCode9"; q$; c$; q$; "RMtrType9"; q$; c$;
  PRINT #EXFile, q$; "RateCode10"; q$; c$; q$; "RMtrType10"; q$; c$; q$; "RateCode11"; q$; c$; q$; "RMtrType11"; q$; c$; q$; "RateCode12"; q$; c$; q$; "RMtrType12"; q$; c$; q$; "RateCode13"; q$; c$; q$; "RMtrType13"; q$; c$; q$; "RateCode14"; q$; c$ _
; q$; "RMtrType14"; q$; c$; q$; "RateCode15"; q$; c$; q$; "RMtrType15"; q$; c$; q$; "FRDESC1"; q$; c$; q$; "FRAMT1"; q$; c$; q$; "FRFREQ1"; q$; c$; q$; "REVSRC1"; q$; c$; q$; "NumMin1"; q$; c$; q$; "FRDESC2"; q$; c$; q$; "FRAMT2"; q$; c$;
  PRINT #EXFile, q$; "FRFREQ2"; q$; c$; q$; "REVSRC2"; q$; c$; q$; "NumMin2"; q$; c$; q$; "FRDESC3"; q$; c$; q$; "FRAMT3"; q$; c$; q$; "FRFREQ3"; q$; c$; q$; "REVSRC3"; q$; c$; q$; "NumMin3"; q$; c$; q$; "FRDESC4"; q$; c$; q$; "FRAMT4"; q$; c$; q$;  _
"FRFREQ4"; q$; c$; q$; "REVSRC4"; q$; c$; q$; "NumMin4"; q$; c$; q$; "AMTOWED1"; q$; c$; q$; "TotAmtPD1"; q$; c$; q$; "PayAmt1"; q$; c$; q$; "RevSource1"; q$; c$; q$; "AMTOWED2"; q$; c$; q$; "TotAmtPD2"; q$; c$; q$; "PayAmt2"; q$; c$; q$;  _
"RevSource2"; q$; c$;
  PRINT #EXFile, q$; "MFEE1"; q$; c$; q$; "MFEE2"; q$; c$; q$; "MTRNUM1"; q$; c$; q$; "MTRMulti1"; q$; c$; q$; "MTRType1"; q$; c$; q$; "MTRUnit1"; q$; c$; q$; "NumUser1"; q$; c$; q$; "InsDate1"; q$; c$; q$; "CurRead1"; q$; c$; q$; "PrevRead1"; q$;  _
c$; q$; "CurDate1"; q$; c$; q$; "PastDate1"; q$; c$; q$; "ReadFlag1"; q$; c$; q$; "AvgUse1"; q$; c$; q$; "UseCnt1"; q$; c$; q$; "MTRNUM2"; q$; c$; q$; "MTRMulti2"; q$; c$; q$; "MTRType2"; q$; c$; q$; "MTRUnit2"; q$; c$; q$; "NumUser2"; q$; c$; q$;  _
"InsDate2"; q$; c$;
  PRINT #EXFile, q$; "CurRead2"; q$; c$; q$; "PrevRead2"; q$; c$; q$; "CurDate2"; q$; c$; q$; "PastDate2"; q$; c$; q$; "ReadFlag2"; q$; c$; q$; "AvgUse2"; q$; c$; q$; "UseCnt2"; q$; c$; q$; "MTRNUM3"; q$; c$; q$; "MTRMulti3"; q$; c$; q$; "MTRType3"; _
 q$; c$; q$; "MTRUnit3"; q$; c$; q$; "NumUser3"; q$; c$; q$; "InsDate3"; q$; c$; q$; "CurRead3"; q$; c$; q$; "PrevRead3"; q$; c$; q$; "CurDate3"; q$; c$; q$; "PastDate3"; q$; c$; q$; "ReadFlag3"; q$; c$; q$; "AvgUse3"; q$; c$; q$; "UseCnt3"; q$; c$; _

  PRINT #EXFile, q$; "MTRNUM4"; q$; c$; q$; "MTRMulti4"; q$; c$; q$; "MTRType4"; q$; c$; q$; "MTRUnit4"; q$; c$; q$; "NumUser4"; q$; c$; q$; "InsDate4"; q$; c$; q$; "CurRead4"; q$; c$; q$; "PrevRead4"; q$; c$; q$; "CurDate4"; q$; c$; q$; "PastDate4" _
; q$; c$; q$; "ReadFlag4"; q$; c$; q$; "AvgUse4"; q$; c$; q$; "UseCnt4"; q$; c$; q$; "MTRNUM5"; q$; c$; q$; "MTRMulti5"; q$; c$; q$; "MTRType5"; q$; c$; q$; "MTRUnit5"; q$; c$; q$; "NumUser5"; q$; c$; q$; "InsDate5"; q$; c$; q$; "CurRead5"; q$; c$;
  PRINT #EXFile, q$; "PrevRead5"; q$; c$; q$; "CurDate5"; q$; c$; q$; "PastDate5"; q$; c$; q$; "ReadFlag5"; q$; c$; q$; "AvgUse5"; q$; c$; q$; "UseCnt5"; q$; c$; q$; "MTRNUM6"; q$; c$; q$; "MTRMulti6"; q$; c$; q$; "MTRType6"; q$; c$; q$; "MTRUnit6"; _
 q$; c$; q$; "NumUser6"; q$; c$; q$; "InsDate6"; q$; c$; q$; "CurRead6"; q$; c$; q$; "PrevRead6"; q$; c$; q$; "CurDate6"; q$; c$; q$; "PastDate6"; q$; c$; q$; "ReadFlag6"; q$; c$; q$; "AvgUse6"; q$; c$; q$; "UseCnt6"; q$; c$; q$; "MTRNUM7"; q$; c$;
  PRINT #EXFile, q$; "MTRMulti7"; q$; c$; q$; "MTRType7"; q$; c$; q$; "MTRUnit7"; q$; c$; q$; "NumUser7"; q$; c$; q$; "InsDate7"; q$; c$; q$; "CurRead7"; q$; c$; q$; "PrevRead7"; q$; c$; q$; "CurDate7"; q$; c$; q$; "PastDate7"; q$; c$; q$;  _
"ReadFlag7"; q$; c$; q$; "AvgUse7"; q$; c$; q$; "UseCnt7"; q$; c$; q$; "CustPIN"; q$; c$; q$; "LastTrans"; q$; c$; q$; "CurrBalance"; q$; c$; q$; "PrevBalance"; q$; c$; q$; "CurrRevAmts1"; q$; c$; q$; "CurrRevAmts2"; q$; c$; q$; "CurrRevAmts3"; q$;  _
c$;
  PRINT #EXFile, q$; "CurrRevAmts4"; q$; c$; q$; "CurrRevAmts5"; q$; c$; q$; "CurrRevAmts6"; q$; c$; q$; "CurrRevAmts7"; q$; c$; q$; "CurrRevAmts8"; q$; c$; q$; "CurrRevAmts9"; q$; c$; q$; "CurrRevAmts10"; q$; c$; q$; "CurrRevAmts11"; q$; c$; q$;  _
"CurrRevAmts12"; q$; c$; q$; "CurrRevAmts13"; q$; c$; q$; "CurrRevAmts14"; q$; c$; q$; "CurrRevAmts15"; q$; c$; q$; "DepositAmt"; q$; c$; q$; "DelFlag"; q$; c$; q$; "PreNoteFlag"; q$; c$;
  PRINT #EXFile, q$; "WOLastTrans"; q$; c$; q$; "EstFlag"; q$; c$; q$; "MessageRec"; q$; c$; q$; "OldRec"; q$; c$; q$; "EPPLastTran"; q$; c$; q$; "NewNotes"; q$; c$; q$; "ChkByte"; q$
  
  FOR Cnt& = 1 TO NumOfRecs&
    GET UBFile, Cnt&, U
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    IF U.DelFlag = False THEN
      PRINT #EXFile, q$; QPTrim$(STR$(Cnt&)); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.Book); q$; c$; q$; QPTrim$(U.SeqNumb); q$; c$; q$; QPTrim$(U.Status); q$; c$;
      TheDate$ = Num2Date(U.OpenDate)
      IF INSTR(TheDate$, "%") > 0 THEN
        TheDate$ = ""
      END IF
      PRINT #EXFile, q$; TheDate$; q$; c$; q$; QPTrim$(U.Search); q$; c$; q$; QPTrim$(U.CustName); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.Addr1); q$; c$; q$; QPTrim$(U.Addr2); q$; c$; q$; QPTrim$(U.ServAddr); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.City); q$; c$; q$; QPTrim$(U.State); q$; c$; q$; QPTrim$(U.ZipCode); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.HPhone); q$; c$; q$; QPTrim$(U.WPhone); q$; c$; q$; QPTrim$(U.SoSec); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.DrvLic); q$; c$; q$; QPTrim$(U.CustType); q$; c$; q$; QPTrim$(U.Addr911); q$; c$;
      IF U.BillCopy < 1 THEN
        U.BillCopy = 1
      END IF
      PRINT #EXFile, q$; QPTrim$(U.BillTo); q$; c$; q$; QPTrim$(STR$(U.BillCopy)); q$; c$; q$; QPTrim$(U.PostRte); q$; c$;
      IF U.BillCycl < 0 THEN
        U.BillCycl = 0
      END IF
      IF U.Seq < 0 THEN
        U.Seq = 0
      END IF
      PRINT #EXFile, q$; QPTrim$(STR$(U.BillCycl)); q$; c$; q$; QPTrim$(U.Zone); q$; c$; q$; QPTrim$(STR$(U.Seq)); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.CashOnly); q$; c$; q$; QPTrim$(U.LateFee); q$; c$; q$; QPTrim$(U.CutOffYN); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.TaxExpt); q$; c$; q$; QPTrim$(U.Srcit); q$; c$; q$; QPTrim$(U.EPPFlag); q$; c$;
      'IF U.UseDraft < -1 THEN
      '  U.UseDraft = 0
      'END IF
      PRINT #EXFile, q$; QPTrim$(U.UseDraft); q$; c$; q$; QPTrim$(U.AcctType); q$; c$; q$; QPTrim$(U.BankName); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.BankLoc); q$; c$; q$; QPTrim$(U.Transit); q$; c$; q$; QPTrim$(U.BankAcct); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.BillCmnt); q$; c$; q$; QPTrim$(U.PayCmnt); q$; c$; q$; QPTrim$(U.PumpCode); q$; c$;
      IF U.ProRatePct < 0 THEN
        U.ProRatePct = 100
      END IF
      PRINT #EXFile, q$; QPTrim$(U.UserCode1); q$; c$; q$; QPTrim$(U.UserCode2); q$; c$; q$; QPTrim$(STR$(U.ProRatePct)); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.HHMsg1); q$; c$; q$; QPTrim$(U.HHMsg2); q$; c$; q$; QPTrim$(U.HHMsg3); q$; c$;

'Service's
      FOR SCnt = 1 TO 15
        PRINT #EXFile, q$; QPTrim$(U.Serv(SCnt).RateCode); q$; c$; q$; QPTrim$(U.Serv(SCnt).RMtrType); q$; c$;
      NEXT
'Flat Rates
      FOR FCnt = 1 TO 4
        IF U.FlatRates(FCnt).FRAmt < 0 THEN
          U.FlatRates(FCnt).FRAmt = 0
        END IF
        IF U.FlatRates(FCnt).RevSrc < 0 THEN
          U.FlatRates(FCnt).RevSrc = 0
        END IF
        PRINT #EXFile, q$; QPTrim$(U.FlatRates(FCnt).FRDesc); q$; c$; q$; QPTrim$(FUsing$(STR$(U.FlatRates(FCnt).FRAmt), "#####.##")); q$; c$;
        PRINT #EXFile, q$; QPTrim$(U.FlatRates(FCnt).FRFreq); q$; c$; q$; QPTrim$(STR$(U.FlatRates(FCnt).RevSrc)); q$; c$;
        IF U.FlatRates(FCnt).NumMin < 0 THEN
          U.FlatRates(FCnt).NumMin = 0
        END IF
        PRINT #EXFile, q$; QPTrim$(STR$(U.FlatRates(FCnt).NumMin)); q$; c$;
      NEXT

'Monthly
      FOR MCnt = 1 TO 2
        IF U.Monthly(MCnt).AmtOwed < 0 THEN
          U.Monthly(MCnt).AmtOwed = 0
        END IF
        PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.Monthly(MCnt).AmtOwed), "#####.##")); q$; c$; q$; QPTrim$(FUsing$(STR$(U.Monthly(MCnt).TotAmtPD), "#####.##")); q$; c$;
        IF U.Monthly(MCnt).PayAmt < 0 THEN
          U.Monthly(MCnt).PayAmt = 0
        END IF
        IF U.Monthly(MCnt).RevSource < 0 THEN
          U.Monthly(MCnt).RevSource = 0
        END IF
        PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.Monthly(MCnt).PayAmt), "#####.##")); q$; c$; q$; QPTrim$(STR$(U.Monthly(MCnt).RevSource)); q$; c$;
      NEXT
      IF U.MFee1 < 0 THEN U.MFee1 = 0
      IF U.MFee2 < 0 THEN U.MFee2 = 0
      PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.MFee1), "#####.##")); q$; c$; q$; QPTrim$(FUsing$(STR$(U.MFee2), "#####.##")); q$; c$;
'Location Meters
      FOR LCnt = 1 TO 7
        IF U.LocMeters(LCnt).MTRMulti < 0 THEN U.LocMeters(LCnt).MTRMulti = 0
        PRINT #EXFile, q$; QPTrim$(U.LocMeters(LCnt).MtrNum); q$; c$; q$; QPTrim$(STR$(U.LocMeters(LCnt).MTRMulti)); q$; c$;
        PRINT #EXFile, q$; QPTrim$(U.LocMeters(LCnt).MtrType); q$; c$; q$; QPTrim$(U.LocMeters(LCnt).MTRUnit); q$; c$;
        TheDate$ = Num2Date(U.LocMeters(LCnt).InsDate)
        IF INSTR(TheDate$, "%") > 0 THEN
          TheDate$ = ""
        END IF
        IF U.LocMeters(LCnt).NumUser < 0 THEN U.LocMeters(LCnt).NumUser = 0
        PRINT #EXFile, q$; QPTrim$(STR$(U.LocMeters(LCnt).NumUser)); q$; c$; q$; TheDate$; q$; c$;
        IF U.LocMeters(LCnt).CurRead < 0 THEN U.LocMeters(LCnt).CurRead = 0
        IF U.LocMeters(LCnt).PrevRead < 0 THEN U.LocMeters(LCnt).PrevRead = 0
        PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.LocMeters(LCnt).CurRead), "#########")); q$; c$; q$; QPTrim$(FUsing$(STR$(U.LocMeters(LCnt).PrevRead), "#########")); q$; c$;
        TheDate$ = Num2Date(U.LocMeters(LCnt).CurDate)
        IF INSTR(TheDate$, "%") > 0 THEN
          TheDate$ = ""
        END IF

        PRINT #EXFile, q$; TheDate$; q$; c$; q$;

        TheDate$ = Num2Date(U.LocMeters(LCnt).PastDate)
        IF INSTR(TheDate$, "%") > 0 THEN
          TheDate$ = ""
        END IF

        PRINT #EXFile, TheDate$; q$; c$;
        IF U.LocMeters(LCnt).AvgUse < 0 THEN U.LocMeters(LCnt).AvgUse = 0
        PRINT #EXFile, q$; QPTrim$(U.LocMeters(LCnt).ReadFlag); q$; c$; q$; QPTrim$(FUsing$(STR$(U.LocMeters(LCnt).AvgUse), "#########")); q$; c$;
        IF U.LocMeters(LCnt).UseCnt < 0 THEN U.LocMeters(LCnt).UseCnt = 0
        PRINT #EXFile, q$; QPTrim$(STR$(U.LocMeters(LCnt).UseCnt)); q$; c$;
      NEXT
      PRINT #EXFile, q$; QPTrim$(STR$(U.CustPin)); q$; c$; q$; QPTrim$(FUsing$(STR$(U.LastTrans), "#########")); q$; c$;
      IF U.CurrBalance < -100000 THEN
        U.CurrBalance = 0
      END IF
      IF U.PrevBalance < -100000 THEN
        U.PrevBalance = 0
      END IF
      PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.CurrBalance), "#####.##")); q$; c$; q$; QPTrim$(FUsing$(STR$(U.PrevBalance), "#####.##")); q$; c$;

      FOR RCnt = 1 TO 15
        PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.CurrRevAmts(RCnt)), "#####.##")); q$; c$;
      NEXT
      PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.DepositAmt), "#####.##")); q$; c$;
      PRINT #EXFile, q$; QPTrim$(STR$(U.DelFlag)); q$; c$; q$; QPTrim$(STR$(U.PreNoteFlag)); q$; c$;
      PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.WOLastTrans), "#########")); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.EstFlag); q$; c$; q$; QPTrim$(FUsing$(STR$(U.MessageRec), "#########")); q$; c$;
      PRINT #EXFile, q$; QPTrim$(FUsing$(STR$(U.OldRec), "#########")); q$; c$;
'HH New notes flag
      PRINT #EXFile, q$; QPTrim$(STR$(U.NewNotes)); q$; c$;
      PRINT #EXFile, q$; QPTrim$(U.ChkByte); q$
    END IF
  NEXT
  CLOSE
LOCATE 10, 5
PRINT "Export Completed"

