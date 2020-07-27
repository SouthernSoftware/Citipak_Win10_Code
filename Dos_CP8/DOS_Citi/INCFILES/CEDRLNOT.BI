'Cedar Bluff VA
'last revised 07-09-97

  CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
  
  TotalAmt# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

  PRINT #UBRpt, "~" '; TAB(47); USING "########"; PrintedCnt
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, USING "#######"; AcctNo&;
  PRINT #UBRpt, TAB(19); NDate$
  PRINT #UBRpt,
  PRINT #UBRpt, 'TAB(35); USING "Acct: #####"; AcctNo&
  
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(35); USING "Acct: #####"; AcctNo&
  
  IF UBCustRec(1).CurrRevAmts(1) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 10); TAB(21); USING "#######.##"; UBCustRec(1).CurrRevAmts(1)
  ELSE
    PRINT #UBRpt,
  END IF
  
  IF UBCustRec(1).CurrRevAmts(2) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 10); TAB(21); USING "#######.##"; UBCustRec(1).CurrRevAmts(2);
  END IF
  PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 27)
  
  IF UBCustRec(1).CurrRevAmts(3) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 10); TAB(21); USING "#######.##"; UBCustRec(1).CurrRevAmts(3);
  END IF
  PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ADDR1, 27)
  PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE

  IF UBCustRec(1).CurrRevAmts(4) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 10); TAB(21); USING "#######.##"; UBCustRec(1).CurrRevAmts(4)
  ELSE
    PRINT #UBRpt,
  END IF
  
  PRINT #UBRpt, TAB(2); "Previous:"; TAB(21); USING "#######.##"; TotalAmt#
  
  PRINT #UBRpt, TAB(35); UBCustRec(1).ServAddr
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(21); USING "#######.##"; TotalAmt#;
  PRINT #UBRpt, TAB(35); PDate$;
  PRINT #UBRpt, TAB(49); USING "#####.##"; TotalAmt#
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine1
  PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine2
  PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine3
  PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine4
  PRINT #UBRpt, "~"
'  PRINT #UBRpt, TAB(2); CustMsg$
  
  
  

