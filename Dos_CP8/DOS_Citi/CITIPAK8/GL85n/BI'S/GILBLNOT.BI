'Gilbert Summit late notice

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    'AcctNum = IndexArray(Cnt).RecNum
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)

    PRINT #UBRpt, "~"  '; TAB(30); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; AcctNum;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 19); TAB(50); USING "########"; AcctNum;
    PRINT #UBRpt, TAB(62); NDate$
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(50); PDate$; TAB(64); USING "#####.##"; ABS(TotalAmt#)
    PRINT #UBRpt, TAB(3); NDate$; 'TAB(15); PrevDate$; TAB(26); DateRead$;
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PCnt = 0
    PRINT #UBRpt, "  IF WE DO NOT RECEIVE PAYMENT"; TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, "  BY "; PDate$; " WE WILL HAVE"; TAB(47); UBCustRec(1).Addr1
    PRINT #UBRpt, "  TO DISCONTINUE SERVICE."; TAB(47); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    IF Previous# <> 0 THEN
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; Current#
    PRINT #UBRpt, "                           --------------"

    PRINT #UBRpt,

    PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "  "; Message$
    PRINT #UBRpt,
    PRINT #UBRpt, "~"


