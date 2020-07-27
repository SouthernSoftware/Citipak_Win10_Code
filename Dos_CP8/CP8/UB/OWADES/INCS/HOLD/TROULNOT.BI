    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    AcctNum = IndexArray(Cnt).RecNum
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
    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 8);
        PRINT #UBRpt, TAB(36); USING "#####.##"; UBCustRec(1).CurrRevAmts(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 4
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
      CASE 5
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
      CASE 6
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
      CASE 7
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT

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
    PRINT #UBRpt, " Payment in full on or, before "; PDate$
    PRINT #UBRpt, " or service will be disconnected."
    PRINT #UBRpt, " "
    PRINT #UBRpt, "~"


'Troutman
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); "TOWN OF TROUTMAN"
'  PRINT #UBRpt, TAB(10); "P.O. BOX 26"
'  PRINT #UBRpt, TAB(10); "TROUTMAN, N.C. 28166"
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); "REMINDER NOTICE!!!"
'  PRINT #UBRpt, TAB(10); "IF YOUR CHECK HAS BEEN MAILED, PLEASE DISREGARD THIS NOTICE!"
'  PRINT #UBRpt, TAB(10); "OUR RECORDS INDICATE THAT YOUR WATER/SEWER BILL HAS NOT BEEN"
'  PRINT #UBRpt, TAB(10); "PAID. IF PAYMENT IS NOT RECIVED BY "; PDate$; " SERVICE WILL BE"
'  PRINT #UBRpt, TAB(10); "DISCONNECTED, WITHOUT FURTHER NOTICE, AND A $25.00 RECONNECT"
'  PRINT #UBRpt, TAB(10); "FEE WILL BE CHARGED IN ADDITION TO THE AMOUNT DUE. PLEASE GIVE"
'  PRINT #UBRpt, TAB(10); "THIS YOUR IMMEDIATE ATTENTION!!!!"
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); " Account# "; Acct$
'  PRINT #UBRpt, TAB(10); "Location# "; Location$; TAB(35); NDate$; TAB(50); "AMOUNT DUE:"; FUsing$(STR$(TotalBal#), "$$,######.##")
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); "THANK YOU, TOWN OF TROUTMAN"
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(46); UBCustRec(1).CustName
'  PRINT #UBRpt, TAB(46); UBCustRec(1).Addr1
'  PRINT #UBRpt, TAB(46); UBCustRec(1).Addr2
'  PRINT #UBRpt, TAB(46); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,

