PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "                              Town of Montross"
PRINT #UBRpt, "                                P.O. BOX 126"
PRINT #UBRpt, "                             Montross, VA  22520"
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "     "; LongNDate$
PRINT #UBRpt, ""
PRINT #UBRpt, "     "; UBCustRec(1).CustName
PRINT #UBRpt, "     "; UBCustRec(1).ADDR1
PRINT #UBRpt, "     "; UBCustRec(1).ADDR2
PRINT #UBRpt, "     "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$; TAB(45); "Loc Number: "; Location$
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "     Dear "; UBCustRec(1).CustName
PRINT #UBRpt, ""
PRINT #UBRpt, "     Your UTILITY account as of "; NDate$; " is past due in the"
PRINT #UBRpt, "     amount of"; USING "$$####,#.##"; TotalBal#;
PRINT #UBRpt, "  Payment in full is expected by 5:00 P.M."
PRINT #UBRpt, "     on "; LongPDate$; " to avoid service Interruption."
PRINT #UBRpt, ""
PRINT #UBRpt, "     If payment is not received by the above stated time, legal"
PRINT #UBRpt, "     action will be taken to collect the balance due. There will"
PRINT #UBRpt, "     be no further notification before legal action is taken."
PRINT #UBRpt, ""
IF GotWater THEN
  PRINT #UBRpt, "     A reconnection fee of $100.00 will be assessed before water"
  PRINT #UBRpt, "     service is resumed."
ELSE
  PRINT #UBRpt,
  PRINT #UBRpt,
END IF
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "     Sincerely,"
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "     Brenda T. Reamy"
PRINT #UBRpt, CHR$(12)

