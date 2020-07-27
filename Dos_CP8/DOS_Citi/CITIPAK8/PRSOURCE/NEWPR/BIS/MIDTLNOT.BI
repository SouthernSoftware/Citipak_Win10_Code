PRINT #UBRpt, TAB(3); "To: "; UBCustRec(1).CustName; "    "; LongNDate$
PRINT #UBRpt,
PRINT #UBRpt, TAB(3); "Our records indicate your water and sewer bill has not yet been"
PRINT #UBRpt, TAB(3); "paid.  Please be advised that if the amount indicated below is not"
PRINT #UBRpt, TAB(3); "paid within five (5) working days from the date of this notice, your"
PRINT #UBRpt, TAB(3); "water service will be discontinued and there will be a $20.00 charge"
PRINT #UBRpt, TAB(3); "to turn the water on again.  Bills must be paid before  NOON  on the"
PRINT #UBRpt, TAB(3); "cutoff date below.  AS DETERMINED BY TOWN COUNCIL, THERE WILL BE NO"
PRINT #UBRpt, TAB(3); "EXCEPTIONS TO THE CUTOFF DATE.  ALL BILLS MUST BE PAID IN FULL."
PRINT #UBRpt,
PRINT #UBRpt, TAB(53); "Water"; TAB(63); USING "####.##"; UBCustRec(1).CurrRevAmts(1)
PRINT #UBRpt, TAB(53); "Sewer"; TAB(63); USING "####.##"; UBCustRec(1).CurrRevAmts(2)
PRINT #UBRpt, TAB(53); "Surcharge"; TAB(63); USING "####.##"; UBCustRec(1).CurrRevAmts(3)
PRINT #UBRpt, TAB(3); USING "####"; AcctNo&;
PRINT #UBRpt, TAB(13); UBCustRec(1).CustName; TAB(53); "Misc"; TAB(63); USING "####.##"; UBCustRec(1).CurrRevAmts(5)
PRINT #UBRpt, TAB(13); UBCustRec(1).Addr1; TAB(53); "Sub Total"; TAB(63); USING "####.##"; Round#((UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) - UBCustRec(1).CurrRevAmts(4))
PRINT #UBRpt, TAB(13); LEFT$(UBCustRec(1).City, 20); " "; UBCustRec(1).State; " "; Zip$; TAB(53); "Penalty"; TAB(63); USING "####.##"; UBCustRec(1).CurrRevAmts(4)
PRINT #UBRpt, TAB(53); "Total"; TAB(63); USING "####.##"; Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
PRINT #UBRpt, TAB(3); "Cut Off Date: "; PDate$
PRINT #UBRpt, TAB(45); "Patricia Mullins - Town Clerk"
PRINT #UBRpt,
PRINT #UBRpt,

