  FOR RevCnt = 1 TO NumOfRevs
    IF UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1 <> 0 OR UBPaymentRec(1).PaidOwed(RevCnt).AmtOwe1 <> 0 THEN
      LSET RecpRev$ = RevText$(RevCnt)
      PRINT #RHandle, RecpRev$; USING "$$#####.##"; UBPaymentRec(1).PaidOwed(RevCnt).AmtPd1
    END IF
  NEXT

