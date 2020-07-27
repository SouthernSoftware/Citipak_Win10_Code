  PRINT #RptHandle, "   Total Cash Received: "; USING Fmt3$; TotalCash#
  PRINT #RptHandle, " Total Checks Received: "; USING Fmt3$; TotalCheck#
  PRINT #RptHandle, "                      -----------------"
  PRINT #RptHandle, "  Total Money Received: "; USING Fmt3$; TotalCash# + TotalCheck#
  PRINT #RptHandle, "Amount Applied to Acct: "; USING Fmt3$; TotalAmount#
  'PRINT #RptHandle, "    Total Change Given: "; USING Fmt3$; (TotalCash# + TotalCheck#) - TotalAmount#
  PRINT #RptHandle, "    Total Change Given: "; USING Fmt3$; TotChange#

