TYPE RePrintIndexType
  BillNum  AS LONG
  BillRec  AS LONG
END TYPE

TYPE UBXferInfoType
  DAcctNo   AS STRING * 14
  DebitAmt  AS DOUBLE
  DRecNo    AS INTEGER
  DTitle    AS STRING * 30
  CAcctNo   AS STRING * 14
  CreditAmt AS DOUBLE
  CRecNo    AS INTEGER
  CTitle    AS STRING * 30
END TYPE

TYPE GJXferRecType
  RevText    AS STRING * 15
  BAcctInfo  AS UBXferInfoType     'Billing Accounts
  PAcctInfo  AS UBXferInfoType     'Payment Accounts
  DAcctInfo  AS UBXferInfoType     'Deposit Accounts
END TYPE

'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount
TYPE UBTransRecType
   TransDate              AS INTEGER      '
   TransType              AS INTEGER      '
   TransDesc              AS STRING * 21  'may change
   TransAmt               AS DOUBLE       'total revenue amount
   RevAmt(1 TO 15)        AS DOUBLE       'Revenue amounts
   TaxAmt(1 TO 15)        AS SINGLE       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 TO 7)       AS INTEGER
'*******************
   CurRead(1 TO 7)        AS LONG         'Last/Current meter readings
   PrevRead(1 TO 7)       AS LONG         'Previous readings
   EstRead(1 TO 7)        AS STRING * 1   'Y/N Flags for meter est's
   BillNumber             AS LONG         'Number on the bill that Printed
   ReadDate               AS INTEGER
   BillDate               AS INTEGER
   PastDueDate            AS INTEGER
   DraftDate              AS INTEGER      'mowasa & plymouths bills. Can be changed
'111398
   ProratePCT             AS INTEGER
   Filler1                AS STRING * 1
   EPPFlag                AS STRING * 1   'Equal Payment Flag
   CustStatus             AS STRING * 1   'Customer Status at Time of Transaction
'020199
   'PenAtBill              AS DOUBLE
   EPPTrans               AS LONG         'Pointer to Equal Pay trans
   Filler2                AS STRING * 4
'****************
   PayTypeCode            AS INTEGER      'Payment Type:  1=Cash, 2=Check
   OperatorNumber         AS INTEGER      '
   CustAcctNo             AS LONG         'Pointer to RecNo in ubcust.dat
   PrevTrans              AS LONG
   ActUsage               AS LONG         'Changed for wadesboro
   ActiveFlag             AS INTEGER      'Valid transaction flag
   RunBalance             AS DOUBLE
   CheckAmount            AS DOUBLE
   CashAmount             AS DOUBLE
   BillMsg                AS STRING * 20
   ApplyDepFlag           AS STRING * 1
   Posted2GL              AS STRING * 1
   PrevDate               AS INTEGER
   PenalFlag              AS STRING * 1
   TaxExempt              AS STRING * 1
   NONProfit              AS STRING * 1
END TYPE

'Trans Types
CONST TranUtilityBill = 1          '   1=Utility bill
CONST TranLateCharge = 2           '   2=late charge
CONST TranReconnectFee = 3         '   3=reconnect fee
CONST TranBillPayment = 4          '   4=Bill Payment
CONST TranAppliedDeposit = 5       '   5=Applied Deposit
CONST TranPenaltyCharge = 6        '   6=Penalty Charge
CONST TranDepositPayment = 7       '   7=Deposit Payment
CONST TranDraftPayment = 8         '   8=Draft Payment
CONST TranRefundDeposit = 9        '   9=Refund Deposit
CONST TranBeginBalance = 10        '  10=Beginning Balance
CONST TranUpwardAdjustment = 11    '  11=Bill Adjustments
CONST TranDownwardAdjustment = 12  '  12=Bill Adjustments
CONST TranMiscPayment = 99         '  99=Misc Payment

CONST MtrWaterOnly = 1
CONST MtrSewerOnly = 2
CONST MtrCombined = 3
CONST MtrElectric = 4
CONST MtrDemand = 5
CONST MtrGas = 6
CONST MtrTouchRead = 7
CONST MtrLightsService = 8

TYPE MessLineType
  Line AS STRING * 59
  LineDate AS STRING * 10
END TYPE

TYPE UBMessRecType
  MessLine(1 TO 15) AS MessLineType
  CustRec AS LONG
END TYPE

