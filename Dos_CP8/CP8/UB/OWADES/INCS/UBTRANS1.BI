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

