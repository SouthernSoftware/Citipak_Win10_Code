Attribute VB_Name = "Module2"
Type RePrintIndexType
  BillNum  As Long
  BillRec  As Long
End Type

Type UBXferInfoType
  DAcctNo   As String * 14
  DebitAmt  As Double
  DRecNo    As Integer
  DTitle    As String * 30
  CAcctNo   As String * 14
  CreditAmt As Double
  CRecNo    As Integer
  CTitle    As String * 30
End Type

Type GJXferRecType
  RevText    As String * 15
  BAcctInfo  As UBXferInfoType     'Billing Accounts
  PAcctInfo  As UBXferInfoType     'Payment Accounts
  DAcctInfo  As UBXferInfoType     'Deposit Accounts
End Type

'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount
'Type UBTransRecType
'   TransDate              As Integer      '
'   TransType              As Integer      '
'   TransDesc              As String * 21  'may change
'   TransAmt               As Double       'total revenue amount
'   RevAmt(1 To 15)        As Double       'Revenue amounts
'   TaxAmt(1 To 15)        As Single       'Tax Amounts
''01-20-97 Added meter types field to hold meter type at time of transaction
'   MtrTypes(1 To 7)       As Integer
''*******************
'   CurRead(1 To 7)        As Long         'Last/Current meter readings
'   PrevRead(1 To 7)       As Long         'Previous readings
'   EstRead(1 To 7)        As String * 1   'Y/N Flags for meter est's
'   BillNumber             As Long         'Number on the bill that Printed
'   ReadDate               As Integer
'   BillDate               As Integer
'   PastDueDate            As Integer
'   DraftDate              As Integer      'mowasa & plymouths bills. Can be changed
''111398
'   ProratePCT             As Integer
'   Filler1                As String * 1
'   EPPFlag                As String * 1   'Equal Payment Flag
'   CustStatus             As String * 1   'Customer Status at Time of Transaction
''020199
'   'PenAtBill              AS DOUBLE
'   EPPTrans               As Long         'Pointer to Equal Pay trans
'   Filler2                As String * 4
''****************
'   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check
'   OperatorNumber         As Integer      '
'   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
'   PrevTrans              As Long
'   ActUsage               As Long         'Changed for wadesboro
'   ActiveFlag             As Integer      'Valid transaction flag
'   RunBalance             As Double
'   CheckAmount            As Double
'   CashAmount             As Double
'   BillMsg                As String * 20
'   ApplyDepFlag           As String * 1
'   Posted2GL              As String * 1
'   PrevDate               As Integer
'   PenalFlag              As String * 1
'   TaxExempt              As String * 1
'   NONProfit              As String * 1
'End Type

'Trans Types
Const TranUtilityBill = 1          '   1=Utility bill
Const TranLateCharge = 2           '   2=late charge
Const TranReconnectFee = 3         '   3=reconnect fee
Const TranBillPayment = 4          '   4=Bill Payment
Const TranAppliedDeposit = 5       '   5=Applied Deposit
Const TranPenaltyCharge = 6        '   6=Penalty Charge
Const TranDepositPayment = 7       '   7=Deposit Payment
Const TranDraftPayment = 8         '   8=Draft Payment
Const TranRefundDeposit = 9        '   9=Refund Deposit
Const TranBeginBalance = 10        '  10=Beginning Balance
Const TranUpwardAdjustment = 11    '  11=Bill Adjustments
Const TranDownwardAdjustment = 12  '  12=Bill Adjustments
Const TranEPPReconcile = 13        '  13=EPP Reconcile Trans
Const TranMiscPayment = 99         '  99=Misc Payment

Const MtrWaterOnly = 1
Const MtrSewerOnly = 2
Const MtrCombined = 3
Const MtrElectric = 4
Const MtrDemand = 5
Const MtrGas = 6
Const MtrTouchRead = 7
Const MtrLightsService = 8

'Type MessLineType
'  Line As String * 59
'  LineDate As String * 10
'End Type
'
'Type UBMessRecType
'  MessLine(1 To 15) As MessLineType
'  CustRec As Long
'End Type

