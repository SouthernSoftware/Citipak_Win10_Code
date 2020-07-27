Attribute VB_Name = "modUBExportData"
DefInt A-Z

Dim FlipFlag As Boolean
Dim BtnFnt As Double
Dim intCount As Long
Dim RptHandle As Integer
Dim RptName As String
Dim FileName As String
Dim PipeSymbol As String
Dim RecLen As Integer

Public Sub ProcessUBData()
  StartPath = App.Path
  PipeSymbol = "|"
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Convert Deposits
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim UsingBook As Boolean, UsingName As Boolean
  Dim AcctNumber As Long, UBCust As Integer, UsingAcct As Boolean
  Dim IndexName As String, UBRpt As Integer, SEQNUMB As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, OKToSkip As Boolean
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, Handle As Integer
  Dim Cnt As Long, TDeposit As Double, ToPrint As String
  Dim BOOK As String, CustCnt As Long, ReportFile As String
  ReDim RevAmts(1 To 15) As Double
  Dim UBTransRecLen As Integer, NextTranRecs As Long
  Dim TransDate As Integer, TransAmt As Double
  Dim RevCnt As Integer
  Dim UBTran As Integer, NumOfTranRecs As Long, PrevLastTrans As Long
  Dim TotalDepAmt As Double, LastTran As Long
  Dim UBCustRec As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec)
  Dim TUB As UBTransRecType
  UBTransRecLen = Len(TUB)
 
  UBTran = FreeFile
  Open UBData + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
  
  ToPrint$ = ""

 ' NumOfRecs = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen

  UBCust = FreeFile
  Open UBData + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs = LOF(UBCust) / UBCustRecLen
  
  PipeSymbol = "|"

  FileName = "\UBDeposit.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  For Cnt = 1 To NumOfRecs
    AcctNumber = Cnt
    Get UBCust, AcctNumber, UBCustRec
      If UBCustRec.DelFlag = 0 Then
        If Round#(UBCustRec.DepositAmt) <> 0 Then
          TotalDepAmt# = 0
        ReDim RevAmts(1 To 15) As Double
        LastTran& = UBCustRec.LastTrans
        If LastTran& > 0 Then
          Do
            Get #UBTran, LastTran&, TUB
            If TUB.TransType = TranDepositPayment Then
              For RevCnt = 1 To 15
                If TUB.RevAmt(RevCnt) > 0 Then
                  RevAmts(RevCnt) = Round#(RevAmts(RevCnt) + TUB.RevAmt(RevCnt))
                  TotalDepAmt# = Round#(TotalDepAmt# + TUB.RevAmt(RevCnt))
                End If
              Next
            ElseIf (TUB.TransType = TranAppliedDeposit) Or (TUB.TransType = TranRefundDeposit) Or (TUB.TransType = TranDepPaymentVoid) Then
              For RevCnt = 1 To 15
                If TUB.RevAmt(RevCnt) > 0 Then
                  RevAmts(RevCnt) = Round#(RevAmts(RevCnt) - TUB.RevAmt(RevCnt))
                  TotalDepAmt# = Round#(TotalDepAmt# - TUB.RevAmt(RevCnt))
                End If
              Next
            End If
            LastTran& = TUB.PrevTrans
          Loop While LastTran& > 0
        End If
        DoEvents
        '_________
          ToPrint$ = Str$(AcctNumber)
          ToPrint$ = ToPrint$ + "|" + Str$(UBCustRec.DepositAmt)
          For RevCnt = 1 To 15
            ToPrint$ = ToPrint$ + "|" + Str$(RevAmts(RevCnt))
          Next
          ToPrint$ = ToPrint$ + "|"
          Print #RptHandle, ToPrint$
          ToPrint$ = ""
        End If
    End If

  Next
  
  Close UBCust, UBRpt
  Close UBTran

ExitDepositListing:

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Notes
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  Dim MsgRecord  As UBMessRecType
  Dim MsgHandle As Integer
  Dim NumOfMsgRecs As Long

  RecLen = Len(MsgRecord)
  MsgHandle = FreeFile
  Open UBData + UBMessage For Random Shared As MsgHandle Len = RecLen
  NumOfMsgRecs = LOF(MsgHandle) / RecLen
  
  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\UBMessage.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To NumOfMsgRecs
    Get MsgHandle, intCount, MsgRecord
     If MsgRecord.CustRec > 0 Then
     If Len(MsgRecord.MessLine(1).Msg) > 0 Then
        Print #RptHandle, QPTrim$(Str$(MsgRecord.CustRec)); q$; C$;
        Print #RptHandle, PipeSymbol; QPTrim$(MsgRecord.MessLine(1).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(2).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(2).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(3).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(3).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(4).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(4).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(5).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(5).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(6).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(6).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(7).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(7).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(8).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(8).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(9).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(9).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(10).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(10).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(11).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(11).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(12).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(12).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(13).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(13).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(14).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(14).Msg) & " ";
    End If
    If Len(MsgRecord.MessLine(15).Msg) > 0 Then
        Print #RptHandle, QPTrim$(MsgRecord.MessLine(15).Msg) & " " & PipeSymbol
    End If
    End If
        
    FrmShowPctComp.ShowPctComp intCount, NumOfMsgRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close MsgHandle
           
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Owners
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim OwnerRecord  As UBOwnerRecType
  Dim OwnerHandle As Integer
  Dim NumOfOwnerRecs As Long

  RecLen = Len(OwnerRecord)
  OwnerHandle = FreeFile
  Open UBData + UBOwner For Random Shared As OwnerHandle Len = RecLen
  NumOfOwnerRecs = LOF(OwnerHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\UBOwner.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To NumOfOwnerRecs
    Get OwnerHandle, intCount, OwnerRecord
    If Len(QPTrim(OwnerRecord.OwnLName)) > 0 Or Len(QPTrim(OwnerRecord.OwnFName)) > 0 Then
        Print #RptHandle, QPTrim$(Str(intCount));
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.ADDR1);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.ADDR2);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.City);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.HPHONE);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.OwnFName);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.OwnLName);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.WPHONE);
        Print #RptHandle, PipeSymbol & QPTrim$(OwnerRecord.ZipCode) & PipeSymbol
    End If
    FrmShowPctComp.ShowPctComp intCount, NumOfOwnerRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close OwnerHandle
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Work Order Transactions
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim WORecord  As WorkOrderRecType
  Dim WOHandle As Integer
  Dim NumOfWORecs As Long
  Dim NumCustRecCheck As Long
  Dim UBCTemp As NewUBCustRecType
  Dim WoText As String
  Dim WoReplies As String
  
  UBCustRecLen = Len(UBCTemp)

  CUBFile = FreeFile
  Open "UBData\UBCUST.DAT" For Random Shared As CUBFile Len = UBCustRecLen
  NumCustRecCheck = LOF(CUBFile) \ UBCustRecLen
  Close CUBFile

  RecLen = Len(WORecord)
  WOHandle = FreeFile
  Open UBData + UBWoTrans For Random Shared As WOHandle Len = RecLen
  NumOfWORecs = LOF(WOHandle) / RecLen
  
  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\WoTrans.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  Dim intThisCount As Integer
  For intCount = 1 To NumOfWORecs
    Get WOHandle, intCount, WORecord
    If WORecord.CustRec < 0 Or WORecord.CustRec > NumCustRecCheck Then
        Stop
        GoTo SkipRec
    End If
    
     TheDate$ = Num2Date(WORecord.CompleteByDate) '1
      If InStr(TheDate$, "%") > 0 Then
        TheDate$ = ""
      End If
     Print #RptHandle, TheDate$;
     TheDate$ = Num2Date(WORecord.CompletedDate) '2
      If InStr(TheDate$, "%") > 0 Then
        TheDate$ = ""
      End If
    Print #RptHandle, PipeSymbol & TheDate$;
    TheDate$ = Num2Date(WORecord.ENTRYDATE) '3
    If InStr(TheDate$, "%") > 0 Then
        TheDate$ = ""
    End If
      
    Print #RptHandle, PipeSymbol & TheDate$;
    
    For intThisCount = 1 To 6
        WoText = Replace(WORecord.OrdersText.Text(intThisCount), "|", " ")
        Print #RptHandle, PipeSymbol & WoText; '9
    Next
    For intThisCount = 1 To 6
         WoReplies = Replace(WORecord.RepliesText.Text(intThisCount), "|", " ")
         Print #RptHandle, PipeSymbol & WoReplies; '15
    Next
        
    Print #RptHandle, PipeSymbol & WORecord.CustRec; '16
    Print #RptHandle, PipeSymbol
    
    FrmShowPctComp.ShowPctComp intCount, NumOfWORecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
SkipRec:
  Next intCount
  Close WOHandle
  


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Group Codes
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim GroupCodeRecord  As GroupCodeRecType
  Dim GroupCodesHandle As Integer
  Dim NumOfGroupCodeRecs As Long

  RecLen = Len(GroupCodeRecord)
  GroupCodesHandle = FreeFile
  Open UBData + UBGrpCde For Random Shared As GroupCodesHandle Len = RecLen
  NumOfGroupCodeRecs = LOF(GroupCodesHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\GroupCodes.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To NumOfGroupCodeRecs
    Get GroupCodesHandle, intCount, GroupCodeRecord
    Print #RptHandle, QPTrim$(GroupCodeRecord.GroupCode);
    Print #RptHandle, PipeSymbol & QPTrim$(GroupCodeRecord.GroupCodeName);
    Print #RptHandle, PipeSymbol & intCount & PipeSymbol

    FrmShowPctComp.ShowPctComp intCount, NumOfGroupCodeRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close GroupCodesHandle


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Town Configuration
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim SetupRecord  As UBSetupRecType
  Dim SetupHandle As Integer
  Dim NumOfSetupRecs As Long

  RecLen = Len(SetupRecord)
  SetupHandle = FreeFile
  Open UBData + UBSetup For Random Shared As SetupHandle Len = RecLen
  NumOfSetupRecs = LOF(SetupHandle) / RecLen

  FileName = "\UBSetup.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To NumOfSetupRecs
    Get SetupHandle, intCount, SetupRecord
    Print #RptHandle, QPTrim(SetupRecord.UTILNAME); '1
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.DEFCITY); '2
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.ZipCode); '3
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.PreByBook); '4
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.LockBoxDef); '5
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.RECPDEFT); '6
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.EstRead); '7
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.BANKDFT); '8
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.UseSeq); '9
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.BILLCYCL); '10
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.DefLook); '11
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.MethAcct); '12
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.SkipInactive); '13
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.SkipSeparator); '14
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.Make99File); '15
    Print #RptHandle, PipeSymbol & SetupRecord.LowRead; '16
    Print #RptHandle, PipeSymbol & SetupRecord.HighRead; '17
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.HHDEVICE) & PipeSymbol; '18
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.DEFSTATE); '19

    FrmShowPctComp.ShowPctComp intCount, NumOfSetupRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Revenues
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FileName = "\UBServices.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To 15
    Get SetupHandle, 1, SetupRecord
    Print #RptHandle, QPTrim$(SetupRecord.Revenues(intCount).RevName);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.Revenues(intCount).Prorate);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.Revenues(intCount).UseDep);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.Revenues(intCount).USERATE);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.Revenues(intCount).UseMtr);
    Print #RptHandle, PipeSymbol & SetupRecord.Revenues(intCount).DistOr & PipeSymbol
    
    FrmShowPctComp.ShowPctComp intCount, 15
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount


  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Bill GL Accounts
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FileName = "\UBBillGLAccounts.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To 15
    Get SetupHandle, 1, SetupRecord
    Print #RptHandle, QPTrim$(SetupRecord.Revenues(intCount).RevName);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.BillAcct(intCount).DebitAcct);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.BillAcct(intCount).CreditAcct) & PipeSymbol


    FrmShowPctComp.ShowPctComp intCount, 15
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Pay GL Accounts
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FileName = "\UBPayGLAccounts.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To 15
    Get SetupHandle, 1, SetupRecord
    Print #RptHandle, QPTrim$(SetupRecord.Revenues(intCount).RevName);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.PayAcct(intCount).DebitAcct);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.PayAcct(intCount).CreditAcct) & PipeSymbol

    FrmShowPctComp.ShowPctComp intCount, 15
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Deposits GL Accounts
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  FileName = "\UBDepGLAccounts.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To 15
    Get SetupHandle, 1, SetupRecord
    Print #RptHandle, QPTrim$(SetupRecord.Revenues(intCount).RevName);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.DepAcct(intCount).DebitAcct);
    Print #RptHandle, PipeSymbol & QPTrim$(SetupRecord.DepAcct(intCount).CreditAcct) & PipeSymbol

    FrmShowPctComp.ShowPctComp intCount, 15
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close SetupHandle


  GoTo SkipDraft
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Draft Record
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim DraftRecord As UBDraftRecType
  Dim DraftHandle As Integer
  Dim NumOfDraftRecs As Long

  RecLen = Len(DraftRecord)
  DraftHandle = FreeFile
  Open UBData + UBDraftRec For Random Shared As DraftHandle Len = RecLen
  NumOfDraftRecs = LOF(DraftHandle) / RecLen

  FileName = "\UBDraftRec.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
 
  For intCount = 1 To 1
    Get DraftHandle, intCount, DraftRecord
    Print #RptHandle, QPTrim$(DraftRecord.BANKDEST);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.BANKLOC);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.BankName);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.BANKORIG);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.COMPACCT);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.FEDID);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.FEDPREFX);
    Print #RptHandle, PipeSymbol & QPTrim$(DraftRecord.FileName) & PipeSymbol

    FrmShowPctComp.ShowPctComp intCount, NumOfDraftRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close DraftHandle
SkipDraft:

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Rate Codes
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim RateCodesRecord As UBRateTblRecType
  Dim RateCodesHandle As Integer
  Dim NumOfRateCodesRecs As Long

  RecLen = Len(RateCodesRecord)
  RateCodesHandle = FreeFile
  Open UBData + UBRateCodes For Random Shared As RateCodesHandle Len = RecLen
  NumOfRateCodesRecs = LOF(RateCodesHandle) / RecLen

  FileName = "\UBRateCodes.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To NumOfRateCodesRecs
    Get RateCodesHandle, intCount, RateCodesRecord
    Print #RptHandle, QPTrim$(RateCodesRecord.ChkByte); '1
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(RateCodesRecord.MaxAmt, "#####.##", False)); '2
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(RateCodesRecord.MINAMT, "#####.##", False)); '3
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(RateCodesRecord.MINUNITS, "#####.##", False)); '4
    Print #RptHandle, PipeSymbol & QPTrim$(RateCodesRecord.Ratecode); '5
    Print #RptHandle, PipeSymbol & QPTrim$(RateCodesRecord.RATEDESC) + PipeSymbol '6

    FrmShowPctComp.ShowPctComp intCount, NumOfRateCodesRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Rate Code Breaks
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    FileName = "\UBRateCodeBreaks.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  Dim intCount02 As Integer
  Dim tempUnitAmt As Double
  Dim tempUnits As Long
  For intCount = 1 To NumOfRateCodesRecs
    Get RateCodesHandle, intCount, RateCodesRecord
    Print #RptHandle, QPTrim$(RateCodesRecord.Ratecode) + PipeSymbol;
    For intCount02 = 1 To 10
        tempUnitAmt = RateCodesRecord.TblBreaks(intCount02).UNITAMT
        tempUnits = RateCodesRecord.TblBreaks(intCount02).UNITS
        If tempUnits >= 0 Then
            If intCount02 = 1 Then
                Print #RptHandle, QPTrim$(Str(tempUnits));
            Else
                Print #RptHandle, PipeSymbol & QPTrim$(Str(tempUnits));
            End If
            Print #RptHandle, PipeSymbol & QPTrim$(Str(tempUnitAmt));
        Else
            Print #RptHandle, PipeSymbol & 0;
            Print #RptHandle, PipeSymbol & 0;
        End If
        If intCount02 = 10 Then
             Print #RptHandle, PipeSymbol
             Exit For
        End If
    Next

    FrmShowPctComp.ShowPctComp intCount, NumOfRateCodesRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close RateCodesHandle
  

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Laser Bill Setup information
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim LaserBillRecord  As UBBillLetterType
  Dim LaserBillHandle As Integer
  Dim NumberOfLaserBillRecs As Long

  RecLen = Len(LaserBillRecord)
  LaserBillHandle = FreeFile
  Open UBData + UBLaserBill For Random Shared As LaserBillHandle Len = RecLen
  NumberOfLaserBillRecs = LOF(LaserBillHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\UBLaserBill.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If

  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  Dim intPrintLogo As Integer
  For intCount = 1 To NumberOfLaserBillRecs
    Get LaserBillHandle, intCount, LaserBillRecord
    intPrintLogo = LaserBillRecord.IncLogoFlag
    If intPrintLogo = 1 Then
        Print #RptHandle, QPTrim$("True");
    Else
        Print #RptHandle, QPTrim$("False");
    End If
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.BL1Head1);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.BL1Head2);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.BL1Head3);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgOpt1);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgOpt2);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgPgph1);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgPgph2);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgPgph3);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgPgph4);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgPgph5);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgOpt3);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgOpt4);
    Print #RptHandle, PipeSymbol & QPTrim$(LaserBillRecord.MsgOpt5);
    Print #RptHandle, PipeSymbol & QPTrim$(Str(LaserBillRecord.MtrNumFlag)) + PipeSymbol

    FrmShowPctComp.ShowPctComp intCount, NumberOfLaserBillRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close LaserBillHandle

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Late notice format
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim LateNoticeRecord  As UBLateLetterType
  Dim LateNoticeHandle As Integer
  Dim NumberOfLateRecs As Long
  
  RecLen = Len(LateNoticeRecord)
  LateNoticeHandle = FreeFile
  Open UBData + UBLateLetter For Random Shared As LateNoticeHandle Len = RecLen
  NumberOfLateRecs = LOF(LateNoticeHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\UBLateLetter.txt"
  If Exist(StartPath + "\UtilityBillingData\" + FileName) Then
    KillFile (StartPath + "\UtilityBillingData\" + FileName)
  End If
 
  Dim strBody As String
  strBody = ""
  RptName$ = StartPath + "\UtilityBillingData\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  For intCount = 1 To NumberOfLateRecs
    Get LateNoticeHandle, intCount, LateNoticeRecord
    Print #RptHandle, QPTrim$(LateNoticeRecord.Head1);
    Print #RptHandle, PipeSymbol & QPTrim$(LateNoticeRecord.Head2);
    Print #RptHandle, PipeSymbol & QPTrim$(LateNoticeRecord.Head3);
    Print #RptHandle, PipeSymbol & QPTrim$(LateNoticeRecord.Head4);
    Print #RptHandle, PipeSymbol & QPTrim$(LateNoticeRecord.Head5);
    For intCount02 = 1 To 20
        strBody = QPTrim$(LateNoticeRecord.Body(intCount02))
        If Len(strBody) > 1 Then
           Print #RptHandle, QPTrim$(LateNoticeRecord.Body(intCount02)) + PipeSymbol;
           Print #RptHandle, Str(intCount02) + PipeSymbol
        End If
    Next
    FrmShowPctComp.ShowPctComp intCount, NumberOfLateRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next intCount
  Close LateNoticeHandle
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Customer information
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  q$ = ""
  C$ = "|"
  ReDim GPC(1 To 1) As GroupCodeRecType
  GPCLen = Len(GPC(1))
  GFile = FreeFile
  Open "UBData\UBGrpCde.DAT" For Random Shared As GFile Len = GPCLen
'  NumGPRec = LOF(GFile) \ GPCLen
'  ReDim Preserve GPC(1 To NumGPRec) As GroupCodeRecType
'  For Cnt& = 1 To NumGPRec
'    Get GFile, Cnt&, GPC(Cnt&)
'  Next
  Close GFile

  Dim PctComp As Double
  Dim PrevTranRec As Long
   
  Dim UBC As NewUBCustRecType
  UBCustRecLen = Len(UBC)          'Length of Cust Record Structure

  Dim WUB As WorkOrderRecType
  UBWorkRecLen = Len(WUB)

'  Dim TUB As UBTransRecType
  UBTransRecLen = Len(TUB)          'Length of Cust Record Structure

  CUBFile = FreeFile
  Open "UBData\UBCUST.DAT" For Random Shared As CUBFile Len = UBCustRecLen
  NumOfRecs& = LOF(CUBFile) \ UBCustRecLen
  
  TUBFile = FreeFile
  Open "UBData\UBTRANS.DAT" For Random Shared As TUBFile Len = UBTransRecLen
  TNumOfRecs& = 0& + LOF(TUBFile) \ UBTransRecLen

  WUBFile = FreeFile
  Open "UBData\UBWRKORD.DAT" For Random Shared As WUBFile Len = UBWorkRecLen
  WNumOfRecs& = LOF(WUBFile) \ UBWorkRecLen

  CEXFILE = FreeFile
  Open "UtilityBillingData\UBCUST.TXT" For Output As CEXFILE Len = 32767
   
  TEXFile = FreeFile
  Open "UtilityBillingData\UBTRAN.TXT" For Output As TEXFile Len = 32767

  WEXFile = FreeFile
  Open "UtilityBillingData\UBWRKO.TXT" For Output As WEXFile Len = 32767

  BEXFile = FreeFile
  Open "UtilityBillingData\UBCBAL.TXT" For Output As BEXFile Len = 32767

  Print #CEXFILE, q$; "ACCTNO"; q$; C$; q$; "BOOK"; q$; C$; q$; "SEQNUMB"; q$; C$; q$; "Status"; q$; C$; q$; "OPENDATE"; q$; C$; q$; "GPC"; q$; C$; q$; "SEARCH"; q$; C$; q$; "CUSTNAME"; q$; C$; q$; "ADDR1"; q$; C$; q$; "ADDR2"; q$; C$; q$; "SERVADDR"; q$; C$; q$; "CITY"; q$; C$; q$; "STATE"; q$; C$; q$; "ZIPCODE"; q$; C$; q$; "DPC"; q$; C$; q$; "HPHONE"; q$; C$; q$; "WPHONE"; q$; C$; q$; "SOSEC"; q$; C$; q$; "DRVLIC"; q$; C$; q$; "CUSTTYPE"; q$; C$; q$; "Addr911"; q$; C$; q$; "BillTo"; q$; C$; q$; "BILLCOPY"; q$; C$; q$; "POSTRTE"; q$; C$; q$; "BILLCYCL"; q$; C$; q$; "ZONE"; q$; C$;
  Print #CEXFILE, q$; "SEQ"; q$; C$; q$; "CASHONLY"; q$; C$; q$; "LATEFEE"; q$; C$; q$; "CUTOFFYN"; q$; C$; q$; "TAXEXPT"; q$; C$; q$; "SRCIT"; q$; C$; q$; "USEDRAFT"; q$; C$; q$; "AcctType"; q$; C$; q$; "BANKNAME"; q$; C$; q$; "BANKLOC"; q$; C$; q$; "TRANSIT"; q$; C$; q$; "BANKACCT"; q$; C$; q$; "BILLCMNT"; q$; C$; q$; "PAYCMNT"; q$; C$; q$; "PUMPCODE"; q$; C$; q$; "USERCODE1"; q$; C$; q$; "USERCODE2"; q$; C$; q$; "ProRatePCT"; q$; C$; q$; "HHMSG1"; q$; C$; q$; "HHMSG2"; q$; C$; q$; "HHMSG3"; q$; C$;
  Print #CEXFILE, q$; "RC1"; q$; C$; q$; "RC2"; q$; C$; q$; "RC3"; q$; C$; q$; "RC4"; q$; C$; q$; "RC5"; q$; C$; q$; "RC6"; q$; C$; q$; "RC7"; q$; C$; q$; "RC8"; q$; C$; q$; "RC9"; q$; C$; q$; "RC10"; q$; C$; q$; "RC11"; q$; C$; q$; "RC12"; q$; C$; q$; "RC13"; q$; C$; q$; "RC14"; q$; C$; q$; "RC15"; q$; C$; q$; "FRDESC1"; q$; C$; q$; "FRAMT1"; q$; C$; q$; "FRFREQ1"; q$; C$; q$; "REVSRC1"; q$; C$; q$; "NumMin1"; q$; C$; q$; "FRDESC2"; q$; C$; q$; "FRAMT2"; q$; C$;
  Print #CEXFILE, q$; "FRFREQ2"; q$; C$; q$; "REVSRC2"; q$; C$; q$; "NumMin2"; q$; C$; q$; "FRDESC3"; q$; C$; q$; "FRAMT3"; q$; C$; q$; "FRFREQ3"; q$; C$; q$; "REVSRC3"; q$; C$; q$; "NumMin3"; q$; C$; q$; "FRDESC4"; q$; C$; q$; "FRAMT4"; q$; C$; q$; "FRFREQ4"; q$; C$; q$; "REVSRC4"; q$; C$; q$; "NumMin4"; q$; C$; q$; "AMTOWED1"; q$; C$; q$; "TotAmtPD1"; q$; C$; q$; "PayAmt1"; q$; C$; q$; "RevSource1"; q$; C$; q$; "AMTOWED2"; q$; C$; q$; "TotAmtPD2"; q$; C$; q$; "PayAmt2"; q$; C$; q$; "RevSource2"; q$; C$;
  Print #CEXFILE, q$; "MFEE1"; q$; C$; q$; "MFEE2"; q$; C$; q$; "MTNO1"; q$; C$; q$; "MTMU1"; q$; C$; q$; "MTRType1"; q$; C$; q$; "MTRUnit1"; q$; C$; q$; "NumUser1"; q$; C$; q$; "InsDate1"; q$; C$; q$; "CurRead1"; q$; C$; q$; "PrevRead1"; q$; C$; q$; "CurDate1"; q$; C$; q$; "PastDate1"; q$; C$; q$; "ReadFlag1"; q$; C$; q$; "AvgUse1"; q$; C$; q$; "UseCnt1"; q$; C$; q$; "MID1"; q$; C$; q$; "MLT1"; q$; C$; q$; "MLG1"; q$; C$; q$; "MTNO2"; q$; C$; q$; "MTMU2"; q$; C$; q$; "MTRType2"; q$; C$; q$; "MTRUnit2"; q$; C$; q$; "NumUser2"; q$; C$; q$; "InsDate2"; q$; C$;
  Print #CEXFILE, q$; "CurRead2"; q$; C$; q$; "PrevRead2"; q$; C$; q$; "CurDate2"; q$; C$; q$; "PastDate2"; q$; C$; q$; "ReadFlag2"; q$; C$; q$; "AvgUse2"; q$; C$; q$; "UseCnt2"; q$; C$; q$; "MID2"; q$; C$; q$; "MLT2"; q$; C$; q$; "MLG2"; q$; C$; q$; "MTNO3"; q$; C$; q$; "MTMU3"; q$; C$; q$; "MTRType3"; q$; C$; q$; "MTRUnit3"; q$; C$; q$; "NumUser3"; q$; C$; q$; "InsDate3"; q$; C$; q$; "CurRead3"; q$; C$; q$; "PrevRead3"; q$; C$; q$; "CurDate3"; q$; C$; q$; "PastDate3"; q$; C$; q$; "ReadFlag3"; q$; C$; q$; "AvgUse3"; q$; C$; q$; "UseCnt3"; q$; C$; q$; "MID3"; q$; C$; q$; "MLT3"; q$; C$; q$; "MLG3"; q$; C$;
  Print #CEXFILE, q$; "MTNO4"; q$; C$; q$; "MTMU4"; q$; C$; q$; "MTRType4"; q$; C$; q$; "MTRUnit4"; q$; C$; q$; "NumUser4"; q$; C$; q$; "InsDate4"; q$; C$; q$; "CurRead4"; q$; C$; q$; "PrevRead4"; q$; C$; q$; "CurDate4"; q$; C$; q$; "PastDate4"; q$; C$; q$; "ReadFlag4"; q$; C$; q$; "AvgUse4"; q$; C$; q$; "UseCnt4"; q$; C$; q$; "MID4"; q$; C$; q$; "MLT4"; q$; C$; q$; "MLG4"; q$; C$; q$; "MTNO5"; q$; C$; q$; "MTMU5"; q$; C$; q$; "MTRType5"; q$; C$; q$; "MTRUnit5"; q$; C$; q$; "NumUser5"; q$; C$; q$; "InsDate5"; q$; C$; q$; "CurRead5"; q$; C$;
  Print #CEXFILE, q$; "PrevRead5"; q$; C$; q$; "CurDate5"; q$; C$; q$; "PastDate5"; q$; C$; q$; "ReadFlag5"; q$; C$; q$; "AvgUse5"; q$; C$; q$; "UseCnt5"; q$; C$; q$; "MID5"; q$; C$; q$; "MLT5"; q$; C$; q$; "MLG5"; q$; C$; q$; "MTNO6"; q$; C$; q$; "MTMU6"; q$; C$; q$; "MTRType6"; q$; C$; q$; "MTRUnit6"; q$; C$; q$; "NumUser6"; q$; C$; q$; "InsDate6"; q$; C$; q$; "CurRead6"; q$; C$; q$; "PrevRead6"; q$; C$; q$; "CurDate6"; q$; C$; q$; "PastDate6"; q$; C$; q$; "ReadFlag6"; q$; C$; q$; "AvgUse6"; q$; C$; q$; "UseCnt6"; q$; C$; q$; "MID6"; q$; C$; q$; "MLT6"; q$; C$; q$; "MLG6"; q$; C$; q$; "MTNO7"; q$; C$;
  Print #CEXFILE, q$; "MTMU7"; q$; C$; q$; "MTRType7"; q$; C$; q$; "MTRUnit7"; q$; C$; q$; "NumUser7"; q$; C$; q$; "InsDate7"; q$; C$; q$; "CurRead7"; q$; C$; q$; "PrevRead7"; q$; C$; q$; "CurDate7"; q$; C$; q$; "PastDate7"; q$; C$; q$; "ReadFlag7"; q$; C$; q$; "AvgUse7"; q$; C$; q$; "UseCnt7"; q$; C$; q$; "MID7"; q$; C$; q$; "MLT7"; q$; C$; q$; "MLG7"; q$   '; c$; ' q$; "CurrBalance"; q$; c$; q$; "PrevBalance"; q$; c$; ' q$; "CRA1"; q$; c$; q$; "CRA2"; q$; c$; q$; "CRA3"; q$; c$;

  Print #BEXFile, q$; "ACCTNO"; q$; C$; q$; "CurrBalance"; q$; C$; q$; "PrevBalance"; q$; C$; q$; "CRA1"; q$; C$; q$; "CRA2"; q$; C$; q$; "CRA3"; q$; C$; q$; "CRA4"; q$; C$; q$; "CRA5"; q$; C$; q$; "CRA6"; q$; C$; q$; "CRA7"; q$; C$; q$; "CRA8"; q$; C$; q$; "CRA9"; q$; C$; q$; "CRA10"; q$; C$; q$; "CRA11"; q$; C$; q$; "CRA12"; q$; C$; q$; "CRA13"; q$; C$; q$; "CRA14"; q$; C$; q$; "CRA15"; q$; C$; q$; "DepositAmt"; q$

  Print #TEXFile, q$; "ACCTNO"; q$; C$; q$; "TRDATE"; q$; C$; q$; "TRTYPE"; q$; C$; q$; "TRDESC"; q$; C$; q$; "TRAMT"; q$; C$; q$; "REV1"; q$; C$; q$; "REV2"; q$; C$; q$; "REV3"; q$; C$; q$; "REV4"; q$; C$; q$; "REV5"; q$; C$; q$; "REV6"; q$; C$; q$; "REV7"; q$; C$; q$; "REV8"; q$; C$; q$; "REV9"; q$; C$; q$; "REV10"; q$; C$; q$; "REV11"; q$; C$; q$; "REV12"; q$; C$; q$; "REV13"; q$; C$; q$; "REV14"; q$; C$; q$; "REV15"; q$; C$;
  Print #TEXFile, q$; "TAX1"; q$; C$; q$; "TAX2"; q$; C$; q$; "TAX3"; q$; C$; q$; "TAX4"; q$; C$; q$; "TAX5"; q$; C$; q$; "TAX6"; q$; C$; q$; "TAX7"; q$; C$; q$; "TAX8"; q$; C$; q$; "TAX9"; q$; C$; q$; "TAX10"; q$; C$; q$; "TAX11"; q$; C$; q$; "TAX12"; q$; C$; q$; "TAX13"; q$; C$; q$; "TAX14"; q$; C$; q$; "TAX15"; q$; C$;
  Print #TEXFile, q$; "MTY1"; q$; C$; q$; "MTY2"; q$; C$; q$; "MTY3"; q$; C$; q$; "MTY4"; q$; C$; q$; "MTY5"; q$; C$; q$; "MTY6"; q$; C$; q$; "MTY7"; q$; C$;
  Print #TEXFile, q$; "CRD1"; q$; C$; q$; "CRD2"; q$; C$; q$; "CRD3"; q$; C$; q$; "CRD4"; q$; C$; q$; "CRD5"; q$; C$; q$; "CRD6"; q$; C$; q$; "CRD7"; q$; C$;
  Print #TEXFile, q$; "PRD1"; q$; C$; q$; "PRD2"; q$; C$; q$; "PRD3"; q$; C$; q$; "PRD4"; q$; C$; q$; "PRD5"; q$; C$; q$; "PRD6"; q$; C$; q$; "PRD7"; q$; C$;
  Print #TEXFile, q$; "ERD1"; q$; C$; q$; "ERD2"; q$; C$; q$; "ERD3"; q$; C$; q$; "ERD4"; q$; C$; q$; "ERD5"; q$; C$; q$; "ERD6"; q$; C$; q$; "ERD7"; q$; C$;
  Print #TEXFile, q$; "BLNO"; q$; C$; q$; "RDATE"; q$; C$; q$; "BDATE"; q$; C$; q$; "PDDATE"; q$; C$; q$; "DDATE"; q$; C$; q$; "PRP"; q$; C$; q$; "OPNO"; q$; C$; q$; "RBAL"; q$; C$; q$; "CHKA"; q$; C$; q$; "CSHA"; q$; C$; q$; "PDATE"; q$

  Print #WEXFile, q$; "ACCTNO"; q$; C$; q$; "EDATE"; q$; C$; q$; "OT1"; q$; C$; q$; "OT2"; q$; C$; q$; "OT3"; q$; C$; q$; "OT4"; q$; C$; q$; "OT5"; q$; C$; q$; "OT6"; q$; C$; q$; "RT1"; q$; C$; q$; "RT2"; q$; C$; q$; "RT3"; q$; C$; q$; "RT4"; q$; C$; q$; "RT5"; q$; C$; q$; "RT6"; q$; C$;
  Print #WEXFile, q$; "CBDATE"; q$; C$; q$; "CDATE"; q$
  'GoTo NowWorkOrders

  FrmShowPctComp.Label1 = "Utility Bililng Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  
  Dim strBook As String
  
  For Cnt& = 1 To NumOfRecs&
    Get CUBFile, Cnt&, UBC
    If UBC.DelFlag = False Then
      If Len(QPTrim(UBC.BOOK)) > 0 Then
        strBook = "0"
        strBook = strBook + UBC.BOOK
      Else
        strBook = ""
      End If
      
      Print #CEXFILE, q$; QPTrim$(Str$(Cnt&)); q$; C$; '1
      Print #CEXFILE, q$; strBook; q$; C$; q$; QPTrim$(UBC.SEQNUMB); q$; C$; q$; QPTrim$(UBC.Status); q$; C$;
      TheDate$ = Num2Date(UBC.OPENDATE)
      If InStr(TheDate$, "%") > 0 Then
        TheDate$ = ""
      End If

'      If UBC.GroupCodeRec <= NumGPRec Then
'        If UBC.GroupCodeRec > 0 Then
'          GPCE$ = QPTrim$(GPC(UBC.GroupCodeRec).GroupCode)
'        Else
          GPCE$ = ""
'        End If
'      End If

      Print #CEXFILE, q$; TheDate$; q$; C$; q$; GPCE$; q$; C$; q$; QPTrim$(UBC.SEARCH); q$; C$; q$; QPTrim$(UBC.CustName); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.ADDR1); q$; C$; q$; QPTrim$(UBC.ADDR2); q$; C$; q$; QPTrim$(UBC.SERVADDR); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.City); q$; C$; q$; QPTrim$(UBC.State); q$; C$; q$; QPTrim$(UBC.ZipCode); q$; C$; q$; ""; q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.HPHONE); q$; C$; q$; QPTrim$(UBC.WPHONE); q$; C$; q$; QPTrim$(UBC.SOSEC); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.DRVLIC); q$; C$; q$; QPTrim$(UBC.CUSTTYPE); q$; C$; q$; QPTrim$(UBC.Addr911); q$; C$;
      If UBC.BILLCOPY < 1 Then
        UBC.BILLCOPY = 1
      End If
      Print #CEXFILE, q$; QPTrim$(UBC.BillTo); q$; C$; q$; QPTrim$(Str$(UBC.BILLCOPY)); q$; C$; q$; QPTrim$(UBC.POSTRTE); q$; C$;
      If UBC.BILLCYCL < 0 Then
        UBC.BILLCYCL = 0
      End If
      If UBC.SEQ < 0 Then
        UBC.SEQ = 0
      End If
      Print #CEXFILE, q$; QPTrim$(Str$(UBC.BILLCYCL)); q$; C$; q$; QPTrim$(UBC.ZONE); q$; C$; q$; QPTrim$(Str$(UBC.SEQ)); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.CASHONLY); q$; C$; q$; QPTrim$(UBC.LATEFEE); q$; C$; q$; QPTrim$(UBC.CUTOFFYN); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.TAXEXPT); q$; C$; q$; QPTrim$(UBC.SRCIT); q$; C$;
      'Print #CEXFILE, q$; QPTrim$(UBC.USEDRAFT); q$; C$; q$; QPTrim$(UBC.AcctType); q$; C$; q$; QPTrim$(UBC.BANKNAME); q$; C$;
      'EPP
      If QPTrim$(UBC.USEDRAFT) = "Y" Then
        Print #CEXFILE, q$; QPTrim$(UBC.USEDRAFT); q$; C$; q$; "C"; q$; C$; q$; QPTrim$(UBC.BankName); q$; C$;
      Else
        Print #CEXFILE, q$; QPTrim$(UBC.USEDRAFT); q$; C$; q$; ""; q$; C$; q$; QPTrim$(UBC.BankName); q$; C$;
      End If
      Print #CEXFILE, q$; QPTrim$(UBC.BANKLOC); q$; C$; q$; QPTrim$(UBC.TRANSIT); q$; C$; q$; QPTrim$(UBC.BankAcct); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.BILLCMNT); q$; C$; q$; QPTrim$(UBC.PAYCMNT); q$; C$; q$; QPTrim$(UBC.PUMPCODE); q$; C$;
      If UBC.ProratePCT < 0 Then
        UBC.ProratePCT = 100
      End If
      Print #CEXFILE, q$; QPTrim$(UBC.USERCODE1); q$; C$; q$; QPTrim$(UBC.USERCODE2); q$; C$; q$; QPTrim$(Str$(UBC.ProratePCT)); q$; C$;
      Print #CEXFILE, q$; QPTrim$(UBC.HHMSG1); q$; C$; q$; QPTrim$(UBC.HHMSG2); q$; C$; q$; QPTrim$(UBC.HHMSG3); q$; C$;
'Service's
      For SCnt = 1 To 15
        Print #CEXFILE, q$; QPTrim$(UBC.Serv(SCnt).Ratecode); q$; C$; q$; QPTrim$(UBC.Serv(SCnt).RMtrType); q$; C$;
      Next
'Flat Rates
      For FCnt = 1 To 4
        If UBC.FlatRates(FCnt).FRAMT < 0 Then
          UBC.FlatRates(FCnt).FRAMT = 0
        End If
        If UBC.FlatRates(FCnt).REVSRC < 0 Then
          UBC.FlatRates(FCnt).REVSRC = 0
        End If
        Print #CEXFILE, q$; QPTrim$(UBC.FlatRates(FCnt).FRDESC);
        Print #CEXFILE, q$; C$; q$; QPTrim$(UBUsing$(Str$(UBC.FlatRates(FCnt).FRAMT), "#####.##")); q$; C$;
        Print #CEXFILE, q$; QPTrim$(UBC.FlatRates(FCnt).FRFREQ); q$; C$; q$; QPTrim$(Str$(UBC.FlatRates(FCnt).REVSRC)); q$; C$;
        If UBC.FlatRates(FCnt).NumMin < 0 Then
          UBC.FlatRates(FCnt).NumMin = 0
        End If
        Print #CEXFILE, q$; QPTrim$(Str$(UBC.FlatRates(FCnt).NumMin)); q$; C$;
      Next
      
'Monthly
      For MCnt = 1 To 2
        If UBC.Monthly(MCnt).AMTOWED < 0 Then
          UBC.Monthly(MCnt).AMTOWED = 0
        End If
        Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(UBC.Monthly(MCnt).AMTOWED), "#####.##")); q$; C$; q$; QPTrim$(UBUsing$(Str$(UBC.Monthly(MCnt).TotAmtPD), "#####.##")); q$; C$;
        If UBC.Monthly(MCnt).PayAmt < 0 Then
          UBC.Monthly(MCnt).PayAmt = 0
        End If
        If UBC.Monthly(MCnt).RevSource < 0 Then
          UBC.Monthly(MCnt).RevSource = 0
        End If
        Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(UBC.Monthly(MCnt).PayAmt), "#####.##")); q$; C$; q$; QPTrim$(Str$(UBC.Monthly(MCnt).RevSource)); q$; C$;
      Next
      If UBC.MFEE1 < 0 Then UBC.MFEE1 = 0
      If UBC.MFEE2 < 0 Then UBC.MFEE2 = 0
      Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(UBC.MFEE1), "#####.##")); q$; C$; q$; QPTrim$(UBUsing$(Str$(UBC.MFEE2), "#####.##")); q$; C$;
'Location Meters
      For lcnt = 1 To 7
        If UBC.LocMeters(lcnt).MTRMulti < 0 Then
          UBC.LocMeters(lcnt).MTRMulti = 0
        End If
        Print #CEXFILE, q$; QPTrim$(UBC.LocMeters(lcnt).MTRNUM); q$; C$; q$; QPTrim$(Str$(UBC.LocMeters(lcnt).MTRMulti)); q$; C$;
        Print #CEXFILE, q$; QPTrim$(UBC.LocMeters(lcnt).MTRType); q$; C$; q$; QPTrim$(UBC.LocMeters(lcnt).MTRUnit); q$; C$;
        TheDate$ = Num2Date(UBC.LocMeters(lcnt).InsDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        If UBC.LocMeters(lcnt).NumUser < 0 Then
          UBC.LocMeters(lcnt).NumUser = 0
        End If
        Print #CEXFILE, q$; QPTrim$(Str$(UBC.LocMeters(lcnt).NumUser)); q$; C$; q$; TheDate$; q$; C$;
        If UBC.LocMeters(lcnt).CurRead < 0 Then
          UBC.LocMeters(lcnt).CurRead = 0
        End If
        If UBC.LocMeters(lcnt).PrevRead < 0 Then
          UBC.LocMeters(lcnt).PrevRead = 0
        End If
        Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(UBC.LocMeters(lcnt).CurRead), "#########")); q$; C$; q$; QPTrim$(UBUsing$(Str$(UBC.LocMeters(lcnt).PrevRead), "#########")); q$; C$;
        TheDate$ = Num2Date(UBC.LocMeters(lcnt).CurDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #CEXFILE, q$; TheDate$; q$; C$; q$;
        TheDate$ = Num2Date(UBC.LocMeters(lcnt).PastDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #CEXFILE, TheDate$; q$; C$;
        If UBC.LocMeters(lcnt).AvgUse < 0 Then UBC.LocMeters(lcnt).AvgUse = 0
        Print #CEXFILE, q$; QPTrim$(UBC.LocMeters(lcnt).ReadFlag); q$; C$; q$; QPTrim$(UBUsing$(Str$(UBC.LocMeters(lcnt).AvgUse), "#########")); q$; C$;
        If UBC.LocMeters(lcnt).UseCnt < 0 Then UBC.LocMeters(lcnt).UseCnt = 0
            Print #CEXFILE, q$; QPTrim$(Str$(UBC.LocMeters(lcnt).UseCnt)); q$; C$;
            'EPP
            'Print #CEXFILE, q$; QPTrim$(UBC.LocMeters(lcnt).MtrIDNO); q$; C$;
            Print #CEXFILE, q$; ""; q$; C$;
            'EPP
            Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(0), "##.######")); q$; C$;
            Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(0), "##.######")); q$; C$;
            'Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(UBC.LocMeters(lcnt).MtrLat), "##.######")); q$; C$;
            'Print #CEXFILE, q$; QPTrim$(UBUsing$(Str$(UBC.LocMeters(lcnt).MtrLng), "##.######")); q$; C$;
        Next
      Print #CEXFILE, q$; QPTrim$(Str$(UBC.PreNoteFlag)); C$ 'Last record on the customer text file

      If UBC.CurrBalance < -100000 Then
        UBC.CurrBalance = 0
      End If
      If UBC.PrevBalance < -100000 Then
        UBC.PrevBalance = 0
      End If
      'If UBC.EPPFlag <> "Y" Then
        Print #BEXFile, q$; QPTrim$(Str$(Cnt&)); q$; C$;
        Print #BEXFile, q$; QPTrim$(UBUsing$(Str$(UBC.CurrBalance), "#####.##")); q$; C$; q$; QPTrim$(UBUsing$(Str$(UBC.PrevBalance), "#####.##")); q$; C$;
        For RCnt = 1 To 15
          Print #BEXFile, q$; QPTrim$(UBUsing$(Str$(UBC.CurrRevAmts(RCnt)), "#####.##")); q$; C$;
        Next
        Print #BEXFile, q$; QPTrim$(UBUsing$(Str$(UBC.DepositAmt), "#####.##")); q$; C$
      'End If
      'DO there transactions
      PrevTranRec& = UBC.LastTrans
      If PrevTranRec& > 0 Then
        Do While PrevTranRec& > 0
        Get TUBFile, PrevTranRec&, TUB
        Print #TEXFile, q$; QPTrim$(Str$(Cnt&)); q$; C$;
        TheDate$ = Num2Date(TUB.TransDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #TEXFile, q$; TheDate$; q$; C$; q$; QPTrim$(Str$(TUB.TransType)); q$; C$; q$; QPTrim$(TUB.TransDesc); q$; C$; q$; QPTrim$(UBUsing$(Str$(TUB.TransAmt), "######.##")); q$; C$;
        For RCnt = 1 To 15
          Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.RevAmt(RCnt)), "#####.##")); q$; C$;
        Next
        For RCnt = 1 To 15
          Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.TaxAmt(RCnt)), "#####.##")); q$; C$;
        Next
        For RCnt = 1 To 7
          Print #TEXFile, q$; QPTrim$(Str$(TUB.MtrTypes(RCnt))); q$; C$;
        Next
        For RCnt = 1 To 7
          Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.CurRead(RCnt)), "#########")); q$; C$;
        Next
        For RCnt = 1 To 7
          Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.PrevRead(RCnt)), "#########")); q$; C$;
        Next
        For RCnt = 1 To 7
          Print #TEXFile, q$; QPTrim$(TUB.EstRead(RCnt)); q$; C$;
        Next
        TheDate$ = Num2Date(TUB.ReadDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #TEXFile, q$; QPTrim$(Str$(TUB.BillNumber)); q$; C$; q$; TheDate$; q$; C$;
        TheDate$ = Num2Date(TUB.BillDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #TEXFile, q$; TheDate$; q$; C$;
        TheDate$ = Num2Date(TUB.PastDueDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #TEXFile, q$; TheDate$; q$; C$;
        TheDate$ = Num2Date(TUB.DraftDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        Print #TEXFile, q$; TheDate$; q$; C$;
        Print #TEXFile, q$; QPTrim$(Str$(TUB.PayTypeCode)); q$; C$; q$; QPTrim$(Str$(TUB.OperatorNumber)); q$; C$;
        Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.RunBalance), "######.##")); q$; C$;
        Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.CheckAmount), "######.##")); q$; C$;
        Print #TEXFile, q$; QPTrim$(UBUsing$(Str$(TUB.CashAmount), "######.##")); q$; C$;
        TheDate$ = Num2Date(TUB.PrevDate)
        If InStr(TheDate$, "%") > 0 Then
          TheDate$ = ""
        End If
        
        Print #TEXFile, q$; TheDate$; C$;
        For RCnt = 1 To 7
          Print #TEXFile, q$; QPTrim$(Str$(UBC.LocMeters(RCnt).MTRMulti)); q$; C$;
        Next
        Print #TEXFile,
        PrevTranRec& = TUB.PrevTrans
        DoEvents
        Loop
      End If
    End If

    FrmShowPctComp.ShowPctComp Cnt&, NumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp

      Exit Sub
    End If
  Next

End Sub


