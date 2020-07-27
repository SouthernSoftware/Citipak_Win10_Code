Attribute VB_Name = "modBLExportData"
DefInt A-Z

Dim FlipFlag As Boolean
Dim BtnFnt As Double
Dim intCount As Long
Dim RptHandle As Integer
Dim RptName As String
Dim FileName As String
Dim PipeSymbol As String
Dim RecLen As Integer
Public Sub ProcessBLData()
  StartPath = App.Path
  PipeSymbol = "|"
  
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Category Codes
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''7''''''''''''''
  Dim CatCodeRecord  As ARNewCatCodeRecType
  Dim CatCodeHandle As Integer
  Dim NumOfCatCodeRecs As Long

  RecLen = Len(CatCodeRecord)
  CatCodeHandle = FreeFile
  Open BLData + BLCatCodeName For Random Shared As CatCodeHandle Len = RecLen
  NumOfCatCodeRecs = LOF(CatCodeHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\CatCode.txt"
  If Exist(StartPath + "\Business License Data\" + FileName) Then
    KillFile (StartPath + "\Business License Data\" + FileName)
  End If

  RptName$ = StartPath + "\Business License Data\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Business License Data Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  Dim Revenue As String
  Dim CashRec As String
  Dim ARRec As String

  For intCount = 1 To NumOfCatCodeRecs
    Get CatCodeHandle, intCount, CatCodeRecord
    Print #RptHandle, QPTrim$(CatCodeRecord.CatCode);                                         '1
    Print #RptHandle, PipeSymbol & QPTrim$(CatCodeRecord.CODEDESC);                           '2
    Print #RptHandle, PipeSymbol & QPTrim$(CatCodeRecord.CodeType);                           '3
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Fee, "#####.##"));           '4
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.BaseAmt1, "#####.##"));      '5
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Recpt1, "#####.##"));        '6
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Percent1, "#####.##"));      '7
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Maximum1, "#####.##"));      '8
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.BaseAmt2, "#####.##"));      '9
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Recpt2, "#####.##"));        '10
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Percent2, "#####.##"));      '11
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Maximum2, "#####.##"));      '12
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.BaseAmt3, "#####.##"));      '13
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Recpt3, "#####.##"));        '14
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Percent3, "#####.##"));      '15
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Maximum3, "#####.##"));      '16
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.BaseAmt4, "#####.##"));      '17
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Recpt4, "#####.##"));        '18
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Percent4, "#####.##"));      '19
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Maximum4, "#####.##"));      '20
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.BaseAmt5, "#####.##"));      '21
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Recpt5, "#####.##"));        '22
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Percent5, "#####.##"));      '23
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Maximum5, "#####.##"));      '24
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.BaseAmt6, "#####.##"));      '25
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Recpt6, "#####.##"));        '26
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Percent6, "#####.##"));      '27
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(CatCodeRecord.Maximum6, "#####.##")); '28
    Revenue = GetAcctNum(CInt(CatCodeRecord.REVGLNUM))
    CashRec = GetAcctNum(CInt(CatCodeRecord.CASHACCT))
    ARRec = GetAcctNum(CInt(CatCodeRecord.ARGLACCT))
    Print #RptHandle, PipeSymbol & Revenue;      '29
    Print #RptHandle, PipeSymbol & CashRec;      '30
    Print #RptHandle, PipeSymbol & ARRec & PipeSymbol           '31
  
  
  FrmShowPctComp.ShowPctComp intCount, NumOfCatCodeRecs
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    Unload FrmShowPctComp
    Exit Sub
  End If
  Next intCount
  Close CatCodeHandle
  
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Town Config information
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim TownRecord  As TownSetUpType
  Dim TownHandle As Integer
  Dim NumOfTownRecs As Long

  RecLen = Len(TownRecord)
  TownHandle = FreeFile
  Open BLData + BLTownSetUpName For Random Shared As TownHandle Len = RecLen
  NumOfTownRecs = LOF(TownHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\BLTownInfo.txt"
  If Exist(StartPath + "\Business License Data\" + FileName) Then
    KillFile (StartPath + "\Business License Data\" + FileName)
  End If

  RptName$ = StartPath + "\Business License Data\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Business License Data Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents

  For intCount = 1 To NumOfTownRecs
    Get TownHandle, intCount, TownRecord
    Print #RptHandle, QPTrim$(TownRecord.TownName); '1
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.AcctMeth); '2
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.TownAdd1); '3
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.TownAdd2); '4
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.State); '5
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.City); '6
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.ZipCode); '7
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.TownPhone); '8
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.UseAmtPctYN); '9
    Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(TownRecord.IssFee, "#####.##", False)); '10
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.LicNumPermYN); '11
    Print #RptHandle, PipeSymbol & QPTrim$(TownRecord.GL2Cats), ' 12
    Revenue = GetAcctNum(CInt(TownRecord.PENREVGLNUM))
    CashRec = GetAcctNum(CInt(TownRecord.PENCASHACCT))
    ARRec = GetAcctNum(CInt(TownRecord.PENRECGLNUM))
    Print #RptHandle, PipeSymbol & Revenue; '13
    Print #RptHandle, PipeSymbol & CashRec; '14
    Print #RptHandle, PipeSymbol & ARRec & PipeSymbol '15
  
  
  FrmShowPctComp.ShowPctComp intCount, NumOfTownRecs
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    Unload FrmShowPctComp
    Exit Sub
  End If
  Next intCount
  Close TownHandle
  
  
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Export Accounts
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Transaction Stuff
  Dim TransRecord  As ARTransRecType
  Dim TransHandle As Integer
  Dim NumOfTransRecs As Long
  Dim FileNameTrans As String
  Dim RptNameTrans As String
  Dim RptHandleTrans As Integer
  RecLen = Len(TransRecord)
  TransHandle = FreeFile
  Open BLData + BLTransFileName For Random Shared As TransHandle Len = RecLen
  NumOfTransRecs = LOF(TransHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileNameTrans = "\BLTransactions.txt"
  If Exist(StartPath + "\Business License Data\" + FileNameTrans) Then
    KillFile (StartPath + "\Business License Data\" + FileNameTrans)
  End If

  RptNameTrans = StartPath + "\Business License Data\" + FileNameTrans
  RptHandleTrans = FreeFile
  Open RptNameTrans$ For Output As #RptHandleTrans
  
  'Account Stuff
  Dim AccountsRecord  As ARCustRecType
  Dim CustAccountHandle As Integer
  Dim NumOfAccounstRecs As Long

  RecLen = Len(AccountsRecord)
  CustAccountHandle = FreeFile
  Open BLData + BLCustFileName For Random Shared As CustAccountHandle Len = RecLen
  NumOfAccountsRecs = LOF(CustAccountHandle) / RecLen

  StartPath = App.Path

  PipeSymbol = "|"

  FileName = "\BLAccounts.txt"
  If Exist(StartPath + "\Business License Data\" + FileName) Then
    KillFile (StartPath + "\Business License Data\" + FileName)
  End If

  RptName$ = StartPath + "\Business License Data\" + FileName
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Business License Data Export"
  FrmShowPctComp.Show , frmCitiPakExportData
  DoEvents
  
  Dim intFirstTrans As Long
  For intCount = 1 To NumOfAccountsRecs
    Get CustAccountHandle, intCount, AccountsRecord
          
        Print #RptHandle, QPTrim$(AccountsRecord.CustNumb);                                       '1
        If QPTrim(AccountsRecord.SortName) = "DELETED" Or AccountsRecord.Deleted = "Y" Then
        Print #RptHandle, PipeSymbol & "Deleted";                          '2
        Else
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.BillName);                          '2
        End If
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.CustName);                          '3
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.ServAdd);                           '4
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.ADDRESS1);                          '5
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.ADDRESS2);                          '6
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.City);                              '7
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.State);                             '8
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.ZipCode);                           '9
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.CustLocation);  'Inside outside     '10
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.LICENSE);                           '11
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.Contact);                           '12
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.Prorate, "#####.##"));      '13
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.WPHONE);                            '14
        TheDate$ = Num2Date(AccountsRecord.VALID)
            If InStr(TheDate$, "%") > 0 Then
              TheDate$ = ""
            End If
        Print #RptHandle, PipeSymbol & TheDate$;                                                  '15
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.Inactive);                          '16
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.REV1, "#######"));          '17
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.REV2, "#######"));          '18
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.REV3, "#######"));          '19
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.REV4, "#######"));          '20
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.REV5, "#######"));          '21
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.BILLCAT1);                          '22
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.BILLCAT2);                          '23
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.BILLCAT3);                          '24
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.BILLCAT4);                          '25
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.BILLCAT5);                          '26
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.Fee1, "#####.##"));         '27
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.Fee2, "#####.##"));         '28
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.Fee3, "#####.##"));         '29
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.Fee4, "#####.##"));         '30
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.Fee5, "#####.##"));         '31
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.IssuanceBal, "#####.##"));  '32
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.PenBal, "#####.##"));       '33
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.AcctBal, "#####.##"));      '34
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.FeeLicBal1, "#####.##"));   '35
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.FeeLicBal2, "#####.##"));   '36
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.FeeLicBal3, "#####.##"));   '37
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.FeeLicBal4, "#####.##"));   '38
        Print #RptHandle, PipeSymbol & QPTrim$(UBUsing(AccountsRecord.FeeLicBal5, "#####.##"));   '39
        Print #RptHandle, PipeSymbol & QPTrim$(AccountsRecord.SSNFID) & PipeSymbol                '41
         
        If QPTrim(AccountsRecord.SortName) <> "DELETED" And AccountsRecord.Deleted <> "Y" Then
             'Get the transactions for this account
            intFirstTrans = AccountsRecord.FirstTrans
            Do While intFirstTrans > 0
            Get TransHandle, intFirstTrans, TransRecord
                Print #RptHandleTrans, QPTrim$(TransRecord.CustomerNumber);               '1
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.TransAmount, "#####.##"));       '2
                TheDate$ = Num2Date(TransRecord.TransDate)
                If InStr(TheDate$, "%") > 0 Then
                  TheDate$ = ""
                End If
                Print #RptHandleTrans, PipeSymbol & TheDate$;                             '3
                Print #RptHandleTrans, PipeSymbol & QPTrim$(TransRecord.TransDesc);       '4
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CashAmount, "#######"));       '5
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.ChkAmount, "#######"));        '6
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatLicAmt1, "#####.##"));       '7
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatLicAmt2, "#####.##"));       '8
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatLicAmt3, "#####.##"));        '9
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatLicAmt4, "#####.##"));        '10
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatLicAmt5, "#####.##"));        '11
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatCodeRec1, "#######"));        '12
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatCodeRec2, "#######"));        '13
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatCodeRec3, "#######"));        '14
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatCodeRec4, "#######"));        '15
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.CatCodeRec5, "#######"));        '16
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.IssAmt, "#######"));        '17
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.PenAmt, "#######"));         '18
                Print #RptHandleTrans, PipeSymbol & QPTrim$(UBUsing(TransRecord.TransType, "#####")) & PipeSymbol  '19
                intFirstTrans = TransRecord.NextTrans
            Loop
        End If
    FrmShowPctComp.ShowPctComp intCount, NumOfAccountsRecs
   If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    Unload FrmShowPctComp
    Exit Sub
  End If
  Next intCount
  Close CustAccountHandle
End Sub
 Public Function GetAcctNum$(RecordNumber)
  Dim AcctFileNum As Integer, NumAccts As Integer
    Dim GLAcct As GLAcctRecType

   OpenAcctFile AcctFileNum, NumAccts
   If RecordNumber > 0 Then
     Get AcctFileNum, RecordNumber, GLAcct
     If GLAcct.Deleted = 0 Then
       GetAcctNum$ = GLAcct.Num
     Else
       GetAcctNum$ = "Invalid Acct"
     End If
   Else
     GetAcctNum$ = "Invalid Acct"
   End If
   Close AcctFileNum

End Function

