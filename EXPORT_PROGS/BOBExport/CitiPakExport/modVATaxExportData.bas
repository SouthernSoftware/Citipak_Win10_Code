Attribute VB_Name = "modVATaxExportData"
Option Explicit

Public Sub ProcessVATaxCust()
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  OpenTaxCustFile THandle, NumOfTRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxCustData.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Customer Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfTRecs
    Get THandle, x, TaxCust
    If TaxCust.Deleted = -1 Then GoTo Skip
'    Print #RptHandle, CStr(x);
    Print #RptHandle, CStr(TaxCust.Acct);
    Print #RptHandle, B & MakeRegDate(TaxCust.OPENDATE);
    Print #RptHandle, B & CheckForBad(TaxCust.CustName);
    Print #RptHandle, B & CheckForBad(TaxCust.SName);
    Print #RptHandle, B & MakePhoneNums13Length(CheckForBad(TaxCust.HPHONE));
    Print #RptHandle, B & MakePhoneNums13Length(CheckForBad(TaxCust.WPHONE));
    Print #RptHandle, B & CheckForBad(TaxCust.CSSN);
    Print #RptHandle, B & CheckForBad(TaxCust.OSSN);
    Print #RptHandle, B & CheckForBad(TaxCust.Addr1);
    Print #RptHandle, B & CheckForBad(TaxCust.Addr2);
    Print #RptHandle, B & CheckForBad(TaxCust.City);
    Print #RptHandle, B & CheckForBad(TaxCust.State);
    Print #RptHandle, B & CheckForBad(TaxCust.Zip);
    If TaxCust.Active = "Y" Then
      Print #RptHandle, B & "N";
    Else
      Print #RptHandle, B & "Y";
    End If
'    Print #RptHandle, B & TaxCust.Active;
    Print #RptHandle, B & TaxCust.Interest;
    Print #RptHandle, B & TaxCust.TaxExempt;
    Print #RptHandle, B & TaxCust.Penalty;
    Print #RptHandle, B & CheckForBad(TaxCust.Employer);
    Print #RptHandle, B & TaxCust.Bankrupt;
    Print #RptHandle, B & CheckForBad(TaxCust.TownShip);
    Print #RptHandle, B & TaxCust.LateNotice;
    Print #RptHandle, B & CStr(TaxCust.PrePayBal);
    Print #RptHandle, B & CStr(TaxCust.PrePayTrans);
'    If CheckForBad(TaxCust.CountyAcctString) <> "" Then Stop
    Print #RptHandle, B & CheckForBad(TaxCust.CountyAcctString);
'    If CStr(TaxCust.CountyAcct) <> "0" Then Stop
    Print #RptHandle, B & CStr(TaxCust.CountyAcct);
'    Print #RptHandle, B & CStr(TaxCust.LastTrans);
'    Print #RptHandle, B & CStr(TaxCust.FirstPropRec);
'    Print #RptHandle, B & CStr(TaxCust.FirstPersRec);
'    Print #RptHandle, B & CStr(TaxCust.PIN);
'    Print #RptHandle, B & CStr(TaxCust.Deleted);
'    Print #RptHandle, B & CStr(TaxCust.FileVer);
    Print #RptHandle, B & CheckForBad(TaxCust.OptSrchDesc);
    Print #RptHandle, B & CheckForBad(TaxCust.ServiceAdd);
    Print #RptHandle, B & CheckForBad(TaxCust.DrvrsLic);
    Print #RptHandle, B & CheckForBad(TaxCust.DeliveryPt);
    Print #RptHandle, B & CheckForBad(TaxCust.PostalRt);
'    Print #RptHandle, B & CStr(TaxCust.Cycle);
    Print #RptHandle, B & CheckForBad(TaxCust.CycleName);
'    Print #RptHandle, B & CStr(TaxCust.County4BillNum);
    Print #RptHandle, B & CheckForBad(TaxCust.County4BillName);
    Print #RptHandle, B
Skip:
    FrmShowPctComp.ShowPctComp x, NumOfTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
   Next x
   Unload FrmShowPctComp
   Close
   
End Sub

Private Function MakePhoneNums13Length(ByRef PHONE As String) As String
  Dim x As Integer, Y As Integer
  Dim ch As String * 1
  Dim thisLen As Integer
  Dim NewPhone As String
  thisLen = Len(PHONE)
  MakePhoneNums13Length = "(000)000-0000"
  If thisLen = 10 Then
   NewPhone = "("
   For x = 1 To 3
     NewPhone = NewPhone + Mid(PHONE, x, 1)
   Next x
   NewPhone = NewPhone + ")"
   For x = 4 To 6
     NewPhone = NewPhone + Mid(PHONE, x, 1)
   Next x
   NewPhone = NewPhone + "-"
   For x = 7 To 10
     NewPhone = NewPhone + Mid(PHONE, x, 1)
   Next x
   MakePhoneNums13Length = NewPhone
  ElseIf thisLen = 7 Then
    NewPhone = "(000)"
    For x = 1 To 3
      NewPhone = NewPhone + Mid(PHONE, x, 1)
    Next x
    NewPhone = NewPhone + "-"
    For x = 4 To 7
      NewPhone = NewPhone + Mid(PHONE, x, 1)
    Next x
    MakePhoneNums13Length = NewPhone
  End If
End Function
'Public Sub ProcessVAMortgageCodes()
'  Dim MortRec As MortCodeRecType
'  Dim MHandle As Integer
'  Dim NumOfMRecs As Integer
'  Dim x As Integer
'  Dim RptHandle As Integer
'  Dim RptName As String
'  Dim B As String
'  Dim ThisFile As String
'
'  OpenMortCodeFile MHandle, NumOfMRecs
'  StartPath = App.Path
'  B = "|"
'  ThisFile = "\VATaxMortCodes.txt"
'  If DirExists(StartPath + "\VATAXConvToTxt") Then
'    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
'      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
'    End If
'  Else
'    MkDir StartPath + "\VATAXConvToTxt"
'  End If
'
'  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
'  RptHandle = FreeFile
'  Open RptName$ For Output As #RptHandle
'
'  FrmShowPctComp.Label1 = "Tax Mortgage Codes Export"
'  FrmShowPctComp.Show
'  DoEvents
'
'  For x = 1 To NumOfMRecs
'    Get MHandle, x, MortRec
'    If MortRec.Deleted = 1 Then GoTo Skip
'    Print #RptHandle, QPTrim$(MortRec.MORTCODE);
'    Print #RptHandle, B & QPTrim$(MortRec.BName);
'    Print #RptHandle, B & QPTrim$(MortRec.Add1);
'    Print #RptHandle, B & QPTrim$(MortRec.Add2);
'    Print #RptHandle, B & QPTrim$(MortRec.Add3);
'    Print #RptHandle, B & QPTrim$(MortRec.Contact);
'    Print #RptHandle, B & QPTrim$(MortRec.PHONE);
'    Print #RptHandle, B & QPTrim$(MortRec.XFileNme);
'    Print #RptHandle, B
'Skip:
'    FrmShowPctComp.ShowPctComp x, NumOfMRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'   Next x
'   Unload FrmShowPctComp
'   Close
'
'End Sub

Public Sub ProcessVATaxReal()
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim NextRec As Integer
  Dim dodo As Integer
  
  M = "###,###,###.##"
  N = "#############"
  OpenTaxCustFile CustHandle, NumOfCustRecs
  OpenRealPropFile RHandle, NumOfRRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxRealProp.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Real Property Export"
  FrmShowPctComp.Show
  DoEvents
  
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    NextRec = CustRec.FirstPropRec
    Do While NextRec > 0
      Get RHandle, NextRec, RealRec
      Print #RptHandle, CStr(NextRec);
      Print #RptHandle, B & QPTrim$(RealRec.RealPin);
      Print #RptHandle, B & MakeRegDate(RealRec.PROPDATE);
      Print #RptHandle, B & QPTrim$(RealRec.GISPOS);
      Print #RptHandle, B & QPTrim$(RealRec.Map);
      Print #RptHandle, B & QPTrim$(RealRec.BLOCK);
      Print #RptHandle, B & QPTrim$(RealRec.LOTNUMB);
      Print #RptHandle, B & RealRec.LOTACRE;
      Print #RptHandle, B & CStr(RealRec.PropSize);
      Print #RptHandle, B & RealRec.PROPDISC;
      Print #RptHandle, B & RealRec.LateList;
      Print #RptHandle, B & CStr(RealRec.OptRev1Chrg);
      Print #RptHandle, B & CStr(RealRec.OptRev2Chrg);
      Print #RptHandle, B & CStr(RealRec.OptRev3Chrg);
      Print #RptHandle, B & QPTrim$(RealRec.TownShip);
      Print #RptHandle, B & QPTrim$(RealRec.MORTCODE);
      Print #RptHandle, B & QPTrim$(Using$(M, RealRec.PROPVALU));
      Print #RptHandle, B & QPTrim$(Using$(M, RealRec.EXMPSENI));
      Print #RptHandle, B & QPTrim$(Using$(M, RealRec.EXMPOTHR));
      Print #RptHandle, B & QPTrim$(RealRec.PROPNOT1);
      Print #RptHandle, B & QPTrim$(RealRec.PROPNOT2);
      Print #RptHandle, B & QPTrim$(RealRec.PROPNOT3);
      Print #RptHandle, B & Using$(N, RealRec.CustPin);
  '    Print #RptHandle, B & RealRec.NextRec;
      Print #RptHandle, B & CStr(RealRec.LastYrPrinted);
  '    Print #RptHandle, B & CStr(RealRec.Deleted);
      Print #RptHandle, B & QPTrim$(RealRec.PropAddr);
      Print #RptHandle, B & Using$(N, RealRec.InternalPin);
      Print #RptHandle, B & RealRec.LienYN;
      Print #RptHandle, B & QPTrim$(RealRec.LienDesc);
      Print #RptHandle, B & RealRec.Mock;
  '    Print #RptHandle, B & QPTrim$(RealRec.Image);
      Print #RptHandle, B & QPTrim$(RealRec.OptSearch);
      Print #RptHandle, B & QPTrim$(RealRec.ICPDesc);
      Print #RptHandle, B & QPTrim$(Using$(M, RealRec.BldgVal));
      Print #RptHandle, B
      dodo = dodo + 1
      NextRec = RealRec.NextRec
      'If RealRec.NextRec > 0 Then Stop
    Loop
Skip:
    FrmShowPctComp.ShowPctComp x, NumOfCustRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
    'If dodo > 120 Then Stop
  Next x
  Close
  'End
  
'  For x = 1 To NumOfRRecs
'    Get RHandle, x, RealRec
'    If RealRec.Deleted = 1 Then GoTo Skip
'    If RealRec.CustPin = 0 Then GoTo Skip
'    Print #RptHandle, CStr(x);
'    Print #RptHandle, B & QPTrim$(RealRec.RealPin);
'    Print #RptHandle, B & MakeRegDate(RealRec.PROPDATE);
'    Print #RptHandle, B & QPTrim$(RealRec.GISPOS);
'    Print #RptHandle, B & QPTrim$(RealRec.Map);
'    Print #RptHandle, B & QPTrim$(RealRec.BLOCK);
'    Print #RptHandle, B & QPTrim$(RealRec.LOTNUMB);
'    Print #RptHandle, B & RealRec.LOTACRE;
'    Print #RptHandle, B & CStr(RealRec.PropSize);
'    Print #RptHandle, B & RealRec.PROPDISC;
'    Print #RptHandle, B & RealRec.LateList;
'    Print #RptHandle, B & CStr(RealRec.OptRev1Chrg);
'    Print #RptHandle, B & CStr(RealRec.OptRev2Chrg);
'    Print #RptHandle, B & CStr(RealRec.OptRev3Chrg);
'    Print #RptHandle, B & QPTrim$(RealRec.TownShip);
'    Print #RptHandle, B & QPTrim$(RealRec.MORTCODE);
'    Print #RptHandle, B & QPTrim$(Using$(M, RealRec.PROPVALU));
'    Print #RptHandle, B & QPTrim$(Using$(M, RealRec.EXMPSENI));
'    Print #RptHandle, B & QPTrim$(Using$(M, RealRec.EXMPOTHR));
'    Print #RptHandle, B & QPTrim$(RealRec.PROPNOT1);
'    Print #RptHandle, B & QPTrim$(RealRec.PROPNOT2);
'    Print #RptHandle, B & QPTrim$(RealRec.PROPNOT3);
'    Print #RptHandle, B & Using$(N, RealRec.CustPin);
''    Print #RptHandle, B & RealRec.NextRec;
'    Print #RptHandle, B & CStr(RealRec.LastYrPrinted);
''    Print #RptHandle, B & CStr(RealRec.Deleted);
'    Print #RptHandle, B & QPTrim$(RealRec.PropAddr);
'    Print #RptHandle, B & Using$(N, RealRec.InternalPin);
'    Print #RptHandle, B & RealRec.LienYN;
'    Print #RptHandle, B & QPTrim$(RealRec.LienDesc);
'    Print #RptHandle, B & RealRec.Mock;
''    Print #RptHandle, B & QPTrim$(RealRec.Image);
'    Print #RptHandle, B & QPTrim$(RealRec.OptSearch);
'    Print #RptHandle, B & QPTrim$(RealRec.ICPDesc);
'    Print #RptHandle, B & QPTrim$(Using$(M, RealRec.BldgVal));
'    Print #RptHandle, B
'Skip:
'    FrmShowPctComp.ShowPctComp x, NumOfRRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      Unload FrmShowPctComp
'      Exit Sub
'    End If
'   Next x
'   Unload FrmShowPctComp
'   Close

End Sub
    
Public Sub ProcessVATaxPers()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim NextRec As Integer
  
  M = "###,###,###.##"
  N = "#############"
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  OpenPersPropFile PHandle, NumOfPRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxPersProp.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Personal Property Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    NextRec = CustRec.FirstPersRec
    Do While NextRec > 0
      Get PHandle, NextRec, PersRec
      Print #RptHandle, CStr(NextRec);
      Print #RptHandle, B & QPTrim$(PersRec.PropPin);
      Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.PersVal));
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.MHVALUE));
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.MCVALUE));
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.CVALUE));
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.MTVALUE));
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.EXMPSENI));
      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.EXMPOTHR));
      Print #RptHandle, B & PersRec.DISCOV;
      Print #RptHandle, B & PersRec.LateList;
      Print #RptHandle, B & QPTrim$(PersRec.DESC1);
      Print #RptHandle, B & QPTrim$(PersRec.DESC2);
      Print #RptHandle, B & QPTrim$(PersRec.DESC3);
      Print #RptHandle, B & QPTrim$(PersRec.Desc4);
      Print #RptHandle, B & QPTrim$(PersRec.Desc5);
      Print #RptHandle, B & Using$(N, PersRec.CustPin);
 '    Print #RptHandle, B & PersRec.NextRec;
      Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
 '    Print #RptHandle, B & CStr(PersRec.Deleted);
      Print #RptHandle, B & CStr(PersRec.VehTaxYear);
      Print #RptHandle, B & PersRec.DMVSubmitted;
      Print #RptHandle, B & Using$(N, PersRec.InternalPin);
      Print #RptHandle, B & CStr(PersRec.TaxBillYear);
      Print #RptHandle, B & PersRec.PPTRAYN;
      Print #RptHandle, B & PersRec.Prorate;
      Print #RptHandle, B & CStr(PersRec.ProrateVal);
      Print #RptHandle, B & QPTrim$(PersRec.Vin);
      Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
      Print #RptHandle, B & CStr(PersRec.Weight);
      Print #RptHandle, B & CStr(PersRec.ModYear);
      Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
      Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
      Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
      Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
      Print #RptHandle, B
    Loop
Skip:
    FrmShowPctComp.ShowPctComp x, NumOfCustRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
  Close
  
End Sub
Public Sub ProcessVATaxPers2()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim NextRec As Integer
  
  M = "###,###,###.##"
  N = "#############"
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  OpenPersPropFile PHandle, NumOfPRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxPersProp.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Personal Property Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    NextRec = CustRec.FirstPersRec
  '  If x = 1909 Then Stop
    Do While NextRec > 0
      Get PHandle, NextRec, PersRec
      If PersRec.PersVal > 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE > 0 Then
        GoSub SavePersAndX
        GoSub SaveMHNoX
        GoSub SaveMCNoX
        GoSub SaveMTNoX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal > 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE = 0 Then
        GoSub SavePersAndX
        GoSub SaveMHNoX
        GoSub SaveMCNoX
        GoSub SaveMTNoX
      ElseIf PersRec.PersVal > 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE = 0 Then
        GoSub SavePersAndX
        GoSub SaveMHNoX
        GoSub SaveMCNoX
      ElseIf PersRec.PersVal > 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE = 0 Then
        GoSub SavePersAndX
        GoSub SaveMHNoX
      ElseIf PersRec.PersVal > 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE = 0 Then
        GoSub SavePersAndX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE > 0 Then
        GoSub SaveMHAndX
        GoSub SaveMCNoX
        GoSub SaveMTNoX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE > 0 Then
        GoSub SaveMHAndX
        GoSub SaveMTNoX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE > 0 Then
        GoSub SaveMHAndX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE > 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE = 0 Then
        GoSub SaveMHAndX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE > 0 Then
        GoSub SaveMCAndX
        GoSub SaveMTNoX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE > 0 Then
        GoSub SaveMCAndX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE > 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE = 0 Then
        GoSub SaveMCAndX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE > 0 Then
        GoSub SaveMTAndX
        GoSub SaveFENoX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE > 0 And PersRec.CVALUE = 0 Then
        GoSub SaveMTAndX
      ElseIf PersRec.PersVal = 0 And PersRec.MHVALUE = 0 And PersRec.MCVALUE = 0 And PersRec.MTVALUE = 0 And PersRec.CVALUE > 0 Then
        GoSub SaveFEAndX
      Else
        GoSub SavePersAndX
      End If
      
      
'      Print #RptHandle, CStr(NextRec);
'      Print #RptHandle, B & QPTrim$(PersRec.PropPin);
'      Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.PersVal));
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.MHVALUE));
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.MCVALUE));
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.CVALUE));
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.MTVALUE));
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.EXMPSENI));
'      Print #RptHandle, B & QPTrim$(Using$(M, PersRec.EXMPOTHR));
'      Print #RptHandle, B & PersRec.DISCOV;
'      Print #RptHandle, B & PersRec.LateList;
'      Print #RptHandle, B & QPTrim$(PersRec.DESC1);
'      Print #RptHandle, B & QPTrim$(PersRec.DESC2);
'      Print #RptHandle, B & QPTrim$(PersRec.DESC3);
'      Print #RptHandle, B & QPTrim$(PersRec.Desc4);
'      Print #RptHandle, B & QPTrim$(PersRec.Desc5);
'      Print #RptHandle, B & Using$(N, PersRec.CustPin);
' '    Print #RptHandle, B & PersRec.NextRec;
'      Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
' '    Print #RptHandle, B & CStr(PersRec.Deleted);
'      Print #RptHandle, B & CStr(PersRec.VehTaxYear);
'      Print #RptHandle, B & PersRec.DMVSubmitted;
'      Print #RptHandle, B & Using$(N, PersRec.InternalPin);
'      Print #RptHandle, B & CStr(PersRec.TaxBillYear);
'      Print #RptHandle, B & PersRec.PPTRAYN;
'      Print #RptHandle, B & PersRec.Prorate;
'      Print #RptHandle, B & CStr(PersRec.ProrateVal);
'      Print #RptHandle, B & QPTrim$(PersRec.Vin);
'      Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
'      Print #RptHandle, B & CStr(PersRec.Weight);
'      Print #RptHandle, B & CStr(PersRec.ModYear);
'      Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
'      Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
'      Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
'      Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
'      Print #RptHandle, B
      NextRec = PersRec.NextRec
    Loop
Skip:
    FrmShowPctComp.ShowPctComp x, NumOfCustRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
  Close
  Exit Sub
  
SavePersAndX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.PersVal));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPSENI));
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPOTHR));
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B
 Return
 
SavePersNoX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.PersVal));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B

  Return
 
SaveMHAndX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MHVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPSENI));
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPOTHR));
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B
  Return
 
SaveMHNoX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MHVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B
  Return

SaveMCAndX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MCVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPSENI));
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPOTHR));
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B

  Return
 
SaveMCNoX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MCVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B

  Return
SaveMTAndX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MTVALUE));
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPSENI));
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPOTHR));
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B
  Return
 
SaveMTNoX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MTVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B

  Return
 
SaveFEAndX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.CVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPSENI));
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.EXMPOTHR));
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & CStr(PersRec.OptRev1Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev2Chrg);
    Print #RptHandle, B & CStr(PersRec.OptRev3Chrg);
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B

  Return
 
SaveFENoX:
    Print #RptHandle, CStr(NextRec);
    Print #RptHandle, B & CheckForBad(PersRec.PropPin);
    Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & CheckForBad(Using$(M, PersRec.CVALUE));
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & PersRec.DISCOV;
    Print #RptHandle, B & PersRec.LateList;
    Print #RptHandle, B & CheckForBad(PersRec.DESC1);
    Print #RptHandle, B & CheckForBad(PersRec.DESC2);
    Print #RptHandle, B & CheckForBad(PersRec.DESC3);
    Print #RptHandle, B & CheckForBad(PersRec.Desc4);
    Print #RptHandle, B & CheckForBad(PersRec.Desc5);
    Print #RptHandle, B & Using$(N, PersRec.CustPin);
'    Print #RptHandle, B & PersRec.NextRec;
    Print #RptHandle, B & CStr(PersRec.LastYrPrinted);
'    Print #RptHandle, B & CStr(PersRec.Deleted);
    Print #RptHandle, B & CStr(PersRec.VehTaxYear);
    Print #RptHandle, B & PersRec.DMVSubmitted;
    Print #RptHandle, B & Using$(N, PersRec.InternalPin);
    Print #RptHandle, B & CStr(PersRec.TaxBillYear);
    Print #RptHandle, B & PersRec.PPTRAYN;
    Print #RptHandle, B & PersRec.Prorate;
    Print #RptHandle, B & CStr(PersRec.ProrateVal);
    Print #RptHandle, B & QPTrim$(PersRec.Vin);
    Print #RptHandle, B & QPTrim$(PersRec.MakeMod);
    Print #RptHandle, B & CStr(PersRec.Weight);
    Print #RptHandle, B & CStr(PersRec.ModYear);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & QPTrim$(PersRec.OptSearch);
    Print #RptHandle, B
  Return
 
 
End Sub
    
    
    
Public Sub ProcessMortCodes()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMRecs As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  OpenMortCodeFile MHandle, NumOfMRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxMortCode.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Mortgage Code Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfMRecs
    Get MHandle, x, MortRec
    If MortRec.Deleted = 1 Then GoTo Skip
'    Print #RptHandle, CStr(x);
    Print #RptHandle, QPTrim$(MortRec.MORTCODE);
    Print #RptHandle, B & QPTrim$(MortRec.BName);
    Print #RptHandle, B & QPTrim$(MortRec.Add1);
    Print #RptHandle, B & QPTrim$(MortRec.Add2);
    Print #RptHandle, B & QPTrim$(MortRec.Add3);
    Print #RptHandle, B & QPTrim$(MortRec.Contact);
    Print #RptHandle, B & QPTrim$(MortRec.PHONE);
'    Print #RptHandle, B & CStr(MortRec.Deleted);
    Print #RptHandle, B & QPTrim$(MortRec.XFileNme);
    Print #RptHandle, B
Skip:
    FrmShowPctComp.ShowPctComp x, NumOfMRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
  Close

End Sub

Public Sub ProcessOptSearches()
  Dim PersOptRec As OptPersIdxType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim RealOptRec As OptRealIdxType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim CustOptRec As OptCustIdxType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  OpenPersOptSearchFile PHandle, NumOfPRecs
  OpenRealOptSearchFile RHandle, NumOfRRecs
  OpenCustOptSearchFile CHandle, NumOfCRecs
  
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxOptSrchPers.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Optional Search Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfPRecs
    Get PHandle, x, PersOptRec
    Print #RptHandle, B & QPTrim$(PersOptRec.OptDesc);
    Print #RptHandle, B & CStr(PersOptRec.PersRec);
    Print #RptHandle, B & QPTrim$(PersOptRec.PersPin);
    Print #RptHandle, B
    FrmShowPctComp.ShowPctComp x, NumOfPRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close PHandle
  
  ThisFile = "\VATaxOptSrchReal.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  For x = 1 To NumOfRRecs
    Get RHandle, x, RealOptRec
    Print #RptHandle, B & QPTrim$(RealOptRec.OptDesc);
    Print #RptHandle, B & CStr(RealOptRec.RealRec);
    Print #RptHandle, B & QPTrim$(RealOptRec.RealPin);
    Print #RptHandle, B
    FrmShowPctComp.ShowPctComp x, NumOfRRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Close RHandle
  
  ThisFile = "\VATaxOptSrchCust.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, CustOptRec
    Print #RptHandle, B & QPTrim$(CustOptRec.OptDesc);
    Print #RptHandle, B & CStr(CustOptRec.CustRec);
    Print #RptHandle, B & CStr(CustOptRec.CustPin);
    Print #RptHandle, B
    FrmShowPctComp.ShowPctComp x, NumOfCRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload FrmShowPctComp
  Close

End Sub

Public Sub ProcessTownships()
  Dim TownRec As TownshipType
  Dim THandle As Integer
  Dim NumOfTRecs As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  OpenTownshipFile THandle, NumOfTRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxTownships.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Townships Export"
  FrmShowPctComp.Show
  DoEvents
  For x = 1 To NumOfTRecs
    Get THandle, x, TownRec
    If x = 1 Then
      Print #RptHandle, QPTrim$(TownRec.TownShip)
    Else
      Print #RptHandle, B & QPTrim$(TownRec.TownShip)
    End If
'    Print #RptHandle, B
    FrmShowPctComp.ShowPctComp x, NumOfTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
  Close
  
End Sub

Public Sub ProcessSystemSetup()
  Dim SysRec As TaxMasterType
  Dim SHandle As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim P As String
  M = "###,###,###.##"
  N = "#############"
  P = "###.###"
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxSystemSetup.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenTaxSetUpFile SHandle

  Get SHandle, 1, SysRec
  Print #RptHandle, QPTrim$(SysRec.Name); '1
  Print #RptHandle, B & QPTrim$(SysRec.Add1); '2
  Print #RptHandle, B & QPTrim$(SysRec.Add2); '3
  Print #RptHandle, B & QPTrim$(SysRec.City); '4
  Print #RptHandle, B & QPTrim$(SysRec.Zip); '5
  Print #RptHandle, B & QPTrim$(SysRec.TaxSt); '6
  Print #RptHandle, B & CStr(SysRec.TaxForm); '7
  Print #RptHandle, B & CStr(SysRec.RTaxYear); '8
  Print #RptHandle, B & CStr(SysRec.LateForm); '9
  Print #RptHandle, B & SysRec.WarnInt; '10
  Print #RptHandle, B & QPTrim$(Using$(M, SysRec.MinBill)); '11
  Print #RptHandle, B & QPTrim$(SysRec.AcctgMethod); '12
  Print #RptHandle, B & CStr(SysRec.MinTxOpt); '13
  Print #RptHandle, B & QPTrim$(SysRec.TownState); '14
'  For x = 1 To 5
'    Print #RptHandle, B & CStr(SysRec.CurrRYrInt(x));
'  Next x
  Print #RptHandle, B & CStr(SysRec.CurrRYrIntInUse); '15
'  For x = 1 To 5
'    Print #RptHandle, B & CStr(SysRec.CurrPYrInt(x));
'  Next x
  Print #RptHandle, B & CStr(SysRec.CurrPYrIntInUse); '16
  Print #RptHandle, B & CStr(SysRec.PastYrInt); '17
'  Print #RptHandle, B & Using$(P, SysRec.PenPct);
'  Print #RptHandle, B & CStr(SysRec.PenIdx);
  Print #RptHandle, B & SysRec.CntrlDepYN; '18
'  Print #RptHandle, B & SysRec.PriorYrMltRevYN;
  Print #RptHandle, B & QPTrim$(SysRec.OverPayGLNum); '19
  Print #RptHandle, B & SysRec.PenPrncTaxYN; '20
  Print #RptHandle, B & SysRec.PenIntYN; '21
  Print #RptHandle, B & SysRec.PenAdvYN; '22
  Print #RptHandle, B & SysRec.PenLateLstYN; '23
  Print #RptHandle, B & SysRec.PenOpt1YN; '24
  Print #RptHandle, B & SysRec.PenOpt2YN; '25
  Print #RptHandle, B & SysRec.PenOpt3YN; '26
  Print #RptHandle, B & SysRec.IntPrncTaxYN; '27
  Print #RptHandle, B & SysRec.IntIntYN; '28
  Print #RptHandle, B & SysRec.IntAdvYN; '29
  Print #RptHandle, B & SysRec.IntLateLstYN; '30
  Print #RptHandle, B & SysRec.IntOpt1YN; '31
  Print #RptHandle, B & SysRec.IntOpt2YN; '32
  Print #RptHandle, B & SysRec.IntOpt3YN; '33
  Print #RptHandle, B & QPTrim$(SysRec.OptRev1); '34
  Print #RptHandle, B & QPTrim$(SysRec.OptRev2); '35
  Print #RptHandle, B & QPTrim$(SysRec.OptRev3); '36
  Print #RptHandle, B & MakeRegDate(SysRec.DiscRXDate);
  Print #RptHandle, B & Using(P, SysRec.DisRPct);
  Print #RptHandle, B & MakeRegDate(SysRec.DiscPXDate);
  Print #RptHandle, B & Using(P, SysRec.DisPPct);
  Print #RptHandle, B & QPTrim$(SysRec.OptSrchCust);
  Print #RptHandle, B & QPTrim$(SysRec.OptSrchProp);
  For x = 1 To 5
    Print #RptHandle, B & QPTrim$(SysRec.CountyName(x));
  Next x
'  For x = 1 To 5
'    Print #RptHandle, B & CStr(SysRec.CountyNum(x));
'  Next x
'  Print #RptHandle, B & SysRec.UseCountyYN;
  Print #RptHandle, B & SysRec.RealPersSplit;
'  For x = 1 To 5
'    Print #RptHandle, B & CStr(SysRec.CycleNum(x));
'  Next x
  For x = 1 To 5
    Print #RptHandle, B & QPTrim$(SysRec.CycleName(x));
  Next x
'  Print #RptHandle, B & SysRec.UseCyclesYN;
  Print #RptHandle, B & QPTrim$(SysRec.CDCashGL);
  Print #RptHandle, B & QPTrim$(SysRec.CDSubGL);
  For x = 1 To 6
    Print #RptHandle, B & QPTrim$(SysRec.ClassName(x));
  Next x
  Print #RptHandle, B & CStr(SysRec.MultiYear);
  Print #RptHandle, B & Using$(P, SysRec.PPTRADisc);
  Print #RptHandle, B & QPTrim$(Using$(M, SysRec.MaxVehTaxVal));
'  Print #RptHandle, B & MakeRegDate(SysRec.LawChngDate);
  Print #RptHandle, B & QPTrim$(Using$(M, SysRec.MinVehTaxVal));
  Print #RptHandle, B & SysRec.PPTRAYN;
  Print #RptHandle, B & SysRec.PenPenaltyYN;
  Print #RptHandle, B & SysRec.IntPenaltyYN;
  Print #RptHandle, B & QPTrim$(SysRec.POptRev1);
  Print #RptHandle, B & QPTrim$(SysRec.POptRev2);
  Print #RptHandle, B & QPTrim$(SysRec.POptRev3);
  Print #RptHandle, B & SysRec.PenPersYN;
  Print #RptHandle, B & SysRec.IntPersYN;
  Print #RptHandle, B & CStr(SysRec.PersPayOrder);
  Print #RptHandle, B & SysRec.PenMTYN;
  Print #RptHandle, B & SysRec.IntMTYN;
  Print #RptHandle, B & CStr(SysRec.MTPayOrder);
  Print #RptHandle, B & SysRec.PenMCYN;
  Print #RptHandle, B & SysRec.IntMCYN;
  Print #RptHandle, B & CStr(SysRec.MCPayOrder);
  Print #RptHandle, B & SysRec.PenFEYN;
  Print #RptHandle, B & SysRec.IntFEYN;
  Print #RptHandle, B & CStr(SysRec.FEPayOrder);
  Print #RptHandle, B & SysRec.PenMHYN;
  Print #RptHandle, B & SysRec.IntMHYN;
  Print #RptHandle, B & CStr(SysRec.MHPayOrder);
  Print #RptHandle, B & SysRec.PenPIntYN;
  Print #RptHandle, B & SysRec.IntPIntYN;
  Print #RptHandle, B & CStr(SysRec.PIntPayOrder);
  Print #RptHandle, B & SysRec.PenPPenYN;
  Print #RptHandle, B & SysRec.IntPPenYN;
  Print #RptHandle, B & CStr(SysRec.PPenPayOrder);
  Print #RptHandle, B & SysRec.PenPOpt1YN;
  Print #RptHandle, B & SysRec.IntPOpt1YN;
  Print #RptHandle, B & CStr(SysRec.POpt1PayOrder);
  Print #RptHandle, B & SysRec.PenPOpt2YN;
  Print #RptHandle, B & SysRec.IntPOpt2YN;
  Print #RptHandle, B & CStr(SysRec.POpt2PayOrder);
  Print #RptHandle, B & SysRec.PenPOpt3YN;
  Print #RptHandle, B & SysRec.IntPOpt3YN;
  Print #RptHandle, B & CStr(SysRec.POpt3PayOrder);
  Print #RptHandle, B & CStr(SysRec.PTaxYear);
  Print #RptHandle, B & QPTrim$(SysRec.OptSrchPers);
  Print #RptHandle, B
  Close

End Sub

Public Sub ProcessLaserItemized()
  Dim LtrRec As TxBillLaserItemized
  Dim LHandle As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxPersLaserItemized.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenLaserPersItemized LHandle
  NumOfRecs = LOF(LHandle) / Len(LtrRec)
  If NumOfRecs = 0 Then GoTo Done
  Get LHandle, 1, LtrRec
  Print #RptHandle, LtrRec.TxtHead1;
  Print #RptHandle, B & LtrRec.TxtHead2;
  Print #RptHandle, B & LtrRec.txtHead3;
  Print #RptHandle, B & LtrRec.txtHead4;
  Print #RptHandle, B & LtrRec.txtHead5;
  Print #RptHandle, B & LtrRec.txtOpt1;
  Print #RptHandle, B & LtrRec.TxtOpt2;
  Print #RptHandle, B & LtrRec.TxtOpt3;
  Print #RptHandle, B & LtrRec.txtPgph0;
  Print #RptHandle, B & LtrRec.txtPgph1;
  Print #RptHandle, B & LtrRec.txtPgph2;
  Print #RptHandle, B & LtrRec.txtPgph3;
  Print #RptHandle, B & LtrRec.txtPgph4;
  Print #RptHandle, B & CStr(LtrRec.dologo);
  Print #RptHandle, B & CStr(LtrRec.UseBarCode);
  Print #RptHandle, B
Done:
  Close
  
End Sub

Public Sub ProcessMessages()
  Dim MessRec As TaxMessRecType
  Dim MHandle As Integer
  Dim NumOfMRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  OpenTaxMessage MHandle, NumOfMRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\VATaxMessages.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Messages Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfMRecs
    Get MHandle, x, MessRec
    For Y = 1 To 15
      If Y = 1 Then
        Print #RptHandle, MessRec.MessLine(Y).Msg;
      Else
        Print #RptHandle, B & MessRec.MessLine(Y).Msg;
      End If
    Next Y
    Print #RptHandle, B & CStr(MessRec.TaxRec);
    Print #RptHandle, B
    FrmShowPctComp.ShowPctComp x, NumOfMRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
  Close


End Sub
Public Sub ProcessLaserStandard()
  Dim LtrRec As TxBillLaser1DefaultsType
  Dim LHandle As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxPersLaserStandard.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenTxBillPersFile LHandle
  NumOfRecs = LOF(LHandle) / Len(LtrRec)
  If NumOfRecs = 0 Then GoTo Real
  Get LHandle, 1, LtrRec
  Print #RptHandle, LtrRec.TxtHead1;
  Print #RptHandle, B & LtrRec.TxtHead2;
  Print #RptHandle, B & LtrRec.txtOpt1;
  Print #RptHandle, B & LtrRec.TxtOpt2;
  Print #RptHandle, B & LtrRec.TxtOpt3;
  Print #RptHandle, B & LtrRec.TxtOpt4;
  Print #RptHandle, B & LtrRec.txtPgph0;
  Print #RptHandle, B & LtrRec.txtPgph1;
  Print #RptHandle, B & LtrRec.txtPgph2;
  Print #RptHandle, B & LtrRec.txtPgph3;
  Print #RptHandle, B & LtrRec.txtPgph4;
  Print #RptHandle, B & LtrRec.txtPgph5;
  Print #RptHandle, B & LtrRec.txtPgph6;
  Print #RptHandle, B & LtrRec.txtPgph7;
  Print #RptHandle, B & LtrRec.TxtOpt5;
  Print #RptHandle, B & LtrRec.txtHead4;
  Print #RptHandle, B & LtrRec.txtHead5;
  Print #RptHandle, B & LtrRec.txtHead6;
  Print #RptHandle, B & LtrRec.TxtOpt6;
  Print #RptHandle, B & LtrRec.TxtOpt7;
  Print #RptHandle, B & CStr(LtrRec.dologo);
  Print #RptHandle, B & CStr(LtrRec.UseBarCode);
  Print #RptHandle, B

Real:
  Close
  
  ThisFile = "\VATaxRealLaserStandard.txt"
  If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
    KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenTxBillRealFile LHandle
  NumOfRecs = LOF(LHandle) / Len(LtrRec)
  If NumOfRecs = 0 Then GoTo Done
  Get LHandle, 1, LtrRec
  Print #RptHandle, LtrRec.TxtHead1;
  Print #RptHandle, B & LtrRec.TxtHead2:
  Print #RptHandle, B & LtrRec.txtOpt1;
  Print #RptHandle, B & LtrRec.TxtOpt2;
  Print #RptHandle, B & LtrRec.TxtOpt3;
  Print #RptHandle, B & LtrRec.TxtOpt4;
  Print #RptHandle, B & LtrRec.txtPgph0;
  Print #RptHandle, B & LtrRec.txtPgph1;
  Print #RptHandle, B & LtrRec.txtPgph2;
  Print #RptHandle, B & LtrRec.txtPgph3;
  Print #RptHandle, B & LtrRec.txtPgph4;
  Print #RptHandle, B & LtrRec.txtPgph5;
  Print #RptHandle, B & LtrRec.txtPgph6;
  Print #RptHandle, B & LtrRec.txtPgph7;
  Print #RptHandle, B & LtrRec.TxtOpt5;
  Print #RptHandle, B & LtrRec.txtHead4;
  Print #RptHandle, B & LtrRec.txtHead5;
  Print #RptHandle, B & LtrRec.txtHead6;
  Print #RptHandle, B & LtrRec.TxtOpt6;
  Print #RptHandle, B & LtrRec.TxtOpt7;
  Print #RptHandle, B & CStr(LtrRec.dologo);
  Print #RptHandle, B & CStr(LtrRec.UseBarCode);
  Print #RptHandle, B
Done:
  Close
  
End Sub
    
Public Sub ProcessLateLetter()
  Dim LtrRec As TAXLateLetterType
  Dim LHandle As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxLateLetter.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenLateLtrFile LHandle
  NumOfRecs = LOF(LHandle) / Len(LtrRec)
  If NumOfRecs = 0 Then GoTo Done
  Get LHandle, 1, LtrRec
  Print #RptHandle, LtrRec.Head1;
  Print #RptHandle, B & LtrRec.Head2;
  Print #RptHandle, B & LtrRec.Head3;
  Print #RptHandle, B & LtrRec.Head4;
  Print #RptHandle, B & LtrRec.Head5;
  For x = 1 To 20
    Print #RptHandle, B & LtrRec.Body(x);
  Next x
Done:
  Close

End Sub

Public Sub ProcessRateTables()
  Dim RevRec As OptRevRateTablesType
  Dim RHandle As Integer
  Dim NumOfRRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  Dim M As String
  Dim N As String
  Dim P As String
  M = "###,###,###.##"
  N = "#############"
  P = "###.###"
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxRateTables.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenTaxRateTables RHandle, NumOfRRecs
  For x = 1 To NumOfRRecs
    Get RHandle, x, RevRec
    Print #RptHandle, CStr(RevRec.OptRevNum);
    Print #RptHandle, B & QPTrim$(RevRec.Desc);
    Print #RptHandle, B & RevRec.Type;
    Print #RptHandle, B & RevRec.StepType;
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RevRec.FromAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RevRec.ToAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RevRec.TaxFAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & Using$(P, RevRec.TaxPAmt(Y));
    Next Y
    Print #RptHandle, B & QPTrim$(Using$(M, RevRec.FlatAmt));
    Print #RptHandle, B & CStr(RevRec.Deleted);
    Print #RptHandle, B & RevRec.RevType;
    Print #RptHandle, B & QPTrim$(RevRec.Comment);
    Print #RptHandle, B
  Next x
  
End Sub

Public Sub ProcessPenaltyTables()
  Dim RevRec As PenaltyRateTablesType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  Dim M As String
  Dim N As String
  Dim P As String
  M = "###,###,###.##"
  N = "#############"
  P = "###.###"
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxPenaltyRateTables.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenTaxPenRateTbls PHandle, NumOfPRecs
  For x = 1 To NumOfPRecs
    Get PHandle, x, RevRec
    Print #RptHandle, QPTrim$(RevRec.Desc);
    For Y = 1 To 10
      Print #RptHandle, B & CStr(RevRec.RateType(Y));
    Next Y
    Print #RptHandle, B & RevRec.StepType;
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RevRec.FromAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RevRec.ToAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RevRec.TaxFAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(P, RevRec.TaxPAmt(Y)));
    Next Y
    Print #RptHandle, B & QPTrim$(Using$(M, RevRec.FlatAmt));
    Print #RptHandle, B & CStr(RevRec.Deleted);
    Print #RptHandle, B & QPTrim$(RevRec.Comment);
    Print #RptHandle, B & RevRec.BillType;
    Print #RptHandle, B
  Next x
  
  Close

End Sub
    
Public Sub ProcessRealGLPay()
  Dim RGLRec As TaxRAcctsType
  Dim RHandle As Integer
  Dim NumOfRRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxGLRealPay.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenRTaxGLInterPay RHandle
  NumOfRRecs = LOF(RHandle) / Len(RGLRec)
  FrmShowPctComp.Label1 = "Tax Real GL Pay Export"
  FrmShowPctComp.Show
  DoEvents
  For x = 1 To NumOfRRecs
    Get RHandle, x, RGLRec
    For Y = 1 To 51
      Print #RptHandle, CStr(RGLRec.TaxAcct(Y).TaxYear);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).TaxDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).TaxCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).IntDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).IntCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).AdvDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).AdvCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).LtLstDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).LtLstCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).PenDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).PenCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt1DBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt1CRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt2DBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt2CRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt3DBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt3CRAcct);
      Print #RptHandle, B
    Next Y
    FrmShowPctComp.ShowPctComp x, NumOfRRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
Done:
  Close
  Unload FrmShowPctComp
  
End Sub
Public Sub ProcessPersGLPay()
  Dim PGLRec As TaxPAcctsType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
   ThisFile = "\VATaxGLPersPay.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenPTaxGLInterPay PHandle
  NumOfPRecs = LOF(PHandle) / Len(PGLRec)
  FrmShowPctComp.Label1 = "Tax Personal GL Pay Export"
  FrmShowPctComp.Show
  DoEvents
  For x = 1 To NumOfPRecs
    Get PHandle, x, PGLRec
    For Y = 1 To 51
      Print #RptHandle, CStr(PGLRec.TaxAcct(Y).TaxYear);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PersDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PersCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MTDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MTCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MCDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MCCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).FEDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).FECRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MHDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MHCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).IntDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).IntCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PenDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PenCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt1DBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt1CRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt2DBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt2CRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt3DBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt3CRAcct);
      Print #RptHandle, B
    Next Y
    FrmShowPctComp.ShowPctComp x, NumOfPRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
Done:
  Close
  Unload FrmShowPctComp


End Sub

Public Sub ProcessRealGLBill()
  Dim RGLRec As TaxRAcctsType
  Dim RHandle As Integer
  Dim NumOfRRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\VATaxGLRealBill.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenRTaxGLInterBill RHandle
  NumOfRRecs = LOF(RHandle) / Len(RGLRec)
  FrmShowPctComp.Label1 = "Tax Real GL Billing Export"
  FrmShowPctComp.Show
  DoEvents
  
  For x = 1 To NumOfRRecs
    Get RHandle, x, RGLRec
    For Y = 1 To 51
      Print #RptHandle, CStr(RGLRec.TaxAcct(Y).TaxYear);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).TaxDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).TaxCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).IntDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).IntCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).AdvDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).AdvCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).LtLstDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).LtLstCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).PenDBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).PenCRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt1DBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt1CRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt2DBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt2CRAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt3DBAcct);
      Print #RptHandle, B & QPTrim$(RGLRec.TaxAcct(Y).Opt3CRAcct);
      Print #RptHandle, B
    Next Y
    FrmShowPctComp.ShowPctComp x, NumOfRRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
Done:
  Close
  Unload FrmShowPctComp
  
End Sub
    
Public Sub ProcessPersGLBill()
  Dim PGLRec As TaxPAcctsType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  ThisFile = "\VATaxGLPersBill.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  B = "|"
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  OpenPTaxGLInterBill PHandle
  NumOfPRecs = LOF(PHandle) / Len(PGLRec)
  FrmShowPctComp.Label1 = "Tax Personal GL Billing Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfPRecs
    Get PHandle, x, PGLRec
    For Y = 1 To 51
      Print #RptHandle, CStr(PGLRec.TaxAcct(Y).TaxYear);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PersDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PersCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MTDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MTCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MCDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MCCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).FEDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).FECRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MHDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).MHCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).IntDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).IntCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PenDBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).PenCRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt1DBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt1CRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt2DBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt2CRAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt3DBAcct);
      Print #RptHandle, B & QPTrim$(PGLRec.TaxAcct(Y).Opt3CRAcct);
      Print #RptHandle, B
    Next Y
    FrmShowPctComp.ShowPctComp x, NumOfPRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
Done:
  Close
  Unload FrmShowPctComp

End Sub

Public Sub ProcessTransHist()
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim P As String
  M = "###,###,###.##"
  N = "#############"
  P = "###.###"
  
  ThisFile = "\VATaxTaxTrans.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  B = "|"
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  OpenTaxTransFile THandle, NumOfTRecs
  
  FrmShowPctComp.Label1 = "Tax Transaction History Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    Print #RptHandle, MakeRegDate(TransRec.TransDate);
    Print #RptHandle, B & CStr(TransRec.TaxYear);
    Print #RptHandle, B & CStr(TransRec.TranType);
    Print #RptHandle, B & TransRec.BillType;
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Amount));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Collection));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.CollectionPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Interest));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.InterestPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.LateList));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.LateListPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Penalty));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PenaltyPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PrePaidAmt));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PrePaidBal));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PrePaidUsed));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle1));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle1Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle2));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle2Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle3));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle3Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle4));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle4Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle5));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle5Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt1));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt1Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt2));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt2Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt3));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt3Pd));
    Print #RptHandle, B & QPTrim(TransRec.Description);
    Print #RptHandle, B & TransRec.Posted2GL;
    Print #RptHandle, B & CStr(TransRec.CustomerRec);
'    Print #RptHandle, B & CStr(TransRec.LastTrans);
    Print #RptHandle, B & CStr(TransRec.BelongTo);
    Print #RptHandle, B & TransRec.DMVSubmitted;
    Print #RptHandle, B & CStr(TransRec.DMVBatch);
    Print #RptHandle, B & CStr(TransRec.Altered);
    Print #RptHandle, B & TransRec.FromPrePay;
    Print #RptHandle, B & QPTrim$(TransRec.PersPin);
    Print #RptHandle, B & QPTrim$(TransRec.RealPin);
    Print #RptHandle, B & CStr(TransRec.CustPin);
    Print #RptHandle, B & CStr(TransRec.InternalPin);
    Print #RptHandle, B & MakeRegDate(TransRec.DiscXDate);
    Print #RptHandle, B & Using$(P, TransRec.DiscAmt);
    Print #RptHandle, B & CStr(TransRec.OperNum);
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PersVal));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PPTRAVal));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PPTRADisc));
    Print #RptHandle, B & QPTrim$(TransRec.CntyPara);
    Print #RptHandle, B & QPTrim$(TransRec.CyclPara);
    Print #RptHandle, B & QPTrim$(TransRec.TShpPara);
    Print #RptHandle, B & CStr(TransRec.PPTRARmvl);
    Print #RptHandle, B & MakeRegDate(TransRec.PPTRARmvlDate);
    Print #RptHandle, B & ParseBillNum(TransRec.Description);
    Print #RptHandle, B
    FrmShowPctComp.ShowPctComp x, NumOfTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
Done:
  Close
  Unload FrmShowPctComp
 
End Sub
    
Public Sub ProcessVAOptRevRateTables()
  Dim RateRec As OptRevRateTablesType
  Dim THandle As Integer
  Dim NumOfTRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim P As String
  M = "###,###,###.##"
  N = "#############"
  P = "###.###"
  
  ThisFile = "\VATaxOptRateTbls.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  B = "|"
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  OpenTaxRateTables THandle, NumOfTRecs
  
  FrmShowPctComp.Label1 = "Tax Optional Rate Tables Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfTRecs
    Get THandle, x, RateRec
    If RateRec.Deleted = True Then GoTo Skip
    Print #RptHandle, RateRec.OptRevNum;
    Print #RptHandle, B & QPTrim$(RateRec.Desc);
    Print #RptHandle, B & RateRec.Type;
    Print #RptHandle, B & RateRec.StepType;
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RateRec.FromAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RateRec.ToAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(M, RateRec.TaxFAmt(Y)));
    Next Y
    For Y = 1 To 10
      Print #RptHandle, B & QPTrim$(Using$(P, RateRec.TaxPAmt(Y)));
    Next Y
    Print #RptHandle, B & QPTrim$(Using$(M, RateRec.FlatAmt));
    Print #RptHandle, B & RateRec.RevType;
    Print #RptHandle, B
Skip:
    FrmShowPctComp.ShowPctComp x, NumOfTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next x
Done:
  Close
  Unload FrmShowPctComp
 

End Sub

Public Function CheckForBad(ByVal TestStr As String, Optional IsMessage As Boolean = False) As String
  Dim chVal As Integer
  Dim x As Integer
  Dim Lth As Integer
  Dim rtn As String
  
  rtn = TestStr
  Lth = Len(TestStr)
  For x = 1 To Lth
    chVal = Asc(Mid(rtn, x, 1))
    Select Case chVal
        Case 32 To 126:
            ' Char is OK
        Case 13, 10:
            ' Carriage Return and Linefeed are OK.
            If IsMessage = False Then Mid$(rtn, x, 1) = " "
        Case Else:
            ' Non-printable... assume whole string is bad.
'            Mid$(rtn, x, 1) = " "
            rtn = ""
            Exit For
    End Select
  Next x
  
  CheckForBad = rtn
End Function
    
Public Sub ProcessVABalance()
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  
  Dim RealRec As PropertyRecType
  
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim RHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim M As String
  Dim N As String
  Dim P As String
  Dim Amount As Double
  Dim Collection As Double
  Dim CollectionPd As Double
  Dim Interest As Double
  Dim InterestPd As Double
  Dim LateList As Double
  Dim LateListPd As Double
  Dim Penalty As Double
  Dim PenaltyPd As Double
  Dim PrePaidAmt As Double
  Dim PrePaidUsed As Double
  Dim Principle1 As Double
  Dim Principle1Pd As Double
  Dim Principle2 As Double
  Dim Principle2Pd As Double
  Dim Principle3 As Double
  Dim Principle3Pd As Double
  Dim Principle4 As Double
  Dim Principle4Pd As Double
  Dim Principle5 As Double
  Dim Principle5Pd As Double
  Dim RevOpt1 As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2 As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3 As Double
  Dim RevOpt3Pd As Double
  Dim Balance As Double
  Dim NextRec As Long
  
  M = "###,###,###.##"
  N = "#############"
  P = "###.###"
  
  ThisFile = "VATaxBalance.txt"
  If DirExists(StartPath + "\VATAXConvToTxt") Then
    If Exist(StartPath + "\VATAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\VATAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\VATAXConvToTxt"
  End If
  
  Dim NumOfRRecs As Long
  
  B = "|"
  RptName$ = StartPath + "\VATAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  OpenTaxTransFile THandle, NumOfTRecs
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenRealPropFile RHandle, NumOfRRecs
  
  FrmShowPctComp.Label1 = "Tax Balance Export"
  FrmShowPctComp.Show
  DoEvents
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    If TaxCust.Deleted = -1 Then GoTo NoBalance
    Balance = GetCustBalance(x, -1)
    If Balance = 0 Then GoTo NoBalance
    If Balance < 0 Then
      GoSub SendThisCreditBal
    Else
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get THandle, NextRec, TransRec
        If TransRec.TranType = 1 Then
          Balance# = OldRound#(TransRec.Revenue.LateList + TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2 + TransRec.Revenue.Principle3 + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TransRec.Revenue.Interest + TransRec.Revenue.Penalty + TransRec.Revenue.Collection + TransRec.Revenue.RevOpt1 + TransRec.Revenue.RevOpt2 + TransRec.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd + TransRec.Revenue.Principle3Pd + TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TransRec.Revenue.InterestPd + TransRec.Revenue.PenaltyPd + TransRec.Revenue.CollectionPd + TransRec.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TransRec.Revenue.RevOpt1Pd + TransRec.Revenue.RevOpt2Pd + TransRec.Revenue.RevOpt3Pd + TransRec.DiscAmt + TransRec.PPTRADisc - TransRec.PPTRARmvl))
          If Balance > 0 Then
            GoSub SendThis
          End If
        End If
        NextRec = TransRec.LastTrans
      Loop
    End If
    FrmShowPctComp.ShowPctComp x, NumOfCRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
NoBalance:
  Next x
Done:
  Close
  Unload FrmShowPctComp
  Exit Sub
  
SendThisCreditBal:
    Print #RptHandle, QPTrim$(Using$(M, Balance));
    Print #RptHandle, B & Date;
    Print #RptHandle, B & CStr(Year(Date));
    Print #RptHandle, B & CStr(22);
    Print #RptHandle, B & "R";
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0)); '10
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0)); '20
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0)); '30
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & "Over Payment";
    Print #RptHandle, B & CStr(TaxCust.Acct);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "0";
    Print #RptHandle, B & Using$(P, 0);
    Print #RptHandle, B & CStr(0);
    Print #RptHandle, B & "0"; '40
    Print #RptHandle, B & "N";
    '-------------------------------------------
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & "12/31/1979"; '46
    Print #RptHandle, B
  Return
  
SendThis:
    'RealRec
    'TransRec.
    'TransRec.CustomerRec
    If TransRec.CustomerRec <= 0 Then Stop
    Get #CHandle, TransRec.CustomerRec, TaxCust
    
    'If TaxCust.FirstPropRec <= 0 Then Stop
    If TaxCust.FirstPropRec > 0 Then
      Get #RHandle, TaxCust.FirstPropRec, RealRec
    End If
    
    'If Len(QPTrim(RealRec.RealPin)) <= 0 Then Stop
    
    'If TransRec.CustomerRec = 1643 Then Stop
    
    Print #RptHandle, QPTrim$(Using$(M, Balance));
    Print #RptHandle, B & MakeRegDate(TransRec.TransDate);
    Print #RptHandle, B & CStr(TransRec.TaxYear);
    Print #RptHandle, B & CStr(TransRec.TranType);
    Print #RptHandle, B & CheckForBad(TransRec.BillType);
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Amount));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Collection));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.CollectionPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Interest));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.InterestPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.LateList));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.LateListPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Penalty));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PenaltyPd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PrePaidAmt));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PrePaidBal));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.PrePaidUsed));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle1));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle1Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle2));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle2Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle3));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle3Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle4));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle4Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle5));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.Principle5Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt1));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt1Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt2));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt2Pd));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt3));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.Revenue.RevOpt3Pd));
    Print #RptHandle, B & QPTrim(TransRec.Description);
    Print #RptHandle, B & CStr(TransRec.CustomerRec);
    Print #RptHandle, B & QPTrim$(TransRec.PersPin);
    Print #RptHandle, B & QPTrim$(RealRec.RealPin);
    
    'If Len(QPTrim$(TransRec.RealPin)) > 1 Then Stop
    'If Len(QPTrim$(TransRec.PersPin)) > 1 Then Stop
        
    Print #RptHandle, B & Using$(P, TransRec.DiscAmt);
    Print #RptHandle, B & CStr(TransRec.OperNum);
    Print #RptHandle, B & ParseBillNum(TransRec.Description);
    Print #RptHandle, B & TransRec.Posted2GL;
    '----------------------------------------------------------------------
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PersVal));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PPTRAVal));
'    If TransRec.PPTRADisc > 0 Then Stop
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PPTRADisc));
    Print #RptHandle, B & QPTrim$(Using$(M, TransRec.PPTRARmvl));
    Print #RptHandle, B & MakeRegDate(TransRec.PPTRARmvlDate);
    '----------------------------------------------------------------------
    Print #RptHandle, B
  Return

End Sub

'Private Sub ProcessVAPersPPTRARemoval()
'  Dim PPRec As VAPPTaxBillType
'  Dim PPHandle As Integer
'  Dim NumOfPPRecs As Long
'  Dim MyPath$, ThisFile$
'  Dim x As Integer
'  Dim RptFile$
'  Dim RptHandle As Integer
'
'
'
'End Sub
