Attribute VB_Name = "modNCTaxExportData"
Option Explicit

Public Sub ProcessNCTaxCust()
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  NCOpenTaxCustFile THandle, NumOfTRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxCustData.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle

  FrmShowPctComp.Label1 = "Tax Customer Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfTRecs
    Get THandle, x, TaxCust
'    If x = 436 Then Stop
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
    Print #RptHandle, B & CheckForBad(TaxCust.CountyAcctString);
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

Public Sub ProcessNCTaxReal()
  Dim RealRec As NCPropertyRecType
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
  
  M = "###,###,###.##"
  N = "#############"
  NCOpenTaxCustFile CustHandle, NumOfCustRecs
  NCOpenRealPropFile RHandle, NumOfRRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxRealProp.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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
'      Print #RptHandle, B & CStr(RealRec.NextRec);
      Print #RptHandle, B & CStr(RealRec.LastYrPrinted);
'      Print #RptHandle, B & CStr(RealRec.Deleted);
      Print #RptHandle, B & QPTrim$(RealRec.PropAddr);
      Print #RptHandle, B & Using$(N, RealRec.InternalPin);
      Print #RptHandle, B & RealRec.LienYN;
      Print #RptHandle, B & QPTrim$(RealRec.LienDesc);
      Print #RptHandle, B & RealRec.Mock;
'      Print #RptHandle, B & QPTrim$(RealRec.Image);
      Print #RptHandle, B & QPTrim$(RealRec.OptSearch);
      Print #RptHandle, B & QPTrim$(RealRec.ICPDesc);
      Print #RptHandle, B
      NextRec = RealRec.NextRec
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
  Close
  
'  For x = 1 To NumOfRRecs
'    Get RHandle, x, RealRec
'    If RealRec.Deleted = 1 Then GoTo Skip
'    If RealRec.CustPin < 1 Then GoTo Skip
'    If RealRec.CustPin > NumOfCustRecs Then GoTo Skip
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
''    Print #RptHandle, B & CStr(RealRec.NextRec);
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

Public Sub ProcessNCTaxPers()
  Dim PersRec As NCPersonalRecType
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
  
  NCOpenTaxCustFile CustHandle, NumOfCustRecs
  NCOpenPersPropFile PHandle, NumOfPRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxPersProp.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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
        Print #RptHandle, B & CheckForBad(PersRec.PropPin);
        Print #RptHandle, B & MakeRegDate(PersRec.PROPDATE);
        Print #RptHandle, B & CheckForBad(Using$(M, PersRec.PersVal));
        Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MHVALUE));
        Print #RptHandle, B & CheckForBad(Using$(M, PersRec.MCVALUE));
        Print #RptHandle, B & CheckForBad(Using$(M, PersRec.CVALUE));
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
        Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
        Print #RptHandle, B
      
      NextRec = PersRec.NextRec
    Loop
  
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

Public Sub ProcessNCTaxPers2()
  Dim PersRec As NCPersonalRecType
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
  Dim cnt As Integer
  M = "###,###,###.##"
  N = "#############"
  
  NCOpenTaxCustFile CustHandle, NumOfCustRecs
  NCOpenPersPropFile PHandle, NumOfPRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxPersProp.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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
      
      NextRec = PersRec.NextRec
    Loop
      
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
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
    Print #RptHandle, B & CheckForBad(PersRec.OptSearch);
    Print #RptHandle, B

  Return
  
End Sub


Public Sub ProcessNCMortCodes()
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMRecs As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  NCOpenMortCodeFile MHandle, NumOfMRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxMortCode.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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

Public Sub ProcessNCOptSearches()
  Dim PersOptRec As OptPersIdxType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim RealOptRec As NCOptRealIdxType
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
  
  NCOpenPersOptSearchFile PHandle, NumOfPRecs
  NCOpenRealOptSearchFile RHandle, NumOfRRecs
  NCOpenCustOptSearchFile CHandle, NumOfCRecs
  
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxOptSrchPers.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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
  
  ThisFile = "\NCTaxOptSrchReal.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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
  
  ThisFile = "\NCTaxOptSrchCust.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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

Public Sub ProcessNCTownships()
  Dim TownRec As TownshipType
  Dim THandle As Integer
  Dim NumOfTRecs As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  NCOpenTownshipFile THandle, NumOfTRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxTownships.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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

Public Sub ProcessNCSystemSetup()
  Dim SysRec As NCTaxMasterType
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
  ThisFile = "\NCTaxSystemSetup.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenTaxSetUpFile SHandle

  Get SHandle, 1, SysRec
  Print #RptHandle, QPTrim$(SysRec.Name);
  Print #RptHandle, B & QPTrim$(SysRec.Add1);
  Print #RptHandle, B & QPTrim$(SysRec.Add2);
  Print #RptHandle, B & QPTrim$(SysRec.City);
  Print #RptHandle, B & QPTrim$(SysRec.Zip);
  Print #RptHandle, B & QPTrim$(SysRec.TaxSt);
  Print #RptHandle, B & CStr(SysRec.TaxForm);
  Print #RptHandle, B & CStr(SysRec.TaxYear);
  Print #RptHandle, B & CStr(SysRec.LateForm);
  Print #RptHandle, B & SysRec.WarnInt;
  Print #RptHandle, B & QPTrim$(Using$(M, SysRec.MinBill));
  Print #RptHandle, B & QPTrim$(SysRec.AcctgMethod);
  Print #RptHandle, B & CStr(SysRec.MinTxOpt);
  Print #RptHandle, B & QPTrim$(SysRec.TownState);
  Print #RptHandle, B & CStr(SysRec.CurrYrInt);
  Print #RptHandle, B & CStr(SysRec.PastYrInt);
'  Print #RptHandle, B & Using$(P, SysRec.PenPct);
'  Print #RptHandle, B & CStr(SysRec.PenIdx);
  Print #RptHandle, B & SysRec.CntrlDepYN;
'  Print #RptHandle, B & SysRec.PriorYrMltRevYN;
  Print #RptHandle, B & QPTrim$(SysRec.OverPayGLNum);
  Print #RptHandle, B & SysRec.PenPrncTaxYN;
  Print #RptHandle, B & SysRec.PenIntYN;
  Print #RptHandle, B & SysRec.PenAdvYN;
  Print #RptHandle, B & SysRec.PenLateLstYN;
  Print #RptHandle, B & SysRec.PenOpt1YN;
  Print #RptHandle, B & SysRec.PenOpt2YN;
  Print #RptHandle, B & SysRec.PenOpt3YN;
  Print #RptHandle, B & SysRec.IntPrncTaxYN;
  Print #RptHandle, B & SysRec.IntIntYN;
  Print #RptHandle, B & SysRec.IntAdvYN;
  Print #RptHandle, B & SysRec.IntLateLstYN;
  Print #RptHandle, B & SysRec.IntOpt1YN;
  Print #RptHandle, B & SysRec.IntOpt2YN;
  Print #RptHandle, B & SysRec.IntOpt3YN;
  Print #RptHandle, B & QPTrim$(SysRec.OptRev1);
  Print #RptHandle, B & QPTrim$(SysRec.OptRev2);
  Print #RptHandle, B & QPTrim$(SysRec.OptRev3);
  Print #RptHandle, B & MakeRegDate(SysRec.DiscXDate);
  Print #RptHandle, B & Using(P, SysRec.DisPct);
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
  Print #RptHandle, B & QPTrim$(SysRec.OptSrchPers);
  Print #RptHandle, B & SysRec.AutoFillSrvAdd;
  Print #RptHandle, B
  Close

End Sub

Public Sub ProcessNCMessages()
  Dim MessRec As TaxMessRecType
  Dim MHandle As Integer
  Dim NumOfMRecs As Integer
  Dim x As Long, Y As Integer
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  
  NCOpenTaxMessage MHandle, NumOfMRecs
  StartPath = App.Path
  B = "|"
  ThisFile = "\NCTaxMessages.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
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

Public Sub ProcessNCTaxBill()
  Dim LtrRec As NCTaxBillType
  Dim LHandle As Integer
  Dim NumOfLRecs As Long
  Dim x As Long
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
  ThisFile = "\NCTaxBill.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenTaxBillFile LHandle, NumOfLRecs
  For x = 1 To NumOfLRecs
    Get LHandle, x, LtrRec
    Print #RptHandle, CStr(LtrRec.CustRec);
    Print #RptHandle, B & QPTrim$(LtrRec.CustName);
    Print #RptHandle, B & QPTrim$(LtrRec.CustAdd1);
    Print #RptHandle, B & QPTrim$(LtrRec.CustAdd2);
    Print #RptHandle, B & QPTrim$(LtrRec.CustAdd3);
    Print #RptHandle, B & QPTrim$(LtrRec.CustZip);
    Print #RptHandle, B & QPTrim$(LtrRec.RDesc1);
    Print #RptHandle, B & QPTrim$(LtrRec.RDesc2);
    Print #RptHandle, B & QPTrim$(LtrRec.RealPin);
    Print #RptHandle, B & QPTrim$(LtrRec.PersPin);
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.RealValue));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.PersValue));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.ExptValue));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.RealTaxDue));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.PersTaxDue));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.LateTaxDue));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.TotalBillDue));
    Print #RptHandle, B & QPTrim$(Using$(N, LtrRec.BillNumber));
    Print #RptHandle, B & QPTrim$(Using$(N, LtrRec.TaxYear));
    Print #RptHandle, B & QPTrim$(Using$(N, LtrRec.BillPrinted));
    Print #RptHandle, B & CStr(LtrRec.RealPropRecord);
    Print #RptHandle, B & CStr(LtrRec.PersPropRecord);
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.PriorYrBalance));
    Print #RptHandle, B & QPTrim$(Using$(P, LtrRec.RealTaxRate));
    Print #RptHandle, B & QPTrim$(Using$(P, LtrRec.PersTaxRate));
    Print #RptHandle, B & CStr(LtrRec.CustPin);
    Print #RptHandle, B & QPTrim$(LtrRec.TownShip);
    Print #RptHandle, B & QPTrim$(LtrRec.MORTCODE);
    Print #RptHandle, B & QPTrim$(LtrRec.LotOrAcre);
    Print #RptHandle, B & QPTrim$(LtrRec.LASize);
    Print #RptHandle, B & CStr(LtrRec.MortRec);
    Print #RptHandle, B & CStr(LtrRec.CarShore);
    Print #RptHandle, B & QPTrim$(LtrRec.RDesc3);
    Print #RptHandle, B & CStr(LtrRec.InternalPin);
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.OptRevTax1));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.OptRevTax2));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.OptRevTax3));
    Print #RptHandle, B & QPTrim$(Using$(M, LtrRec.OverPayAmt));
    Print #RptHandle, B & LtrRec.SetDscvry2No;
    Print #RptHandle, B
  Next x
  Close
  
End Sub
Public Sub ProcessNCLateLetter()
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
  ThisFile = "\NCTaxLateLetter.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenLateLtrFile LHandle
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

Public Sub ProcessNCRateTables()
  Dim RevRec As NCOptRevRateTablesType
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
  ThisFile = "\NCTaxRateTables.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenTaxRateTables RHandle, NumOfRRecs
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
    Print #RptHandle, B
  Next x
  Close
End Sub

Public Sub ProcessNCGLPay()
  Dim RGLRec As NCTaxAcctsType
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
  ThisFile = "\NCTaxGLRealPay.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenTaxGLInterPay RHandle
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

Public Sub ProcessNCGLBill()
  Dim RGLRec As NCTaxAcctsType
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
  ThisFile = "\NCTaxGLRealBill.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenTaxGLInterBill RHandle
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
Public Sub ProcessNCTransHist()
  Dim TransRec As NCTaxTransactionType
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
  
  ThisFile = "\NCTaxTaxTrans.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  B = "|"
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  NCOpenTaxTransFile THandle, NumOfTRecs
  
  FrmShowPctComp.Label1 = "Tax Transaction History Export"
  FrmShowPctComp.Show
  DoEvents

  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
'    If x = 1650 Then Stop
    Print #RptHandle, MakeRegDate(TransRec.TransDate);
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
    Print #RptHandle, B & TransRec.Posted2GL;
    Print #RptHandle, B & CStr(TransRec.CustomerRec);
'    Print #RptHandle, B & CStr(TransRec.LastTrans);
    Print #RptHandle, B & CStr(TransRec.BelongTo);
    Print #RptHandle, B & CheckForBad(TransRec.DMVSubmitted);
    Print #RptHandle, B & CStr(TransRec.DMVBatch);
    Print #RptHandle, B & CStr(TransRec.Altered);
    Print #RptHandle, B & CheckForBad(TransRec.FromPrePay);
    Print #RptHandle, B & QPTrim$(TransRec.PersPin);
    Print #RptHandle, B & QPTrim$(TransRec.RealPin);
    Print #RptHandle, B & CStr(TransRec.CustPin);
    Print #RptHandle, B & CStr(TransRec.InternalPin);
    Print #RptHandle, B & MakeRegDate(TransRec.DiscXDate);
    Print #RptHandle, B & Using$(P, TransRec.DiscAmt);
    Print #RptHandle, B & CStr(TransRec.OperNum);
    Print #RptHandle, B & QPTrim$(TransRec.CntyPara);
    Print #RptHandle, B & QPTrim$(TransRec.CyclPara);
    Print #RptHandle, B & QPTrim$(TransRec.TShpPara);
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

Public Sub ProcessNCOptRevRateTables()
  Dim RateRec As NCOptRevRateTablesType
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
  
  ThisFile = "\NCTaxOptRateTbls.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  B = "|"
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  NCOpenTaxRateTables THandle, NumOfTRecs
  
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

Public Sub ProcessNCLaserStandard()
  Dim LtrRec As NCTxBill1DefaultsType
  Dim LHandle As Integer
  Dim x As Long
  Dim RptHandle As Integer
  Dim RptName As String
  Dim B As String
  Dim ThisFile As String
  Dim NumOfRecs As Integer
  
  StartPath = App.Path

  B = "|"
  ThisFile = "\NCTaxLaserStandard.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  NCOpenTxBill1File LHandle
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
  
End Sub

Public Sub ProcessNCBalance()
  Dim TransRec As NCTaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
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
  
  ThisFile = "NCTaxBalance.txt"
  If DirExists(StartPath + "\NCTAXConvToTxt") Then
    If Exist(StartPath + "\NCTAXConvToTxt\" + ThisFile) Then
      KillFile (StartPath + "\NCTAXConvToTxt\" + ThisFile)
    End If
  Else
    MkDir StartPath + "\NCTAXConvToTxt"
  End If
  
  B = "|"
  RptName$ = StartPath + "\NCTAXConvToTxt\" + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  NCOpenTaxTransFile THandle, NumOfTRecs
  NCOpenTaxCustFile CHandle, NumOfCRecs
  FrmShowPctComp.Label1 = "Tax Balance Export"
  FrmShowPctComp.Show
  DoEvents
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    If TaxCust.Deleted = -1 Then GoTo NoBalance
    Balance = NCGetCustBalance(x, -1)
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
          Balance# = OldRound#(Balance# - (TransRec.Revenue.RevOpt1Pd + TransRec.Revenue.RevOpt2Pd + TransRec.Revenue.RevOpt3Pd + TransRec.DiscAmt))
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
    Print #RptHandle, B & "";
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & QPTrim$(Using$(M, 0));
    Print #RptHandle, B & "Over Payment";
    Print #RptHandle, B & CStr(TaxCust.Acct);
    Print #RptHandle, B & "";
    Print #RptHandle, B & "";
    Print #RptHandle, B & Using$(P, 0);
    Print #RptHandle, B & CStr(0);
    Print #RptHandle, B & "0";
    Print #RptHandle, B & "";
    Print #RptHandle, B
  Return
  
SendThis:
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
    Print #RptHandle, B & QPTrim$(TransRec.RealPin);
    Print #RptHandle, B & Using$(P, TransRec.DiscAmt);
    Print #RptHandle, B & CStr(TransRec.OperNum);
    Print #RptHandle, B & ParseBillNum(TransRec.Description);
    Print #RptHandle, B & TransRec.Posted2GL;
    Print #RptHandle, B
  Return

End Sub
