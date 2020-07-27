Attribute VB_Name = "PRSaveCommon"
'This mod was created to save large amounts of data
'requiring error checking in detail or with more than one form
Public Sub SaveFedTax(frm1 As Form, frm2 As Form)
   Dim FedTaxHandle As Integer, x As Integer
   Dim FedTaxFileRec As FederalTaxRecType
   
   OpenFedTaxFile FedTaxHandle
   'if a field is left empty then save the field as zero
   'to keep the program from crashing
   If QPTrim$(frm1.fptxtEmployeeSSPer.Text) = "" Then
     FedTaxFileRec.FTSEMPSS = 0
   Else
     FedTaxFileRec.FTSEMPSS = QPTrim$(frm1.fptxtEmployeeSSPer.Text)
   End If
   If QPTrim$(frm1.fptxtEmployerSSPer.Text) = "" Then
     FedTaxFileRec.FTSEMRSS = 0
   Else
     FedTaxFileRec.FTSEMRSS = QPTrim$(frm1.fptxtEmployerSSPer.Text)
   End If
   If QPTrim$(frm1.fptxtSocSecMaxWages.Text) = "" Then
     FedTaxFileRec.FTSSSMW = 0
   Else
     FedTaxFileRec.FTSSSMW = QPTrim$(frm1.fptxtSocSecMaxWages.Text)
   End If
   If QPTrim$(frm1.fptxtEmployeeMedPer.Text) = "" Then
     FedTaxFileRec.FTSEMPM = 0
   Else
     FedTaxFileRec.FTSEMPM = QPTrim$(frm1.fptxtEmployeeMedPer.Text)
   End If
   If QPTrim$(frm1.fptxtEmployerMedPer.Text) = "" Then
     FedTaxFileRec.FTSEMRM = 0
   Else
     FedTaxFileRec.FTSEMRM = QPTrim$(frm1.fptxtEmployerMedPer.Text)
   End If
   If QPTrim$(frm1.fptxtMedMaxWages.Text) = "" Then
     FedTaxFileRec.FTSMMW = 0
   Else
     FedTaxFileRec.FTSMMW = QPTrim$(frm1.fptxtMedMaxWages.Text)
   End If
   If QPTrim$(frm1.fptxtStdDed.Text) = "" Then
     FedTaxFileRec.FTSSDAA = 0
   Else
     FedTaxFileRec.FTSSDAA = QPTrim$(frm1.fptxtStdDed.Text)
   End If
   
   For x = 1 To 10
     If QPTrim$(frm1.fpText1(x).Text) = "" Then
       FedTaxFileRec.FTS(1, x) = 0
     Else
       FedTaxFileRec.FTS(1, x) = QPTrim$(frm1.fpText1(x).Text)
     End If
     If QPTrim$(frm1.fpText2(x).Text) = "" Then
       FedTaxFileRec.FTS(2, x) = 0
     Else
       FedTaxFileRec.FTS(2, x) = QPTrim$(frm1.fpText2(x).Text)
     End If
     If QPTrim$(frm1.fpText3(x).Text) = "" Then
       FedTaxFileRec.FTS(3, x) = 0
     Else
       FedTaxFileRec.FTS(3, x) = QPTrim$(frm1.fpText3(x).Text)
     End If
   Next x
   
   If QPTrim$(frm2.fptxtEmployeeSSPer.Text) = "" Then
     FedTaxFileRec.FTMEMPSS = 0
   Else
     FedTaxFileRec.FTMEMPSS = QPTrim$(frm2.fptxtEmployeeSSPer.Text)
   End If
   If QPTrim$(frm2.fptxtEmployerSSPer.Text) = "" Then
     FedTaxFileRec.FTMEMRSS = 0
   Else
     FedTaxFileRec.FTMEMRSS = QPTrim$(frm2.fptxtEmployerSSPer.Text)
   End If
   If QPTrim$(frm2.fptxtSocSecMaxWages.Text) = "" Then
     FedTaxFileRec.FTMSSMW = 0
   Else
     FedTaxFileRec.FTMSSMW = QPTrim$(frm2.fptxtSocSecMaxWages.Text)
   End If
   If QPTrim$(frm2.fptxtEmployeeMedPer.Text) = "" Then
     FedTaxFileRec.FTMEMPM = 0
   Else
     FedTaxFileRec.FTMEMPM = QPTrim$(frm2.fptxtEmployeeMedPer.Text)
   End If
   If QPTrim$(frm2.fptxtEmployerMedPer.Text) = "" Then
     FedTaxFileRec.FTMEMRM = 0
   Else
     FedTaxFileRec.FTMEMRM = QPTrim$(frm2.fptxtEmployerMedPer.Text)
   End If
   If QPTrim$(frm2.fptxtMedMaxWages.Text) = "" Then
     FedTaxFileRec.FTMMMW = 0
   Else
     FedTaxFileRec.FTMMMW = QPTrim$(frm2.fptxtMedMaxWages.Text)
   End If
   If QPTrim$(frm2.fptxtStdDed.Text) = "" Then
     FedTaxFileRec.FTMSDAA = 0
   Else
     FedTaxFileRec.FTMSDAA = QPTrim$(frm2.fptxtStdDed.Text)
   End If
   
   For x = 1 To 10
     If QPTrim$(frm2.fpText1(x).Text) = "" Then
       FedTaxFileRec.FTM(1, x) = 0
     Else
       FedTaxFileRec.FTM(1, x) = QPTrim$(frm2.fpText1(x).Text)
     End If
     If QPTrim$(frm2.fpText2(x).Text) = "" Then
       FedTaxFileRec.FTM(2, x) = 0
     Else
       FedTaxFileRec.FTM(2, x) = QPTrim$(frm2.fpText2(x).Text)
     End If
     If QPTrim$(frm2.fpText3(x).Text) = "" Then
       FedTaxFileRec.FTM(3, x) = 0
     Else
       FedTaxFileRec.FTM(3, x) = QPTrim$(frm2.fpText3(x).Text)
     End If
   Next x
   FedTaxFileRec.FTMEMRM = FedTaxFileRec.FTMEMRM
   Put FedTaxHandle, 1, FedTaxFileRec
   
   Close FedTaxHandle
   
   MsgBox "Your Information has been saved.", vbOKOnly
   frmControlFileMaint.Show
   DoEvents
   Unload frm1
   DoEvents
   Unload frm2
End Sub
Public Sub checkExitEmpFedTax(frm1 As Form, frm2 As Form)
   Dim DoWhatFlag As SaveChangeOptions1
   Dim save As Integer, review As Integer, abandon As Integer
   Dim FedTaxHandle As Integer, x As Integer
   Dim FedTaxFileRec As FederalTaxRecType
   Dim changeFlag As Boolean, FileLen As Long
   
   changeFlag = False
   OpenFedTaxFile FedTaxHandle
   FileLen = LOF(FedTaxHandle) / Len(FedTaxFileRec)
   If FileLen > 0 Then
     Get FedTaxHandle, 1, FedTaxFileRec
   End If
   Close FedTaxHandle
   'First save
   If FileLen = 0 Then
     If Val(frm1.fptxtEmployeeSSPer.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm1.fptxtEmployerSSPer.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm1.fptxtSocSecMaxWages.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm1.fptxtEmployeeMedPer.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm1.fptxtEmployerMedPer.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm1.fptxtMedMaxWages.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm1.fptxtStdDed.Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 10
       If Val(frm1.fpText1(x).Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fpText1(x).SetFocus
       End If
       If Val(frm1.fpText2(x).Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fpText2(x).SetFocus
       End If
       If Val(frm1.fpText3(x).Text) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
     
     If Val(frm2.fptxtEmployeeSSPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtEmployerSSPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtSocSecMaxWages.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtEmployeeMedPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtEmployerMedPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtMedMaxWages.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtStdDed.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 10
       If Val(frm2.fpText1(x).Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fpText1(x).SetFocus
       End If
       If Val(frm2.fpText2(x).Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fpText2(x).SetFocus
       End If
       If Val(frm2.fpText3(x).Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
      
   ElseIf FileLen > 0 Then
     'update
     If QPTrim$(frm1.fptxtEmployeeSSPer.Text) = "" Then frm1.fptxtEmployeeSSPer.Text = "0"
     If QPTrim$(frm1.fptxtEmployeeSSPer.Text) <> FedTaxFileRec.FTSEMPSS Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtEmployerSSPer.Text) = "" Then frm1.fptxtEmployerSSPer.Text = "0"
     If QPTrim$(frm1.fptxtEmployerSSPer.Text) <> FedTaxFileRec.FTSEMRSS Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtSocSecMaxWages.Text) = "" Then frm1.fptxtSocSecMaxWages.Text = "0"
     If QPTrim$(frm1.fptxtSocSecMaxWages.Text) <> FedTaxFileRec.FTSSSMW Then
         changeFlag = True
         frm1.Show
         frm1.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtEmployeeMedPer.Text) = "" Then frm1.fptxtEmployeeMedPer.Text = "0"
     If QPTrim$(frm1.fptxtEmployeeMedPer.Text) <> FedTaxFileRec.FTSEMPM Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtEmployerMedPer.Text) = "" Then frm1.fptxtEmployerMedPer.Text = "0"
     If QPTrim$(frm1.fptxtEmployerMedPer.Text) <> FedTaxFileRec.FTSEMRM Then
         changeFlag = True
         frm1.Show
         frm1.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtMedMaxWages.Text) = "" Then frm1.fptxtMedMaxWages.Text = "0"
     If QPTrim$(frm1.fptxtMedMaxWages.Text) <> FedTaxFileRec.FTSMMW Then
         changeFlag = True
         frm1.Show
         frm1.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtStdDed.Text) = "" Then frm1.fptxtStdDed.Text = "0"
     If QPTrim$(frm1.fptxtStdDed.Text) <> FedTaxFileRec.FTSSDAA Then
         changeFlag = True
         frm1.Show
         frm1.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 10
       If QPTrim$(frm1.fpText1(x).Text) = "" Then frm1.fpText1(x).Text = "0"
       If QPTrim$(frm1.fpText1(x).Text) <> FedTaxFileRec.FTS(1, x) Then
         changeFlag = True
         frm1.Show
         frm1.fpText1(x).SetFocus
       End If
       If QPTrim$(frm1.fpText2(x).Text) = "" Then frm1.fpText2(x).Text = "0"
       If QPTrim$(frm1.fpText2(x).Text) <> FedTaxFileRec.FTS(2, x) Then
         changeFlag = True
         frm1.Show
         frm1.fpText2(x).SetFocus
       End If
       If QPTrim$(frm1.fpText3(x).Text) = "" Then frm1.fpText3(x).Text = "0"
       If QPTrim$(frm1.fpText3(x).Text) <> FedTaxFileRec.FTS(3, x) Then
         changeFlag = True
         frm1.Show
         frm1.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
     
     If QPTrim$(frm2.fptxtEmployeeSSPer.Text) = "" Then frm2.fptxtEmployeeSSPer.Text = "0"
     If QPTrim$(frm2.fptxtEmployeeSSPer.Text) <> FedTaxFileRec.FTMEMPSS Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtEmployerSSPer.Text) = "" Then frm2.fptxtEmployerSSPer.Text = "0"
     If QPTrim$(frm2.fptxtEmployerSSPer.Text) <> FedTaxFileRec.FTMEMRSS Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtSocSecMaxWages.Text) = "" Then frm2.fptxtSocSecMaxWages.Text = "0"
     If QPTrim$(frm2.fptxtSocSecMaxWages.Text) <> FedTaxFileRec.FTMSSMW Then
         changeFlag = True
         frm2.Show
         frm2.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtEmployeeMedPer.Text) = "" Then frm2.fptxtEmployeeMedPer.Text = "0"
     If QPTrim$(frm2.fptxtEmployeeMedPer.Text) <> FedTaxFileRec.FTMEMPM Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtEmployerMedPer.Text) = "" Then frm2.fptxtEmployerMedPer.Text = "0"
     If QPTrim$(frm2.fptxtEmployerMedPer.Text) <> FedTaxFileRec.FTMEMRM Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtMedMaxWages.Text) = "" Then frm2.fptxtMedMaxWages.Text = "0"
     If QPTrim$(frm2.fptxtMedMaxWages.Text) <> FedTaxFileRec.FTMMMW Then
         changeFlag = True
         frm2.Show
         frm2.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtStdDed.Text) = "" Then frm2.fptxtStdDed.Text = "0"
     If QPTrim$(frm2.fptxtStdDed.Text) <> FedTaxFileRec.FTMSDAA Then
         changeFlag = True
         frm2.Show
         frm2.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 10
       If QPTrim$(frm2.fpText1(x).Text) = "" Then frm2.fpText1(x).Text = "0"
       If QPTrim$(frm2.fpText1(x).Text) <> FedTaxFileRec.FTM(1, x) Then
         changeFlag = True
         frm2.Show
         frm2.fpText1(x).SetFocus
       End If
       If QPTrim$(frm2.fpText2(x).Text) = "" Then frm2.fpText2(x).Text = "0"
       If QPTrim$(frm2.fpText2(x).Text) <> FedTaxFileRec.FTM(2, x) Then
         changeFlag = True
         frm2.Show
         frm2.fpText2(x).SetFocus
       End If
       If QPTrim$(frm2.fpText3(x).Text) = "" Then frm2.fpText3(x).Text = "0"
       If QPTrim$(frm2.fpText3(x).Text) <> FedTaxFileRec.FTM(3, x) Then
         changeFlag = True
         frm2.Show
         frm2.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
   End If
     
changeFound:
    Close
    If changeFlag = False Then 'no changes detected
       frmControlFileMaint.Show
       DoEvents
       Unload frm1
       DoEvents
       Unload frm2
       GoTo endClick
     'if a change was made then bring up a warning window that forces
     'the user to decide whether to save, review or abandon changes
    Else
       DoWhatFlag = PromptSaveChanges(frm1)
       Select Case DoWhatFlag
       Case SaveChangeOptions1.scoSaveChanges 'save changes
         Call SaveFedTax(frm1, frm2)
       Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
       Case SaveChangeOptions1.scoAbandonChanges 'abandon
         frmControlFileMaint.Show
         DoEvents
         Unload frm1
         DoEvents
         Unload frm2
       Case Else:
          'Do nothing because we don't know about any options except
          'save, review or abandon...used as a placeholder for adding
          'other options at a later date
       End Select
         
    End If
endClick:
     
   
End Sub

Public Sub SaveStaTax(frm1 As Form, frm2 As Form, frm3 As Form)
   Dim StaTaxHandle As Integer, x As Integer
   Dim StaTaxFileRec As StateTaxRecType
   
   'if a field is left empty then save the field as zero
   'to keep the program from crashing
   If QPTrim$(frm1.fptxtEmployeeSSPer.Text) = "" Then
     StaTaxFileRec.TAX101 = 0
   Else
     StaTaxFileRec.TAX101 = CDbl(frm1.fptxtEmployeeSSPer.Text)
   End If
   If QPTrim$(frm1.fptxtEmployerSSPer.Text) = "" Then
     StaTaxFileRec.TAX102 = 0
   Else
     StaTaxFileRec.TAX102 = CDbl(frm1.fptxtEmployerSSPer.Text)
   End If
   If QPTrim$(frm1.fptxtSocSecMaxWages.Text) = "" Then
     StaTaxFileRec.TAX103 = 0
   Else
     StaTaxFileRec.TAX103 = CDbl(frm1.fptxtSocSecMaxWages.Text)
   End If
   If QPTrim$(frm1.fptxtEmployeeMedPer.Text) = "" Then
     StaTaxFileRec.TAX104 = 0
   Else
     StaTaxFileRec.TAX104 = CDbl(frm1.fptxtEmployeeMedPer.Text)
   End If
   If QPTrim$(frm1.fptxtEmployerMedPer.Text) = "" Then
     StaTaxFileRec.TAX105 = 0
   Else
     StaTaxFileRec.TAX105 = CDbl(frm1.fptxtEmployerMedPer.Text)
   End If
   If QPTrim$(frm1.fptxtMedMaxWages.Text) = "" Then
     StaTaxFileRec.TAX106 = 0
   Else
     StaTaxFileRec.TAX106 = CDbl(frm1.fptxtMedMaxWages.Text)
   End If
   If QPTrim$(frm1.fptxtStdDed.Text) = "" Then
     StaTaxFileRec.TAX107 = 0
   Else
     StaTaxFileRec.TAX107 = CDbl(frm1.fptxtStdDed.Text)
   End If
   
   For x = 1 To 12
      If QPTrim$(frm1.fpText1(x)) <> "" Then
        StaTaxFileRec.STS(1, x) = CDbl(frm1.fpText1(x))
      Else
        StaTaxFileRec.STS(1, x) = 0
      End If
      If QPTrim$(frm1.fpText2(x)) <> "" Then
        StaTaxFileRec.STS(2, x) = CDbl(frm1.fpText2(x))
      Else
        StaTaxFileRec.STS(2, x) = 0
      End If
      If QPTrim$(frm1.fpText3(x)) <> "" Then
        StaTaxFileRec.STS(3, x) = CDbl(frm1.fpText3(x))
      Else
        StaTaxFileRec.STS(3, x) = 0
      End If
   Next x
   
   If QPTrim$(frm2.fptxtEmployeeSSPer.Text) = "" Then
     StaTaxFileRec.TAX201 = 0
   Else
     StaTaxFileRec.TAX201 = CDbl(frm2.fptxtEmployeeSSPer.Text)
   End If
   If QPTrim$(frm2.fptxtEmployerSSPer.Text) = "" Then
     StaTaxFileRec.TAX202 = 0
   Else
     StaTaxFileRec.TAX202 = CDbl(frm2.fptxtEmployerSSPer.Text)
   End If
   If QPTrim$(frm2.fptxtSocSecMaxWages.Text) = "" Then
     StaTaxFileRec.TAX203 = 0
   Else
     StaTaxFileRec.TAX203 = CDbl(frm2.fptxtSocSecMaxWages.Text)
   End If
   If QPTrim$(frm2.fptxtEmployeeMedPer.Text) = "" Then
     StaTaxFileRec.TAX204 = 0
   Else
     StaTaxFileRec.TAX204 = CDbl(frm2.fptxtEmployeeMedPer.Text)
   End If
   If QPTrim$(frm2.fptxtEmployerMedPer.Text) = "" Then
     StaTaxFileRec.TAX205 = 0
   Else
     StaTaxFileRec.TAX205 = CDbl(frm2.fptxtEmployerMedPer.Text)
   End If
   If QPTrim$(frm2.fptxtMedMaxWages.Text) = "" Then
     StaTaxFileRec.TAX206 = 0
   Else
     StaTaxFileRec.TAX206 = CDbl(frm2.fptxtMedMaxWages.Text)
   End If
   If QPTrim$(frm2.fptxtStdDed.Text) = "" Then
     StaTaxFileRec.TAX207 = 0
   Else
     StaTaxFileRec.TAX207 = CDbl(frm2.fptxtStdDed.Text)
   End If
   
   For x = 1 To 12
      If QPTrim$(frm2.fpText1(x)) <> "" Then
        StaTaxFileRec.STM(1, x) = CDbl(frm2.fpText1(x))
      Else
        StaTaxFileRec.STM(1, x) = 0
      End If
      If QPTrim$(frm2.fpText2(x)) <> "" Then
        StaTaxFileRec.STM(2, x) = CDbl(frm2.fpText2(x))
      Else
        StaTaxFileRec.STM(2, x) = 0
      End If
      If QPTrim$(frm2.fpText3(x)) <> "" Then
        StaTaxFileRec.STM(3, x) = CDbl(frm2.fpText3(x))
      Else
        StaTaxFileRec.STM(3, x) = 0
      End If
   Next x
   If QPTrim$(frm3.fptxtEmployeeSSPer.Text) = "" Then
     StaTaxFileRec.TAX301 = 0
   Else
     StaTaxFileRec.TAX301 = CDbl(frm3.fptxtEmployeeSSPer.Text)
   End If
   If QPTrim$(frm3.fptxtEmployerSSPer.Text) = "" Then
     StaTaxFileRec.TAX302 = 0
   Else
     StaTaxFileRec.TAX302 = CDbl(frm3.fptxtEmployerSSPer.Text)
   End If
   If QPTrim$(frm3.fptxtSocSecMaxWages.Text) = "" Then
     StaTaxFileRec.TAX303 = 0
   Else
     StaTaxFileRec.TAX303 = CDbl(frm3.fptxtSocSecMaxWages.Text)
   End If
   If QPTrim$(frm3.fptxtEmployeeMedPer.Text) = "" Then
     StaTaxFileRec.TAX304 = 0
   Else
     StaTaxFileRec.TAX304 = CDbl(frm3.fptxtEmployeeMedPer.Text)
   End If
   If QPTrim$(frm3.fptxtEmployerMedPer.Text) = "" Then
     StaTaxFileRec.TAX305 = 0
   Else
     StaTaxFileRec.TAX305 = CDbl(frm3.fptxtEmployerMedPer.Text)
   End If
   If QPTrim$(frm3.fptxtMedMaxWages.Text) = "" Then
     StaTaxFileRec.TAX306 = 0
   Else
     StaTaxFileRec.TAX306 = CDbl(frm3.fptxtMedMaxWages.Text)
   End If
   If QPTrim$(frm3.fptxtStdDed.Text) = "" Then
     StaTaxFileRec.TAX307 = 0
   Else
     StaTaxFileRec.TAX307 = CDbl(frm3.fptxtStdDed.Text)
   End If
   'if the text field is not empty then save the value but if
   'it is empty then save a 0
   For x = 1 To 12
      If QPTrim$(frm3.fpText1(x)) <> "" Then
        StaTaxFileRec.STH(1, x) = CDbl(frm3.fpText1(x))
      Else
        StaTaxFileRec.STH(1, x) = 0
      End If
      If QPTrim$(frm3.fpText2(x)) <> "" Then
        StaTaxFileRec.STH(2, x) = CDbl(frm3.fpText2(x))
      Else
        StaTaxFileRec.STH(2, x) = 0
      End If
      If QPTrim$(frm3.fpText3(x)) <> "" Then
        StaTaxFileRec.STH(3, x) = CDbl(frm3.fpText3(x))
      Else
        StaTaxFileRec.STH(3, x) = 0
      End If
   Next x
   OpenStateTaxFileName StaTaxHandle
   Put StaTaxHandle, 1, StaTaxFileRec
   Close StaTaxHandle
  
   MsgBox "Your Information has been saved.", vbOKOnly
   frmControlFileMaint.Show
   DoEvents
   Unload frm1
   DoEvents
   Unload frm2
   DoEvents
   Unload frm3
   MainLog ("State tax data saved.")
End Sub

Public Sub checkExitEmpStaTax(frm1 As Form, frm2 As Form, frm3 As Form)
   Dim DoWhatFlag As SaveChangeOptions1
   Dim save As Integer, review As Integer, abandon As Integer
   Dim StaTaxHandle As Integer, x As Integer
   Dim StaTaxFileRec As StateTaxRecType
   Dim changeFlag As Boolean, FileLen As Long
   
   OpenStateTaxFileName StaTaxHandle
   Get StaTaxHandle, 1, StaTaxFileRec
   changeFlag = False
   FileLen = LOF(StaTaxHandle) / Len(StaTaxFileRec)
   Close StaTaxHandle
   If FileLen = 0 Then
     If Val(frm1.fptxtEmployeeSSPer.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployeeSSPer.SetFocus
       GoTo changeFound
     End If
     If Val(frm1.fptxtEmployerSSPer.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployerSSPer.SetFocus
       GoTo changeFound
     End If
     If Val(frm1.fptxtSocSecMaxWages.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtSocSecMaxWages.SetFocus
       GoTo changeFound
     End If
     If Val(frm1.fptxtEmployeeMedPer.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployeeMedPer.SetFocus
       GoTo changeFound
     End If
     If Val(frm1.fptxtEmployerMedPer.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployerMedPer.SetFocus
       GoTo changeFound
     End If
     If Val(frm1.fptxtMedMaxWages.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtMedMaxWages.SetFocus
       GoTo changeFound
     End If
     If Val(frm1.fptxtStdDed.Text) <> 0 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtStdDed.SetFocus
       GoTo changeFound
     End If
      
     For x = 1 To 12
       If Val(frm1.fpText1(x)) <> 0 Then '"$.00"
         changeFlag = True
         frm1.Show
         frm1.fpText1(x).SetFocus
       End If
       If Val(frm1.fpText2(x)) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fpText2(x).SetFocus
       End If
       If Val(frm1.fpText3(x)) <> 0 Then
         changeFlag = True
         frm1.Show
         frm1.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
     
     If Val(frm2.fptxtEmployeeSSPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtEmployerSSPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtSocSecMaxWages.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtEmployeeMedPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtEmployerMedPer.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtMedMaxWages.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm2.fptxtStdDed.Text) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 12
       If Val(frm2.fpText1(x)) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fpText1(x).SetFocus
       End If
       If Val(frm2.fpText2(x)) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fpText2(x).SetFocus
       End If
       If Val(frm2.fpText3(x)) <> 0 Then
         changeFlag = True
         frm2.Show
         frm2.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
     
     If Val(frm3.fptxtEmployeeSSPer.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm3.fptxtEmployerSSPer.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm3.fptxtSocSecMaxWages.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm3.fptxtEmployeeMedPer.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm3.fptxtEmployerMedPer.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If Val(frm3.fptxtMedMaxWages.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If Val(frm3.fptxtStdDed.Text) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
       
     For x = 1 To 12
       If Val(frm3.fpText1(x)) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fpText1(x).SetFocus
       End If
       If Val(frm3.fpText2(x)) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fpText2(x).SetFocus
       End If
       If Val(frm3.fpText3(x)) <> 0 Then
         changeFlag = True
         frm3.Show
         frm3.fpText2(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
      
   ElseIf FileLen <> 0 Then
     If QPTrim$(frm1.fptxtEmployeeSSPer.Text) = "" Then frm1.fptxtEmployeeSSPer.Text = 0
     If CDbl(frm1.fptxtEmployeeSSPer.Text) <> StaTaxFileRec.TAX101 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployeeSSPer.SetFocus
       GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtEmployerSSPer.Text) = "" Then frm1.fptxtEmployerSSPer.Text = 0
     If CDbl(frm1.fptxtEmployerSSPer.Text) <> StaTaxFileRec.TAX102 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployerSSPer.SetFocus
       GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtSocSecMaxWages.Text) = "" Then frm1.fptxtSocSecMaxWages.Text = 0
     If CDbl(frm1.fptxtSocSecMaxWages.Text) <> StaTaxFileRec.TAX103 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtSocSecMaxWages.SetFocus
       GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtEmployeeMedPer.Text) = "" Then frm1.fptxtEmployeeMedPer.Text = 0
     If CDbl(frm1.fptxtEmployeeMedPer.Text) <> StaTaxFileRec.TAX104 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployeeMedPer.SetFocus
       GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtEmployerMedPer.Text) = "" Then frm1.fptxtEmployerMedPer.Text = 0
     If CDbl(frm1.fptxtEmployerMedPer.Text) <> StaTaxFileRec.TAX105 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtEmployerMedPer.SetFocus
       GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtMedMaxWages.Text) = "" Then frm1.fptxtMedMaxWages.Text = 0
     If CDbl(frm1.fptxtMedMaxWages.Text) <> StaTaxFileRec.TAX106 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtMedMaxWages.SetFocus
       GoTo changeFound
     End If
     If QPTrim$(frm1.fptxtStdDed.Text) = "" Then frm1.fptxtStdDed.Text = 0
     If CDbl(frm1.fptxtStdDed.Text) <> StaTaxFileRec.TAX107 Then
       changeFlag = True
       frm1.Show
       frm1.fptxtStdDed.SetFocus
       GoTo changeFound
     End If
       
       For x = 1 To 12
       If QPTrim$(frm1.fpText1(x).Text) = "" Then frm1.fpText1(x).Text = 0
       If CDbl(frm1.fpText1(x).Text) <> StaTaxFileRec.STS(1, x) Then
         changeFlag = True
         frm1.Show
         frm1.fpText1(x).SetFocus
       End If
       If QPTrim$(frm1.fpText2(x).Text) = "" Then frm1.fpText2(x).Text = 0
       If CDbl(frm1.fpText2(x).Text) <> StaTaxFileRec.STS(2, x) Then
         changeFlag = True
         frm1.Show
         frm1.fpText2(x).SetFocus
       End If
       If QPTrim$(frm1.fpText3(x).Text) = "" Then frm1.fpText3(x).Text = 0
       If CDbl(frm1.fpText3(x).Text) <> StaTaxFileRec.STS(3, x) Then
         changeFlag = True
         frm1.Show
         frm1.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
     
     If QPTrim$(frm2.fptxtEmployeeSSPer.Text) = "" Then frm2.fptxtEmployeeSSPer.Text = 0
     If CDbl(frm2.fptxtEmployeeSSPer.Text) <> StaTaxFileRec.TAX201 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtEmployerSSPer.Text) = "" Then frm2.fptxtEmployerSSPer.Text = 0
     If CDbl(frm2.fptxtEmployerSSPer.Text) <> StaTaxFileRec.TAX202 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtSocSecMaxWages.Text) = "" Then frm2.fptxtSocSecMaxWages.Text = 0
     If CDbl(frm2.fptxtSocSecMaxWages.Text) <> StaTaxFileRec.TAX203 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtEmployeeMedPer.Text) = "" Then frm2.fptxtEmployeeMedPer.Text = 0
     If CDbl(frm2.fptxtEmployeeMedPer.Text) <> StaTaxFileRec.TAX204 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtEmployerMedPer.Text) = "" Then frm2.fptxtEmployerMedPer.Text = 0
     If CDbl(frm2.fptxtEmployerMedPer.Text) <> StaTaxFileRec.TAX205 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtMedMaxWages.Text) = "" Then frm2.fptxtMedMaxWages.Text = 0
     If CDbl(frm2.fptxtMedMaxWages.Text) <> StaTaxFileRec.TAX206 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm2.fptxtStdDed.Text) = "" Then frm2.fptxtStdDed.Text = 0
     If CDbl(frm2.fptxtStdDed.Text) <> StaTaxFileRec.TAX207 Then
         changeFlag = True
         frm2.Show
         frm2.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 12
       If QPTrim$(frm2.fpText1(x).Text) = "" Then frm2.fpText1(x).Text = 0
       If CDbl(frm2.fpText1(x).Text) <> StaTaxFileRec.STM(1, x) Then
         changeFlag = True
         frm2.Show
         frm2.fpText1(x).SetFocus
       End If
       If QPTrim$(frm2.fpText2(x).Text) = "" Then frm2.fpText2(x).Text = 0
       If CDbl(frm2.fpText2(x).Text) <> StaTaxFileRec.STM(2, x) Then
         changeFlag = True
         frm2.Show
         frm2.fpText2(x).SetFocus
       End If
       If QPTrim$(frm2.fpText3(x).Text) = "" Then frm2.fpText3(x).Text = 0
       If CDbl(frm2.fpText3(x).Text) <> StaTaxFileRec.STM(3, x) Then
         changeFlag = True
         frm2.Show
         frm2.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
     If changeFlag = True Then GoTo changeFound
     
     If QPTrim$(frm3.fptxtEmployeeSSPer.Text) = "" Then frm3.fptxtEmployeeSSPer.Text = 0
     If CDbl(frm3.fptxtEmployeeSSPer.Text) <> StaTaxFileRec.TAX301 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployeeSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm3.fptxtEmployerSSPer.Text) = "" Then frm3.fptxtEmployerSSPer.Text = 0
     If CDbl(frm3.fptxtEmployerSSPer.Text) <> StaTaxFileRec.TAX302 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployerSSPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm3.fptxtSocSecMaxWages.Text) = "" Then frm3.fptxtSocSecMaxWages.Text = 0
     If CDbl(frm3.fptxtSocSecMaxWages.Text) <> StaTaxFileRec.TAX303 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtSocSecMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm3.fptxtEmployeeMedPer.Text) = "" Then frm3.fptxtEmployeeMedPer.Text = 0
     If CDbl(frm3.fptxtEmployeeMedPer.Text) <> StaTaxFileRec.TAX304 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployeeMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm3.fptxtEmployerMedPer.Text) = "" Then frm3.fptxtEmployerMedPer.Text = 0
     If CDbl(frm3.fptxtEmployerMedPer.Text) <> StaTaxFileRec.TAX305 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtEmployerMedPer.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm3.fptxtMedMaxWages.Text) = "" Then frm3.fptxtMedMaxWages.Text = 0
     If CDbl(frm3.fptxtMedMaxWages.Text) <> StaTaxFileRec.TAX306 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtMedMaxWages.SetFocus
         GoTo changeFound
     End If
     If QPTrim$(frm3.fptxtStdDed.Text) = "" Then frm3.fptxtStdDed.Text = 0
     If CDbl(frm3.fptxtStdDed.Text) <> StaTaxFileRec.TAX307 Then
         changeFlag = True
         frm3.Show
         frm3.fptxtStdDed.SetFocus
         GoTo changeFound
     End If
      
     For x = 1 To 12
       If QPTrim$(frm3.fpText1(x).Text) = "" Then frm3.fpText1(x).Text = 0
       If CDbl(frm3.fpText1(x).Text) <> StaTaxFileRec.STH(1, x) Then
         changeFlag = True
         frm3.Show
         frm3.fpText1(x).SetFocus
       End If
       If QPTrim$(frm3.fpText2(x).Text) = "" Then frm3.fpText2(x).Text = 0
       If CDbl(frm3.fpText2(x).Text) <> StaTaxFileRec.STH(2, x) Then
         changeFlag = True
         frm3.Show
         frm3.fpText2(x).SetFocus
       End If
       If QPTrim$(frm3.fpText3(x).Text) = "" Then frm3.fpText3(x).Text = 0
       If CDbl(frm3.fpText3(x).Text) <> StaTaxFileRec.STH(3, x) Then
         changeFlag = True
         frm3.Show
         frm3.fpText3(x).SetFocus
       End If
       If changeFlag = True Then Exit For
     Next x
changeFound:
   End If
     If changeFlag = False Then 'no changes detected
        frmControlFileMaint.Show
        DoEvents
        Unload frm1
        DoEvents
        Unload frm2
        DoEvents
        Unload frm3
        DoEvents
        GoTo endClick
      'if a change was made then bring up a warning window that forces
      'the user to decide whether to save, review or abandon changes
     Else
        DoWhatFlag = PromptSaveChanges(frm1)
        Select Case DoWhatFlag
        Case SaveChangeOptions1.scoSaveChanges 'save changes
          Call SaveStaTax(frm1, frm2, frm3)
        Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
        Case SaveChangeOptions1.scoAbandonChanges 'abandon
          Unload frm1
          DoEvents
          Unload frm2
          DoEvents
          Unload frm3
          DoEvents
          frmControlFileMaint.Show
           'Do nothing because we don't know about any options except
           'save, review or abandon...used as a placeholder for adding
           'other options at a later date
        End Select
         
    End If
endClick:
    Close
End Sub

Public Sub SaveEmpInfo(newEmpFlag As Boolean, thisRecordNum As Integer, frm1 As Form)
   
   Dim EPN As Integer
   Dim LastPin As Integer, NextPin As Integer
   Dim EmpData2FileHandle As Integer, EmpData1FileHandle As Integer
   Dim EmpData2FileRec As EmpData2Type, EmpData1FileRec As EmpData1Type
   Dim EmpData3FileHandle As Integer, EmpData3FileRec As EmpData3Type
   Dim EmpRecNo As Long, saveHere As Long
   Dim DedCodeFileHandle As Integer
   Dim DedCodeFileRec As DedCodeRecType
   Dim THandle As Integer
   Dim EHandle As Integer
   Dim CHandle As Integer
   Dim EmpNumRec As EmpNumType
   Dim TransRec As TransRecType
   Dim CheckRec As PRCheckRecType
   
   Dim tempEmpNum As String, tempEmpSSN As String * 11
   Dim tempEmpLName As String, tempEmpFName As String
   Dim tempAddr1 As String, tempAddr2 As String
   Dim tempCity As String, tempState As String
   Dim tempZip As String, tempBDay As String
   Dim tempGender As String, tempRetNo As String
   Dim tempRace As String
   Dim tempRetType As String, tempBankDraft As String
   Dim tempBankNum As String, tempPreNoted As String
   Dim tempBankName As String, tempBankLoc As String
   Dim tempBankTransNo As String, tempJobTitle As String
   Dim tempWCCode As String, tempStatus As String
   Dim tempBenefitPct As Double, tempPayType As String
   Dim tempFreq As String, tempNext As Date
   Dim tempHireDate As Date, tempOTRate As Double
   Dim tempTerm As Date, tempRate As Double
   Dim tempFedX As String, tempFedAmtPct As String
   Dim tempFedFig As Double, tempFedStatus As String
   Dim tempAllowNumFed As Integer, tempAddWHFed As Double
   Dim tempStateX As String, tempStateAmtPct As String
   Dim tempStateFig As Double, tempStateStatus As String
   Dim tempAllowNumState As Integer, tempAddWHState As Double
   Dim tempSocX As String, tempMedX As String
   Dim tempEIC As String, x As Integer
   Dim tempComment As String 'added 9/1/04
   Dim result As String
   Dim AN As String, AP As String
   Dim Amt As Double, INCOT As String
   Dim DedRec As DedCodeRecType
   Dim DHandle As Integer
   Dim DedCnt As Integer
   Dim Nextx As Integer
   Dim TotalWDDD As Double
   Dim tempHomePhone$, tempMainDept$, tempEmerCntctName$
   Dim tempEmerCntctPhone$, tempEmerRelationship$
   Dim CriticalDataChange As Boolean '7/25/03
   Dim PayUpdate As Boolean
   Dim UnitHandle As Integer
   Dim UnitFileRec As UnitFileRecType
   
   OpenUnitFile UnitHandle
   RecSize = LOF(UnitHandle) / Len(UnitFileRec)
   If RecSize = 0 Then
     MsgBox ("Please make sure the state is saved on the Employer Setup screen.")
     Exit Sub
   End If
   Get UnitHandle, 1, UnitFileRec
   Close UnitHandle
   
   PayUpdate = False
   frmLoadingRpt.Label1.Caption = "Saving......"
   frmLoadingRpt.Label2.Visible = False
   DoEvents
   frmLoadingRpt.Show
   DoEvents
   CriticalDataChange = False '7/25/03
   OpenDedCodeFile DHandle
   DedCnt = LOF(DHandle) / Len(DedRec)
   Close DHandle
   'If this is an existing employee then in order to properly
   'save any data that does not have screen fields we must
   'open the record for this employee then close it to capture
   'existing data for these fields and save it as is
   If Not newEmpFlag Then
     OpenEmpData2File EmpData2FileHandle
     Get #EmpData2FileHandle, thisRecordNum, EmpData2FileRec
     Close EmpData2FileHandle
   Else
     PayUpdate = True
   End If
'  fields that are required are examined for an entry and if
'  none is found then a message box issues an alert and the
'  process sends theuser back to the field where an entry must
'  be made...multiple alerts are queued "in effect" with the
'  re-focus moving backwards until the incorrect first field
'  is corrected at which time the process saves all data
   tempEmpNum = QPTrim$(frm1.txtNumber.Text)
   If tempEmpNum = "" Then
      MsgBox "Please enter a valid value in the Employee Number field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtNumber.SetFocus
      GoTo BadUnitData
   End If
   
   If QPTrim$(frm1.txtNumber.Text) <> QPTrim$(EmpData2FileRec.EmpNo) Then
     CriticalDataChange = True 'added 7/25/03...used to tell program to
     'unload the frmCustomerLookUp when exiting because if that form
     'appears it will not be updated with the number change
     'and any attempt to bring this employee back up will result in
     'a crash (bad record number)
     PayUpdate = True 'added 1/22/08
   End If
   
   tempEmpSSN = QPTrim$(Mid$(frm1.fpMaskSoc.Text, 1, 3) & Mid$(frm1.fpMaskSoc.Text, 5, 2) & Mid$(frm1.fpMaskSoc.Text, 8, 4))
   
   If Not Len(QPTrim(frm1.fpMaskSoc.Text)) = 11 Then
      MsgBox "Please enter a valid nine digit value in the Employee's Social Security Number field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.fpMaskSoc.SetFocus
      GoTo BadUnitData
   End If
   tempEmpLName = QPTrim$(frm1.txtLastName.Text)
   If tempEmpLName = "" Then
      MsgBox "Please enter a value in the Employee's Last Name field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtLastName.SetFocus
      GoTo BadUnitData
   End If
   
   If QPTrim$(frm1.txtLastName.Text) <> QPTrim$(EmpData2FileRec.EmpLName) Then
     CriticalDataChange = True 'added 7/25/03...used to tell program to
     'unload the frmCustomerLookUp when exiting because if that form
     'appears it will not be updated with the last name change
   End If
   
   tempEmpFName = QPTrim$(frm1.txtFirstName.Text)
   If tempEmpFName = "" Then
      MsgBox "Please enter a value in the Employee's First Name field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtFirstName.SetFocus
      GoTo BadUnitData
   End If
   If QPTrim$(frm1.txtFirstName.Text) <> QPTrim$(EmpData2FileRec.EmpFName) Then
     CriticalDataChange = True 'added 7/25/03...used to tell program to
     'unload the frmCustomerLookUp when exiting because if that form
     'appears it will not be updated with the first name change
   End If
   
   
   tempAddr1 = QPTrim$(frm1.txtAddress1.Text)
   tempAddr2 = QPTrim$(frm1.txtAddress2.Text)
   If tempAddr1 = "" And tempAddr2 = "" Then
      MsgBox "Please enter an address in one or both of the Address fields"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtAddress1.SetFocus
      GoTo BadUnitData
   End If
   tempCity = QPTrim$(frm1.txtCity.Text)
   If tempCity = "" Then
      MsgBox "Please enter a city in the Employee's City field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtCity.SetFocus
      GoTo BadUnitData
   End If
   tempState = QPTrim$(frm1.txtState.Text)
   If tempState = "" Then
      MsgBox "Please enter a state in the Employee's State field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtState.SetFocus
      GoTo BadUnitData
   End If
   tempZip = QPTrim$(ReplaceString(frm1.txtZip.Text, "-", ""))
   If tempZip = "" Then
      MsgBox "Please enter a zip code in the Employee's Zip Code field"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.txtZip.SetFocus
      GoTo BadUnitData
   End If
   If Len(QPTrim$(frm1.fpMaskBDay.Text)) = 0 Or CheckValDate(frm1.fpMaskBDay.Text) = False Then
      EmpData2FileRec.EMPBDAY = 0 'birthday is optional but
   'we don't want to save an empty field as NULL
   Else
      EmpData2FileRec.EMPBDAY = DateDiff("d", "12/31/1979", frm1.fpMaskBDay.Text)
      
   End If
   tempGender = QPTrim$(frm1.fpcomboGender.Text)
   If tempGender = "" Then
      MsgBox "Please select the Employee's Gender"
      frm1.vaTabPro1.ActiveTab = 0
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboGender.SetFocus
      GoTo BadUnitData
   End If
   tempRace = QPTrim$(frm1.fptxtRace.Text)
   tempRetNo = QPTrim$(frm1.fptxtRetNum.Text) 'if a retirement number is
   'entered then a retirement type must also be entered
   tempRetType = QPTrim$(frm1.fpcomboRetType.Text)
   If tempRetNo = Null Then tempRetNo = ""
   'if a user enters a value in the retirement type field and
   'does not enter a value in the retirement number field then
   'the retirement type value is not saved...an alert triggered
   'when the retirement type field loses focus tells the user
   'to put an entry in the retirement number field if he/she
   'has not already done so
   If tempRetNo = "" Then
      tempRetType = ""
   ElseIf Len(QPTrim$(frm1.fptxtRetNum.Text)) > 0 Then
      If tempRetType = "" Then
         MsgBox "Since there is a value in the Retirement Number field you must select a value in the Retirement Type field"
         frm1.vaTabPro1.ActiveTab = 0
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboRetType.SetFocus
         GoTo BadUnitData
      End If
   End If
   
   '**************************added 11/12/2002***************
   
   tempHomePhone$ = QPTrim$(frm1.fptxtHomePhone.Text)
   tempMainDept$ = QPTrim$(frm1.fptxtMainDept.Text)
   tempEmerCntctName$ = QPTrim$(frm1.fptxtContactName.Text)
   tempEmerCntctPhone$ = QPTrim$(frm1.fptxtContactPhone.Text)
   tempEmerRelationship$ = QPTrim$(frm1.fptxtRelationship.Text)
   
   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   'if the user did not fill in the bank draft code field
   'then none of the following 5 fields should have values
   'saved in them...an alert is issued at the form if any of these
   '5 fields are filled in without the bank draft code field
   'filled in
   
  'if bank draft code is left empty then the following 5 fields
  'should also be left empty
   frm1.fpcomboBankdraft.Col = 0
   If frm1.fpcomboBankdraft.ColText = "" Then
     If Len(frm1.txtBankAcctNo.Text) > 0 Then GoTo noCode
     If Len(frm1.txtBankName.Text) > 0 Then GoTo noCode
     If Len(frm1.fpcomboPrenoted.Text) > 0 Then GoTo noCode
     If Len(frm1.txtBankLocation.Text) > 0 Then GoTo noCode
     If Len(frm1.txtBankTransNo.Text) > 0 Then GoTo noCode
   End If
   GoTo codeOK
noCode:
   result = MsgBox("A value in the BankDraft Code field is required or no entries in that block will be saved. Do you wish to exit anyway?", vbYesNo)
   If result = vbNo Then
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboBankdraft.SetFocus
      GoTo BadUnitData
   End If
codeOK:
   frm1.fpcomboBankdraft.Col = 0
   tempBankDraft = QPTrim$(frm1.fpcomboBankdraft.ColText)
   If tempBankDraft = "" Then
      tempBankNum = ""
      tempPreNoted = ""
      tempBankName = ""
      tempBankLoc = ""
      tempBankTransNo = ""
   Else
   'if bank draft code is filled then the next 5 fields get
   'whatever value is in their fields except Prenoted which
   'isn't required
      tempBankNum = QPTrim$(frm1.txtBankAcctNo.Text)
      'give prenoted an empty string value if it is NULL
      If frm1.fpcomboPrenoted.Text = Null Then
         frm1.fpcomboPrenoted.Text = ""
      ElseIf Len(frm1.fpcomboPrenoted.Text) > 0 Then
         tempPreNoted = QPTrim$(frm1.fpcomboPrenoted.Text)
      Else
         tempPreNoted = ""
      End If
      
      tempBankName = QPTrim$(frm1.txtBankName.Text)
      tempBankLoc = QPTrim$(frm1.txtBankLocation.Text)
      'some of the old records were getting zeros which
      'caused a flag
      If QPTrim$(frm1.txtBankTransNo.Text) = "0" Then
         tempBankTransNo = ""
      Else
         tempBankTransNo = QPTrim$(frm1.txtBankTransNo.Text)
      End If
   End If
   
   'if bank draft code field has a value then the rest of the
   'values in that block must also be filled in except
   'Prenoted and if they aren't the user gets a message box alert
   If Not tempBankDraft = "" Then
      If tempBankNum = "" Then
         MsgBox "Please select the Employee's Bank Account Number"
         frm1.vaTabPro1.ActiveTab = 1
         frm1.vaTabPro1.SetFocus
         frm1.txtBankAcctNo.SetFocus
         GoTo BadUnitData
      End If
      If tempBankName = "" Then
         MsgBox "Please select the Employee's Bank's Name"
         frm1.vaTabPro1.ActiveTab = 1
         frm1.vaTabPro1.SetFocus
         frm1.txtBankName.SetFocus
         GoTo BadUnitData
      End If
      If tempBankLoc = "" Then
         MsgBox "Please select the Employee's Bank's Location"
         frm1.vaTabPro1.ActiveTab = 1
         frm1.vaTabPro1.SetFocus
         frm1.txtBankLocation.SetFocus
         GoTo BadUnitData
      End If
      If tempBankTransNo = "" Then
         MsgBox "Please select the Employee's Bank Transit Number"
         frm1.vaTabPro1.ActiveTab = 1
         frm1.vaTabPro1.SetFocus
         frm1.txtBankTransNo.SetFocus
         GoTo BadUnitData
      End If
   End If
   
   tempJobTitle = QPTrim$(frm1.txtTitle.Text)
   If tempJobTitle = Null Then tempJobTitle = "" 'No NULL
   'values should be saved in this optional field
   tempWCCode = QPTrim$(frm1.fptxtWCCode.Text)
   If tempWCCode = "" Then
      MsgBox "Please select the Employee's W/C Code"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtWCCode.SetFocus
      GoTo BadUnitData
   End If
   tempStatus = QPTrim$(frm1.fpcomboStatus.Text)
   If tempStatus = "" Then
      MsgBox "Please select a status in the Employee's Status field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboStatus.SetFocus
      GoTo BadUnitData
   End If
   tempBenefitPct = Val(ReplaceString(frm1.fptxtBenefitPct.Text, "%", ""))
   If tempBenefitPct < 0 Or tempBenefitPct > 100 Or QPTrim$(frm1.fptxtBenefitPct.Text) = "" Then
      MsgBox "Please enter a valid percentage in the Employee's Benefit Pct field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtBenefitPct.SetFocus
      GoTo BadUnitData
   End If
   
   tempPayType = QPTrim$(frm1.fpcomboPayType.Text)
   If tempPayType = "" Then
      MsgBox "Please select a pay type in the Employee's Pay Type field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboPayType.SetFocus
      GoTo BadUnitData
   End If
   PayType = tempPayType 'added 7/28/04
   If thisRecordNum > 0 Then 'added 7/28/04
     If tempPayType <> QPTrim$(EmpData2FileRec.EMPPTYPE) Then
       PayUpdate = True
     End If
   End If
   
   tempFreq = QPTrim$(frm1.fpcomboFreq.Text)
   If tempFreq = "" Then
      MsgBox "Please select a pay frequency in the Employee's Frequency field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboFreq.SetFocus
      GoTo BadUnitData
   End If
   ThisFreq = tempFreq 'added 7/28/04
   If thisRecordNum > 0 Then 'added 7/28/04
     If tempFreq <> QPTrim$(EmpData2FileRec.EMPPFREQ) Then
       PayUpdate = True
     End If
   End If
   
   tempRate = frm1.fptxtRate.Text
   If tempRate <= 0 Or QPTrim$(frm1.fptxtRate.Text) = "" Then
      MsgBox "Please enter a pay rate in the Employee's Rate field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtRate.SetFocus
      GoTo BadUnitData
   End If
   RegRate = Val(tempRate) 'added 7/28/04
   If thisRecordNum > 0 Then 'added 7/28/04
     If Val(tempRate) <> EmpData2FileRec.EMPPRATE Then
       PayUpdate = True
     End If
   End If
   
   
   tempOTRate = frm1.fptxtOTRate.Text
   If tempOTRate < 0 Or QPTrim$(frm1.fptxtOTRate.Text) = "" Then
      MsgBox "Please enter an overtime pay rate in the Employee's OT Rate field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtOTRate.SetFocus
      GoTo BadUnitData
   End If
   OTRate = Val(tempOTRate) 'added 7/28/04
   If thisRecordNum > 0 Then 'added 7/28/04
     If Val(tempOTRate) <> EmpData2FileRec.EMPORATE Then
       PayUpdate = True
     End If
   End If
   
   If Not CheckValDate(frm1.fpMaskHire.Text) Then
      MsgBox "Please enter a valid hire date in the Employee's Hire Date field"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpMaskHire.SetFocus
      GoTo BadUnitData
   Else
      EmpData2FileRec.EMPHDATE = DateDiff("d", "12/31/1979", frm1.fpMaskHire.Text)
   End If
   'optional fields
   
   If Len(QPTrim$(frm1.fpMaskNext.Text)) = 0 Then
      EmpData2FileRec.EMPRDATE = 0
   ElseIf CheckValDate(frm1.fpMaskNext.Text) = False Then
      MsgBox "The date entered in the Review Next field is not valid"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpMaskNext.SetFocus
      GoTo BadUnitData
   Else
      EmpData2FileRec.EMPRDATE = DateDiff("d", "12/31/1979", frm1.fpMaskNext.Text)  'No Nulls should be saved in
   End If
   
   'optional fields
   If Len(QPTrim$(frm1.fpMaskTerm.Text)) = 0 Then
      EmpData2FileRec.EMPTDATE = 0
   ElseIf CheckValDate(frm1.fpMaskTerm.Text) = False Then
      MsgBox "The date entered in the Termination Date field is not valid"
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpMaskTerm.SetFocus
      GoTo BadUnitData
   Else
      EmpData2FileRec.EMPTDATE = DateDiff("d", "12/31/1979", frm1.fpMaskTerm.Text) 'No Nulls should be saved in
   End If
   
'___Check for blank required fields____________________________________________________
   tempFedX = QPTrim$(frm1.fpcomboFedX.Text)
   If tempFedX = "" Then 'tempFedX = "N"
     MsgBox "Please select a value in the Federal Exempt field"
     frm1.vaTabPro1.ActiveTab = 3
     frm1.vaTabPro1.SetFocus
     frm1.fpcomboFedX.SetFocus
     GoTo BadUnitData
   End If
   tempFedAmtPct = QPTrim$(frm1.fpcomboFedAmtPct.Text)
   If tempFedAmtPct = Null Then tempFedAmtPct = "" 'No NULL values
   'should be saved for optional fields
   tempFedFig = Val(frm1.fptxtFedFig.Text)
   If tempFedFig = Null Then tempFedFig = 0 'No NULL values
   'should be saved for optional fields
   tempFedStatus = QPTrim$(frm1.fpcomboFedStatus.Text)
   If tempFedStatus = "" Then
      MsgBox "Please select a value in the Federal Status field"
      frm1.vaTabPro1.ActiveTab = 3
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboFedStatus.SetFocus
      GoTo BadUnitData
   End If
   tempAllowNumFed = Val(frm1.fptxtAllowNumFed.Text)
   If tempAllowNumFed < 0 Or QPTrim$(frm1.fptxtAllowNumFed.Text) = "" Then
      MsgBox "Please enter a value in the # Federal Allowances field"
      frm1.vaTabPro1.ActiveTab = 3
      frm1.vaTabPro1.SetFocus
      frm1.fptxtAllowNumFed.SetFocus
      GoTo BadUnitData
   End If
   
   tempAddWHFed = frm1.fptxtAddWHFed.Text
   If tempAddWHFed = Null Then tempAddWHFed = 0 'No NULL values
   'should be saved for optional fields
   tempStateX = QPTrim$(frm1.fpcomboStateX.Text)
   If tempStateX = "" Then 'tempStateX = "N"
     MsgBox "Please select a value in the State Exempt field"
     frm1.vaTabPro1.ActiveTab = 3
     frm1.vaTabPro1.SetFocus
     frm1.fpcomboStateX.SetFocus
     GoTo BadUnitData
   End If
   tempStateAmtPct = QPTrim$(frm1.fpcomboStateAmtPct.Text)
   If tempStateAmtPct = Null Then tempStateAmtPct = "" 'No NULL values
   'should be saved for optional fields
   tempStateFig = Val(frm1.fptxtStateFig.Text)
   If tempStateFig = Null Then tempStateFig = 0 'No NULL values
   'should be saved for optional fields
   tempStateStatus = QPTrim$(frm1.fpcomboStateStatus.Text)
   If tempStateStatus = "" Then
      MsgBox "Please select a value in the State Status field"
      frm1.vaTabPro1.ActiveTab = 3
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboStateStatus.SetFocus
      GoTo BadUnitData
   End If
   tempAllowNumState = Val(frm1.fptxtAllowNumState.Text)
   If tempAllowNumState < 0 Or QPTrim$(frm1.fptxtAllowNumState.Text) = "" Then
      MsgBox "Please enter a value in the # State Allowances field"
      frm1.vaTabPro1.ActiveTab = 3
      frm1.vaTabPro1.SetFocus
      frm1.fptxtAllowNumState.SetFocus
      GoTo BadUnitData
   End If
   tempAddWHState = frm1.fptxtAddWHState.Text
   If tempAddWHState = Null Then tempAddWHState = 0 'No NULL values
   'should be saved for optional fields
   tempSocX = QPTrim$(frm1.fpcomboSocX.Text)
   If tempSocX = "" Then
     MsgBox "Please select a value in the Social Security Exempt field"
     frm1.vaTabPro1.ActiveTab = 3
     frm1.vaTabPro1.SetFocus
     frm1.fpcomboSocX.SetFocus
     GoTo BadUnitData
   End If
   tempMedX = QPTrim$(frm1.fpcomboMedX.Text)
   If tempMedX = "" Then
     MsgBox "Please select a value in the Medicare Exempt field"
     frm1.vaTabPro1.ActiveTab = 3
     frm1.vaTabPro1.SetFocus
     frm1.fpcomboMedX.SetFocus
     GoTo BadUnitData
   End If
   tempEIC = QPTrim$(frm1.fpcomboEIC.Text)
   If tempEIC = "" Then
      MsgBox "Please enter a value in the EIC Code field"
      frm1.vaTabPro1.ActiveTab = 3
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboEIC.SetFocus
      GoTo BadUnitData
   End If
   
   tempComment = QPTrim$(frm1.fptxtComment.Text) 'added 9/1/2004
   
'PAGE 1__Assign Values to Types_______________________________________________________________
'NOTE: Dates are saved in the validation routines above
   EmpData1FileRec.EmpNo = QPTrim$(tempEmpNum)
   RSet EmpData1FileRec.EmpNo = QPTrim$(EmpData1FileRec.EmpNo)
   EmpData1FileRec.EmpLName = tempEmpLName
   EmpData1FileRec.EmpFName = tempEmpFName
   EmpData2FileRec.EmpNo = QPTrim$(tempEmpNum)
   RSet EmpData2FileRec.EmpNo = QPTrim$(EmpData2FileRec.EmpNo)
   EmpData2FileRec.EmpSSN = tempEmpSSN
   EmpData2FileRec.EmpLName = tempEmpLName
   EmpData2FileRec.EmpFName = tempEmpFName
   EmpData2FileRec.EmpAddr1 = tempAddr1
   EmpData2FileRec.EMPADDR2 = tempAddr2
   
   EmpData2FileRec.EmpCity = tempCity
   EmpData2FileRec.EmpState = tempState
   EmpData2FileRec.EmpZip = tempZip
   EmpData2FileRec.EMPGENDR = tempGender
   EmpData2FileRec.EMPRACE = tempRace
   EmpData2FileRec.EMPRETNO = tempRetNo
   
   EmpData2FileRec.EMPRETTP = tempRetType
   
   '***********added 11/12/2002
   EmpData2FileRec.HomePhone = tempHomePhone$
   EmpData2FileRec.PrimeDept = tempMainDept$
   EmpData2FileRec.EmrgncyCntctName = tempEmerCntctName$
   EmpData2FileRec.EmrgncyCntctPhnNum = tempEmerCntctPhone$
   EmpData2FileRec.EmrgncyCntctRelation = tempEmerRelationship$
   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
   EmpData2FileRec.DRAFTCOD = tempBankDraft
   EmpData2FileRec.EMPDDACC = tempBankNum
   EmpData2FileRec.PRENOTED = tempPreNoted
   EmpData2FileRec.BankName = tempBankName
   EmpData2FileRec.BANKLOC = tempBankLoc
   EmpData2FileRec.TRANSIT = tempBankTransNo
   
   EmpData2FileRec.EMPJOB = tempJobTitle
   EmpData2FileRec.EMPWCCLS = tempWCCode
   EmpData2FileRec.EMPSTATS = tempStatus
   EmpData2FileRec.EMPBCODE = tempBenefitPct
   EmpData2FileRec.EMPPTYPE = tempPayType
   EmpData2FileRec.EMPPFREQ = tempFreq
   EmpData2FileRec.EMPPRATE = tempRate
   EmpData2FileRec.EMPORATE = tempOTRate
  
'__Assign Values to Types_______________________________________________________________
   
   EmpData2FileRec.EMPFEDX = tempFedX
   EmpData2FileRec.EMPFEDO2 = tempFedAmtPct
   EmpData2FileRec.EMPFEDO1 = tempFedFig
   EmpData2FileRec.EMPFEDS = tempFedStatus
   EmpData2FileRec.EMPFEDA = tempAllowNumFed       'num of allowance
   EmpData2FileRec.EMPFEDAA = tempAddWHFed
   EmpData2FileRec.EMPSTAX = tempStateX
   EmpData2FileRec.EMPSTAO2 = tempStateAmtPct
   EmpData2FileRec.EMPSTAO1 = tempStateFig
   EmpData2FileRec.EMPSTAS = tempStateStatus
   EmpData2FileRec.EMPSTAA = tempAllowNumState
   EmpData2FileRec.EMPSTAAA = tempAddWHState
   EmpData2FileRec.EMPSOCX = tempSocX
   EmpData2FileRec.EMPMEDX = tempMedX
   EmpData2FileRec.EMPEIC = tempEIC
   For x = 1 To 50
      frm1.vaSpreadMisc.Col = 2
      frm1.vaSpreadMisc.Row = x
      AP = QPTrim$(frm1.vaSpreadMisc.Text) 'AP = Amount/Percent column
      frm1.vaSpreadMisc.Col = 4
      frm1.vaSpreadMisc.Row = x
      INCOT = QPTrim$(frm1.vaSpreadMisc.Text) 'INCOT = Include Overtime column
      If AP = "AMOUNT" Or AP = "" Or AP = "Amount" Then
         If Len(INCOT) > 0 Then
            MsgBox "A value is not allowed in the Inc O/T field if AMOUNT is selected in the Amt/Pct field"
            frm1.vaSpreadMisc.Text = ""
            frm1.vaTabPro1.ActiveTab = 4
            frm1.vaSpreadMisc.SetFocus
            frm1.vaSpreadMisc.SetActiveCell 4, x
            GoTo BadUnitData
         End If
      End If
      If UCase$(AP) = "AMOUNT" And x > DedCnt Then
        MsgBox "Error: Cannot save a value where there is no description."
        frm1.vaTabPro1.ActiveTab = 4
        frm1.vaSpreadMisc.SetFocus
        frm1.vaSpreadMisc.SetActiveCell 2, x
        GoTo BadUnitData
      End If
      If AP = "PERCENT" Or AP = "Percent" Then
         If Len(INCOT) = 0 Then
            MsgBox "You must enter a value in the Inc O/T field if PERCENT is selected in the Amt/Pct field"
            frm1.vaTabPro1.ActiveTab = 4
            frm1.vaSpreadMisc.SetFocus
            frm1.vaSpreadMisc.SetActiveCell 4, x
            GoTo BadUnitData
          End If
      End If
   Next x
   
   For x = 1 To 50
       frm1.vaSpreadMisc.Col = 2
       frm1.vaSpreadMisc.Row = x
       'DPct is either the word "amount" or "percent"
       EmpData2FileRec.EmpDed(x).DPct = QPTrim$(frm1.vaSpreadMisc.Text)
       If frm1.vaSpreadMisc.Text = "" Or frm1.vaSpreadMisc.Text = Null Then EmpData2FileRec.EmpDed(x).DPct = ""
       frm1.vaSpreadMisc.Col = 3
       frm1.vaSpreadMisc.Row = x
       'DAmt is the numeric value of DPct
       EmpData2FileRec.EmpDed(x).DAmt = Val(frm1.vaSpreadMisc.Text)
       'if a choice is made in the Amt/Pct field then a value must be in the
       'withholding field
       If Val(EmpData2FileRec.EmpDed(x).DAmt) = 0 And Not QPTrim$(EmpData2FileRec.EmpDed(x).DPct) = "" Then
          MsgBox "Please enter a value in the Withholding field"
          frm1.vaTabPro1.ActiveTab = 4
          frm1.vaSpreadMisc.SetFocus
          frm1.vaSpreadMisc.SetActiveCell 3, x
          GoTo BadUnitData
       'if a choice is not made in the Amt/Pct field but a value is in the
       'Withholding field
       ElseIf QPTrim$(EmpData2FileRec.EmpDed(x).DPct) = "" And EmpData2FileRec.EmpDed(x).DAmt > 0 Then
          MsgBox "Please either make a selection in the Amt/Pct field or delete the value in the Withholding field"
          frm1.vaTabPro1.ActiveTab = 4
          frm1.vaSpreadMisc.SetFocus
          frm1.vaSpreadMisc.SetActiveCell 3, x
          GoTo BadUnitData
       ElseIf frm1.vaSpreadMisc.Text = "" Or frm1.vaSpreadMisc.Text = Null Then
          EmpData2FileRec.EmpDed(x).DAmt = 0
       End If
       frm1.vaSpreadMisc.Col = 4
       frm1.vaSpreadMisc.Row = x
       EmpData2FileRec.EmpDed(x).DOTI = QPTrim$(frm1.vaSpreadMisc.Text)
       If frm1.vaSpreadMisc.Text = "" And QPTrim$(UCase$(EmpData2FileRec.EmpDed(x).DPct)) = "PERCENT" Then
          frm1.vaTabPro1.ActiveTab = 4
          frm1.vaSpreadMisc.SetFocus
          frm1.vaSpreadMisc.SetActiveCell 4, x
          GoTo BadUnitData
       ElseIf frm1.vaSpreadMisc.Text = "" Or frm1.vaSpreadMisc.Text = Null Then
          EmpData2FileRec.EmpDed(x).DOTI = ""
       End If
   Next x

'__Assign Values to Types_______________________________________________________________
   
   If frm1.fptxtAN(1).Text = Null Or frm1.fptxtAN(1).Text = "" Then
      EmpData2FileRec.EMPEACT1 = ""
   Else
      EmpData2FileRec.EMPEACT1 = QPTrim$(frm1.fptxtAN(1).Text)
   End If
   
   If frm1.fptxtE(1).Text = Null Or frm1.fptxtE(1).Text = "" Then
      EmpData2FileRec.EMPEAMT1 = 0
   Else
      EmpData2FileRec.EMPEAMT1 = frm1.fptxtE(1).Text
   End If
   
   If frm1.fptxtAN(2).Text = Null Or frm1.fptxtAN(2).Text = "" Then
      EmpData2FileRec.EMPEACT2 = ""
   Else
      EmpData2FileRec.EMPEACT2 = QPTrim$(frm1.fptxtAN(2).Text)
   End If
   
   If frm1.fptxtE(2).Text = Null Or frm1.fptxtE(2).Text = "" Then
      EmpData2FileRec.EMPEAMT2 = 0
   Else
      EmpData2FileRec.EMPEAMT2 = frm1.fptxtE(2).Text
   End If
   
   If frm1.fptxtAN(3).Text = Null Or frm1.fptxtAN(3).Text = "" Then
      EmpData2FileRec.EMPEACT3 = ""
   Else
      EmpData2FileRec.EMPEACT3 = QPTrim$(frm1.fptxtAN(3).Text)
   End If
   If frm1.fptxtE(3).Text = Null Or frm1.fptxtE(3).Text = "" Then
      EmpData2FileRec.EMPEAMT3 = 0
   Else
      EmpData2FileRec.EMPEAMT3 = frm1.fptxtE(3).Text
   End If
   'ALERT: A new version will have 5 entries above instead of 3 ^
   'this for loop insures that at least one entry is made
   'in the WDAN field
   For x = 1 To 8
     If Len(QPTrim$(frm1.fptxtWDAN(x).Text)) > 0 Then Exit For
   Next x
   If x = 9 Then
     MsgBox "Please make a selection in the Account Number field"
     frm1.vaTabPro1.ActiveTab = 6
     frm1.fptxtWDAN(1).SetFocus
     GoTo BadUnitData
   End If
   TotalWDDD = 0
   For x = 1 To 8
      AN = QPTrim$(frm1.fptxtWDAN(x).Text)
      Amt = Val(frm1.fptxtWDDD(x).Text)
      TotalWDDD = TotalWDDD + Amt
      If Amt > 0 Then
         If Len(AN) = 0 Then
            MsgBox "Please enter an Account Number or delete the current value in Default Distribution."
            frm1.vaTabPro1.ActiveTab = 6
            frm1.vaTabPro1.SetFocus
            frm1.fptxtWDAN(x).SetFocus
            GoTo BadUnitData
         End If
      End If
   Next x
   'trap for erroneous total distribution tally for salaried employees
   If TotalWDDD > 100 And Mid(EmpData2FileRec.EMPPTYPE, 1, 1) = "S" Then
     MsgBox "Total default distribution cannot be more than 100% for salaried employees."
     frm1.vaTabPro1.ActiveTab = 6
     frm1.fptxtWDDD(1).SetFocus
     GoTo BadUnitData
   End If
     
   Nextx = 1
   
   For x = 1 To 8
      If frm1.fptxtWDAN(x).Text = Null Then
         EmpData2FileRec.EDist(x).DAcct = ""
      Else
         EmpData2FileRec.EDist(x).DAcct = QPTrim$(frm1.fptxtWDAN(x).Text)
      End If
      If frm1.fptxtWDDD(x).Text = Null Then
         EmpData2FileRec.EDist(x).DAmt = 0
      Else
         EmpData2FileRec.EDist(x).DAmt = Val(frm1.fptxtWDDD(x).Text)
      End If
   Next x
   
   If frm1.fptxtEarned(1).Text = Null Or QPTrim$(frm1.fptxtEarned(1).Text) = "" Then
      EmpData2FileRec.EMPVACE = 0
   Else
      EmpData2FileRec.EMPVACE = Val(frm1.fptxtEarned(1).Text)
   End If
   
   If frm1.fptxtUsed(1).Text = Null Or QPTrim$(frm1.fptxtUsed(1).Text) = "" Then
      EmpData2FileRec.EMPVUSED = 0
   Else
      EmpData2FileRec.EMPVUSED = Val(frm1.fptxtUsed(1).Text)
   End If
   
   If frm1.fptxtBal(1).Text = Null Or QPTrim$(frm1.fptxtBal(2)) = "" Then
      EmpData2FileRec.EMPVBAL = 0
   Else
      EmpData2FileRec.EMPVBAL = Val(frm1.fptxtBal(1).Text)
   End If
   
   If frm1.fptxtEarned(2).Text = Null Or QPTrim$(frm1.fptxtEarned(2).Text) = "" Then
      EmpData2FileRec.EMPSLE = 0
   Else
      EmpData2FileRec.EMPSLE = Val(frm1.fptxtEarned(2).Text)
   End If
   
   If frm1.fptxtUsed(2).Text = Null Or QPTrim$(frm1.fptxtUsed(2).Text) = "" Then
      EmpData2FileRec.EMPSLUSE = 0
   Else
      EmpData2FileRec.EMPSLUSE = Val(frm1.fptxtUsed(2).Text)
   End If
   
   If frm1.fptxtBal(2).Text = Null Or QPTrim$(frm1.fptxtBal(2)) = "" Then
      EmpData2FileRec.EMPSLBAL = 0
   Else
      EmpData2FileRec.EMPSLBAL = Val(frm1.fptxtBal(2).Text)
   End If
   
   If frm1.fptxtEarned(3).Text = Null Or QPTrim$(frm1.fptxtEarned(3).Text) = "" Then
      EmpData2FileRec.EMPCTE = 0
   Else
      EmpData2FileRec.EMPCTE = Val(frm1.fptxtEarned(3).Text)
   End If
   
   If frm1.fptxtUsed(3).Text = Null Or QPTrim$(frm1.fptxtUsed(3).Text) = "" Then
      EmpData2FileRec.EMPCTUSE = 0
   Else
      EmpData2FileRec.EMPCTUSE = Val(frm1.fptxtUsed(3).Text)
   End If
   
   If frm1.fptxtBal(3).Text = Null Or QPTrim$(frm1.fptxtBal(3)) = "" Then
      EmpData2FileRec.EMPCTBAL = 0
   Else
      EmpData2FileRec.EMPCTBAL = Val(frm1.fptxtBal(3).Text)
   End If
   
   If frm1.fptxtEarned(4).Text = Null Or QPTrim$(frm1.fptxtEarned(4).Text) = "" Then
      EmpData2FileRec.PERERN = 0
   Else
      EmpData2FileRec.PERERN = Val(frm1.fptxtEarned(4).Text)
   End If
   
   If frm1.fptxtUsed(4).Text = Null Or QPTrim$(frm1.fptxtUsed(4).Text) = "" Then
      EmpData2FileRec.PerUsed = 0
   Else
      EmpData2FileRec.PerUsed = Val(frm1.fptxtUsed(4).Text)
   End If
   
   If frm1.fptxtBal(4).Text = Null Or QPTrim$(frm1.fptxtBal(4)) = "" Then
      EmpData2FileRec.PERBAL = 0
   Else
      EmpData2FileRec.PERBAL = Val(frm1.fptxtBal(4).Text)
   End If
   
   If frm1.fptxtEarned(5).Text = Null Or QPTrim$(frm1.fptxtEarned(5).Text) = "" Then
      EmpData2FileRec.HOLERN = 0
   Else
      EmpData2FileRec.HOLERN = Val(frm1.fptxtEarned(5).Text)
   End If
   
   If frm1.fptxtUsed(5).Text = Null Or QPTrim$(frm1.fptxtUsed(5).Text) = "" Then
      EmpData2FileRec.HolUsed = 0
   Else
      EmpData2FileRec.HolUsed = Val(frm1.fptxtUsed(5).Text)
   End If
   
   If frm1.fptxtBal(5).Text = Null Or QPTrim$(frm1.fptxtBal(5)) = "" Then
      EmpData2FileRec.HOLBAL = 0
   Else
      EmpData2FileRec.HOLBAL = Val(frm1.fptxtBal(5).Text)
   End If
   
   If frm1.fpcomboLT.Text = Null Or QPTrim$(frm1.fpcomboLT.Text) = "" Then
      EmpData2FileRec.LeaveTbl = 0
   Else
      EmpData2FileRec.LeaveTbl = Val(frm1.fpcomboLT.Text)
   End If

   '************added 11/12/2002
   If frm1.fpcombo401K.Text = Null Or QPTrim$(frm1.fpcombo401K.Text) = "" Then
      EmpData2FileRec.YN401K = "N"
   Else
      EmpData2FileRec.YN401K = QPTrim$(frm1.fpcombo401K.Text)
   End If
   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   
   If frm1.fpcomboESC.Text = Null Or QPTrim$(frm1.fpcomboESC.Text) = "" Then
      EmpData2FileRec.ExcludeESC = "N"
   Else
      EmpData2FileRec.ExcludeESC = QPTrim$(frm1.fpcomboESC.Text)
   End If
   
   EmpData2FileRec.Comment = tempComment 'added 9/01/04
'Assign Emp3 all zeros since this is a new employee plus
'assign all Emp2 data with no screen fields either 0 or ""
   If newEmpFlag = True Then
     EmpData2FileRec.UseLife = ""
     EmpData2FileRec.LastTransRec = 0
     EmpData2FileRec.EmpPin = 0
     EmpData2FileRec.Deleted = 0
     EmpData2FileRec.LDTDate = 0
     EmpData2FileRec.CDTDate = 0
     EmpData2FileRec.InprocFlag = 0
     EmpData2FileRec.Unused = ""
     EmpData3FileRec.Data1RecNum = 0
     EmpData3FileRec.YTDGrossPay = 0
     EmpData3FileRec.YTDSocGrossPay = 0
     EmpData3FileRec.YTDMedGrossPay = 0
     EmpData3FileRec.YTDFedGrossPay = 0
     EmpData3FileRec.YTDStaGrossPay = 0
     EmpData3FileRec.YTDOTPay = 0
     EmpData3FileRec.YTDRegPay = 0
     EmpData3FileRec.YTDNet = 0
     EmpData3FileRec.YTDSocial = 0
     EmpData3FileRec.YTDMedicare = 0
     EmpData3FileRec.YTDFederal = 0
     EmpData3FileRec.YTDState = 0
     EmpData3FileRec.YTDRetire = 0
     For x = 1 To 50
       EmpData3FileRec.YTDDAmt(x) = 0
     Next x
     EmpData3FileRec.YTDDAmtT = 0
     EmpData3FileRec.YTDEarn1 = 0
     EmpData3FileRec.YTDEarn2 = 0
     EmpData3FileRec.YTDEarn3 = 0
     EmpData3FileRec.YTDEarnT = 0
     EmpData3FileRec.YTDEIC = 0
     EmpData3FileRec.YTDOther2 = 0
     TransRec.BaseRate = 0
     TransRec.CheckDate = 0
     TransRec.CheckNum = 0
     TransRec.CompUsed = 0
     For x = 1 To 50
       TransRec.DAmt(x) = 0
     Next x
     For x = 1 To 3
       TransRec.EAmt(x) = 0
     Next x
     For x = 1 To 6
       TransRec.EDist(x).EAcct = ""
       TransRec.EDist(x).EAmt = 0
     Next x
     TransRec.EICAmt = 0
     TransRec.EmpPin = 0
     TransRec.FedGrossPay = 0
     TransRec.FedTaxAmt = 0
     TransRec.GrossPay = 0
     TransRec.GrossWage = 0
     TransRec.HOLHOURS = 0
     TransRec.MatchMedAmt = 0
     TransRec.MatchRetAmt = 0
     TransRec.MatchSocAmt = 0
     TransRec.MedGrossPay = 0
     TransRec.MedTaxAmt = 0
     TransRec.NetPay = 0
     TransRec.OT2Comp = 0
     TransRec.OTHours = 0
     TransRec.OTHrsPaid = 0
     TransRec.OTRate = 0
     TransRec.Pad1 = ""
     TransRec.PayPdEnd = 0
     TransRec.PayPdStart = 0
     TransRec.PaySFlag = ""
     TransRec.PayType = ""
     TransRec.PerHours = 0
     TransRec.PeriodHistRec = 0
     TransRec.PostDate = 0
     TransRec.PrevTransRec = 0
     TransRec.RegHrsPaid = 0
     TransRec.RegHrsWork = 0
     TransRec.RetGrossPay = 0
     TransRec.RetireAmt = 0
     TransRec.SickUsed = 0
     TransRec.SocGrossPay = 0
     TransRec.SocTaxAmt = 0
     TransRec.StaGrossPay = 0
     TransRec.StaTaxAmt = 0
     TransRec.TActive = 0
     For x = 1 To 8
        TransRec.TDist(x).DAcct = ""
        TransRec.TDist(x).DOHrs = 0
        TransRec.TDist(x).DOWage = 0
        TransRec.TDist(x).DPct = 0
        TransRec.TDist(x).DRHrs = 0
        TransRec.TDist(x).DRWage = 0
     Next x
     TransRec.TaxFring = 0
     TransRec.TotAdditEarn = 0
     TransRec.TotDedAmt = 0
     TransRec.TotOTWage = 0
     TransRec.TotRegWage = 0
     TransRec.TotTaxAmt = 0
     TransRec.VacUsed = 0
     For x = 1 To 3
       CheckRec.AEarn(x).DAmt = 0
       CheckRec.AEarn(x).DCode = ""
       CheckRec.AEarn(x).YTDDAmt = 0
     Next x
     CheckRec.BaseRate = 0
     CheckRec.CActive = 0
     For x = 1 To 50
       CheckRec.CDED(x).DAmt = 0
       CheckRec.CDED(x).DCode = ""
       CheckRec.CDED(x).YTDDAmt = 0
     Next x
     CheckRec.CheckDate = 0
     CheckRec.CheckNum = 0
     CheckRec.CompBal = 0
     CheckRec.CompEarn = 0
     CheckRec.CompUsed = 0
     CheckRec.DDFlag = 0
     CheckRec.EICAmt = 0
     CheckRec.EmpAddr1 = ""
     CheckRec.EmpCity = ""
     CheckRec.EmpName = ""
     CheckRec.EmpNo = ""
     CheckRec.EmpSSN = ""
     CheckRec.EmpState = ""
     CheckRec.EmpZip = ""
     CheckRec.FedTaxAmt = 0
     CheckRec.GrossPay = 0
     CheckRec.HolUsed = 0
     CheckRec.MedTaxAmt = 0
     CheckRec.NetPay = 0
     CheckRec.OTHrsPaid = 0
     CheckRec.PayEndDate = 0
     CheckRec.PerUsed = 0
     CheckRec.RegHrsPaid = 0
     CheckRec.RegHrsWork = 0
     CheckRec.RetireAmt = 0
     CheckRec.SickBal = 0
     CheckRec.SickUsed = 0
     CheckRec.SocTaxAmt = 0
     CheckRec.StaTaxAmt = 0
     CheckRec.TaxFring = 0
     CheckRec.TotAdditEarn = 0
     CheckRec.TotDedAmt = 0
     CheckRec.TotOTWage = 0
     CheckRec.TotRegWage = 0
     CheckRec.VactBal = 0
     CheckRec.VacUsed = 0
     CheckRec.YTDFederal = 0
     CheckRec.YTDGrossPay = 0
     CheckRec.YTDMedicare = 0
     CheckRec.YTDNetPay = 0
     CheckRec.YTDRetire = 0
     CheckRec.YTDSocial = 0
     CheckRec.YTDState = 0
     CheckRec.YTDTotDed = 0
     
  End If

'Save Procedure_______________________________________________________
   OpenEmpData2File EmpData2FileHandle
   OpenEmpData1File EmpData1FileHandle
   EmpRecNo = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
   If newEmpFlag = True Then
      EPN = FreeFile
      Open PRData + EMPPinFileName For Random As EPN Len = 2
      Get EPN, 1, LastPin
      NextPin = LastPin + 1
      Put EPN, 1, NextPin
      Close EPN
      saveHere = EmpRecNo + 1
      EmpData1FileRec.Data1RecNum = EmpRecNo + 1
      EmpData1FileRec.TransRecNum = EmpRecNo + 1
      EmpData1FileRec.Deleted = 0
      EmpData2FileRec.EmpPin = NextPin
      
   'thisRecordNum is passed to this function
   'saveHere can be either the new next record number or the old
   'record number of the employee being edited
   ElseIf newEmpFlag = False Then
      EmpData1FileRec.Data1RecNum = thisRecordNum
      EmpData1FileRec.TransRecNum = thisRecordNum
      saveHere = thisRecordNum
   End If
   Put EmpData1FileHandle, saveHere, EmpData1FileRec
   Close EmpData1FileHandle
   
   If PayUpdate = True Then
     If newEmpFlag = False Then
       Call UpdatePayRate(QPTrim$(EmpData2FileRec.EMPJOB), PayType, OTRate, RegRate, ThisFreq, saveHere, False)
     End If
   End If
   
   Put EmpData2FileHandle, saveHere, EmpData2FileRec
   Close EmpData2FileHandle
   
   OpenEmpNumFile EHandle
   EmpNumRec.EmpNum = tempEmpNum
   Put EHandle, saveHere, EmpNumRec
   Close EHandle
   
   If newEmpFlag = True Then
     Call UpdatePayRate(QPTrim(EmpData2FileRec.EMPJOB), PayType, OTRate, RegRate, ThisFreq, saveHere, True)
     OpenPRChecksFile CHandle 'save all CheckRec values
     Put CHandle, saveHere, CheckRec
     Close CHandle
     OpenTransWorkFile THandle 'save all TransRec values
     Put THandle, saveHere, TransRec
     Close THandle
     OpenEmpData3File EmpData3FileHandle
     Put EmpData3FileHandle, saveHere, EmpData3FileRec
     Close EmpData3FileHandle
   End If
   
   PayUpdate = False
   PayType = ""
   OTRate = 0
   RegRate = 0
   ThisFreq = ""
   If QPTrim$(UnitFileRec.UFSTATE) = "NC" And Exist(OrbitEmpDataBatch) Then
'   If frm1.txtState.Text = "NC" And Exist(OrbitEmpDataBatch) Then  'new for ORBIT
     Dim ORec As OrbitEmpData
     Dim OHandle As Integer
     Dim NumOfORecs As Integer
     Dim t As Integer
     Dim ThisSSN As String
     Dim OBRec As OrbitEmpDataBatch
     Dim OBHandle As Integer
     Dim NumOfOBRecs As Integer
     OpenOrbEmpDataBatch OBHandle, NumOfOBRecs
     OpenOrbEmpData OHandle, NumOfORecs
     If newEmpFlag = True Then
       saveHere = NumOfORecs + 1
     ElseIf Not Exist(OrbitEmpData) Then
       saveHere = 1
     Else
       For x = 1 To NumOfORecs
         Get OHandle, x, ORec
         If ORec.EmpRecNum = RecNum Then
           saveHere = x
           Exit For
         End If
       Next x
       If x > NumOfORecs Then
         saveHere = NumOfORecs + 1
       End If
     End If
     Get OBHandle, 1, OBRec
     If newEmpFlag = True Then
       OBRec.EmpRecNum = NextPin
     Else
       OBRec.EmpRecNum = RecNum
     End If
     Put OBHandle, 1, OBRec
     ORec.DateOfBirth = OBRec.DateOfBirth
     ORec.PlanCode = OBRec.PlanCode
     ORec.MemberID = OBRec.MemberID
     ORec.MiddleName = OBRec.MiddleName
     ORec.OutOfCntryAdd = OBRec.OutOfCntryAdd
     ORec.DeptNum = OBRec.DeptNum
     ORec.JobClass = OBRec.JobClass
     ORec.EligibleDate = OBRec.EligibleDate
     ORec.Adjustment = OBRec.Adjustment
     ORec.PayType = OBRec.PayType
     ORec.VacHours = OBRec.VacHours
     ORec.ContrPdEmpPrd = OBRec.ContrPdEmpPrd
     ORec.TerminationDate = OBRec.TerminationDate
     ORec.TermType = OBRec.TermType
     ORec.Adjustment = OBRec.Adjustment
     ORec.VacHours = OBRec.VacHours
     ORec.ContrPdEmpPrd = OBRec.ContrPdEmpPrd
     ORec.TermType = OBRec.TermType
     ORec.LastName = OBRec.LastName
     ORec.FirstName = OBRec.FirstName
     ORec.SSN = OBRec.SSN
     ORec.Gender = OBRec.Gender
     ORec.AddLine1 = OBRec.AddLine1
     ORec.AddLine2 = OBRec.AddLine2
     ORec.City = OBRec.City
     ORec.State = OBRec.State
     ORec.Zip = OBRec.Zip
     ORec.EmployDate = OBRec.EmployDate
     ORec.EmpRecNum = OBRec.EmpRecNum
     ORec.AgencyNum = OBRec.AgencyNum
     ORec.ContrPdEmpBegDate = OBRec.ContrPdEmpBegDate
     ORec.ContrPdEmpEndDate = OBRec.ContrPdEmpEndDate
     ORec.SharedPosition = OBRec.SharedPosition
     ORec.Suffix = OBRec.Suffix
     ORec.Deleted = OBRec.Deleted
     ORec.EmpNum = OBRec.EmpNum
     Put OHandle, saveHere, ORec
     Close OHandle
     Close OBHandle
     KillFile OrbitEmpDataBatch
     ThisSSN = ReplaceString(frm1.fpMaskSoc.Text, "-", "")
     
     If newEmpFlag = True Then GoTo NoNeed
     If QPTrim$(frm1.fptxtRetNum.Text) <> "" And QPTrim$(frm1.fpcomboRetType.Text) <> "" Then
       OpenOrbEmpData OHandle, NumOfORecs
       For t = 1 To NumOfORecs
         Get OHandle, t, ORec
         If ReplaceString(ORec.SSN, "_", "") = ThisSSN Then
           If ORec.EmpRecNum <> RecNum Then
             MainLog ("User alerted to employee record # " & CStr(RecNum) & " had an ORBIT EmpRecNum of " & CStr(ORec.EmpRecNum) & " that needs correcting.")
             MsgBox ("There is a problem with this employee's ORBIT record number. Please call Southern Software @ 1-800-842-8190 to correct this situation.")
           End If
         End If
       Next t
     End If
     Close OHandle
   End If
NoNeed:
   MsgBox "Your information has been saved"
   If newEmpFlag = True Then '7/25/03
     frmEmployeeMaintMenu.Show
     DoEvents
   Else
     Call frmEmployeeLookUp.ActivateControls '7/25/03
     If CriticalDataChange = True Then
       frmEmployeeMaintMenu.Show
       DoEvents
       Unload frmEmployeeLookUp
     End If
   End If
     
   Unload frmEditEmpData
   'reset RecNum to 0 because it is a global and it's value can
   'be carried over...its value is set as it is needed in procedures
   RecNum = 0
   Call MakeEmpIndexs
BadUnitData:
 Close
 Unload frmLoadingRpt

End Sub

Public Sub checkExitEmp(newEmpFlag As Boolean, thisRecordNum As Integer, frm1 As Form)

   Dim DoWhatFlag As SaveChangeOptions1
   Dim save As Integer, review As Integer, abandon As Integer
   Dim EmpData2FileHandle As Integer, changeFlag As Boolean
   Dim EmpData2FileRec As EmpData2Type, x As Integer
   Dim ErnCodeFileHandle As Integer
   Dim ErnCodeFileRec As ErnCodeRecType
   Dim tempBDay As Integer
   Dim tempHDate As Integer
   Dim tempNDate As Integer
   Dim tempTDate As Integer
   Dim SSN As String

   changeFlag = False
   If newEmpFlag = True Then 'no check done if employee is a new entry
     GoTo ChangeTrue:
   End If
   OpenEmpData2File EmpData2FileHandle
   Get EmpData2FileHandle, thisRecordNum, EmpData2FileRec
   Close EmpData2FileHandle
   'check each textbox to see if a change has been made
      

    If QPTrim$(EmpData2FileRec.EmpNo) <> QPTrim$(frm1.txtNumber.Text) Then
    'if it has then set the changeFlag to 1 and reset focus to
    'where the change was made
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtNumber.SetFocus
       GoTo ChangeTrue
    End If
    SSN = Mid$(frm1.fpMaskSoc.Text, 1, 3) + Mid$(frm1.fpMaskSoc.Text, 5, 2) + Mid$(frm1.fpMaskSoc.Text, 8, 4)
    If QPTrim$(EmpData2FileRec.EmpSSN) <> QPTrim$(SSN) And QPTrim$(EmpData2FileRec.EmpSSN) <> QPTrim$(frm1.fpMaskSoc.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fpMaskSoc.SetFocus
       GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EmpLName) <> QPTrim$(frm1.txtLastName.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtLastName.SetFocus
       GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EmpFName) <> QPTrim$(frm1.txtFirstName.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtFirstName.SetFocus
       GoTo ChangeTrue
    End If
      
    If QPTrim$(EmpData2FileRec.EmpAddr1) <> QPTrim$(frm1.txtAddress1.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtAddress1.SetFocus
       GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPADDR2) <> QPTrim$(frm1.txtAddress2.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtAddress2.SetFocus
       GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EmpCity) <> QPTrim$(frm1.txtCity.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtCity.SetFocus
       GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.EmpState) <> QPTrim$(frm1.txtState.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtState.SetFocus
       GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EmpZip) <> QPTrim$(ReplaceString$(frm1.txtZip.Text, "-", "")) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.txtZip.SetFocus
       GoTo ChangeTrue
    End If
    
    If newEmpFlag = True Then
      If EmpData2FileRec.EMPBDAY <> 0 Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 0
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskBDay.SetFocus
          GoTo ChangeTrue
      End If
    End If
    If newEmpFlag = False Then
       If EmpData2FileRec.EMPBDAY = 0 And frm1.fpMaskBDay.Text = "" Then
          GoTo BDay0
       ElseIf EmpData2FileRec.EMPBDAY < -21914 Then GoTo BDay0 ' -21914 is 01/01/1920
       ElseIf EmpData2FileRec.EMPBDAY <> 0 And frm1.fpMaskBDay.Text = "" Then
          changeFlag = True 'happens if an old employee's birthday date is deleted
          frm1.vaTabPro1.ActiveTab = 0
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskBDay.SetFocus
         GoTo ChangeTrue
       ElseIf EmpData2FileRec.EMPBDAY <> DateDiff("d", "12/31/1979", frm1.fpMaskBDay.Text) Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 0
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskBDay.SetFocus
          GoTo ChangeTrue
      End If
    End If
BDay0:
    If QPTrim$(EmpData2FileRec.EMPGENDR) <> QPTrim$(frm1.fpcomboGender.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fpcomboGender.SetFocus
       GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.EMPRACE) <> QPTrim$(frm1.fptxtRace.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fptxtRace.SetFocus
       GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.EMPRETNO) <> QPTrim$(frm1.fptxtRetNum.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fptxtRetNum.SetFocus
       GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPRETTP) <> QPTrim$(frm1.fpcomboRetType.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fpcomboRetType.SetFocus
       GoTo ChangeTrue
    End If
    
    '**************added 11/12/2002
'    If QPTrim$(EmpData2FileRec.PrimeDept) <> QPTrim$(frm1.fptxtMainDept.Text) Then
'       changeFlag = True
'       frm1.vaTabPro1.ActiveTab = 0
'       frm1.vaTabPro1.SetFocus
'       frm1.fpcomboMainDept.SetFocus
'       GoTo ChangeTrue
'    End If
    
    If QPTrim$(EmpData2FileRec.HomePhone) <> QPTrim$(frm1.fptxtHomePhone.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fptxtHomePhone.SetFocus
       GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.EmrgncyCntctName) <> QPTrim$(frm1.fptxtContactName.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fptxtContactName.SetFocus
       GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.EmrgncyCntctPhnNum) <> QPTrim$(frm1.fptxtContactPhone.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fptxtContactPhone.SetFocus
       GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.EmrgncyCntctRelation) <> QPTrim$(frm1.fptxtRelationship.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 0
       frm1.vaTabPro1.SetFocus
       frm1.fptxtRelationship.SetFocus
       GoTo ChangeTrue
    End If
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    
    frm1.fpcomboBankdraft.Col = 0
    If QPTrim$(EmpData2FileRec.DRAFTCOD) <> QPTrim$(frm1.fpcomboBankdraft.ColText) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboBankdraft.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPDDACC) <> QPTrim$(frm1.txtBankAcctNo.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.txtBankAcctNo.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.PRENOTED) <> QPTrim$(frm1.fpcomboPrenoted.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboPrenoted.SetFocus
      GoTo ChangeTrue
    End If
    
    If QPTrim$(EmpData2FileRec.BankName) <> QPTrim$(frm1.txtBankName.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.txtBankName.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.BANKLOC) <> QPTrim$(frm1.txtBankLocation.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.txtBankLocation.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.TRANSIT) <> QPTrim$(frm1.txtBankTransNo.Text) Then
      If EmpData2FileRec.TRANSIT = 0 Then GoTo TransitIs0
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 1
      frm1.vaTabPro1.SetFocus
      frm1.txtBankTransNo.SetFocus
      GoTo ChangeTrue
    End If
TransitIs0:
    If QPTrim$(EmpData2FileRec.EMPJOB) <> QPTrim$(frm1.txtTitle.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.txtTitle.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPWCCLS) <> QPTrim$(frm1.fptxtWCCode.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtWCCode.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPSTATS) <> QPTrim$(frm1.fpcomboStatus.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboStatus.SetFocus
      GoTo ChangeTrue
    End If
    
    If EmpData2FileRec.EMPBCODE <> Val(ReplaceString(frm1.fptxtBenefitPct.Text, "%", "")) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtBenefitPct.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPPTYPE) <> QPTrim$(frm1.fpcomboPayType.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboPayType.SetFocus
      GoTo ChangeTrue
    End If
    If QPTrim$(EmpData2FileRec.EMPPFREQ) <> QPTrim$(frm1.fpcomboFreq.Text) Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fpcomboFreq.SetFocus
      GoTo ChangeTrue
    End If
      
    If EmpData2FileRec.EMPPRATE <> frm1.fptxtRate.Text Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtRate.SetFocus
      GoTo ChangeTrue
    End If
    If EmpData2FileRec.EMPORATE <> frm1.fptxtOTRate.Text Then
      changeFlag = True
      frm1.vaTabPro1.ActiveTab = 2
      frm1.vaTabPro1.SetFocus
      frm1.fptxtOTRate.SetFocus
      GoTo ChangeTrue
    End If
    If newEmpFlag = True Then
      If EmpData2FileRec.EMPHDATE <> 0 Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 2
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskHire.SetFocus
          GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPRDATE <> 0 Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 2
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskNext.SetFocus
          GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPTDATE <> 0 Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 2
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskTerm.SetFocus
          GoTo ChangeTrue
      End If
    Else 'not a new employee entry
      If EmpData2FileRec.EMPHDATE < -21914 Then GoTo HireDate0 ' -21914 is 01/01/1920
      'eliminates DOS empty field code from trapping
      If EmpData2FileRec.EMPHDATE <> 0 And frm1.fpMaskHire.Text = "" Then
        changeFlag = True 'happens if an existing date field is deleted
        frm1.vaTabPro1.ActiveTab = 2
        frm1.vaTabPro1.SetFocus
        frm1.fpMaskHire.SetFocus
        GoTo ChangeTrue
      ElseIf EmpData2FileRec.EMPHDATE = 0 And frm1.fpMaskHire.Text = "" Then
        GoTo HireDate0
      ElseIf EmpData2FileRec.EMPHDATE <> DateDiff("d", "12/31/1979", frm1.fpMaskHire.Text) Then
        changeFlag = True 'no need for empty string/0 comparison because this is a required field
        frm1.vaTabPro1.ActiveTab = 2
        frm1.vaTabPro1.SetFocus
        frm1.fpMaskHire.SetFocus
        GoTo ChangeTrue
      End If
HireDate0:
      If EmpData2FileRec.EMPRDATE < -21914 Then GoTo NextDate0 ' -21914 is 01/01/1920
      If EmpData2FileRec.EMPRDATE = 0 And frm1.fpMaskNext.Text = "" Then
         GoTo NextDate0
      ElseIf EmpData2FileRec.EMPRDATE <> 0 And frm1.fpMaskNext.Text = "" Then
         changeFlag = True 'occurs if an existing date field is deleted
         frm1.vaTabPro1.ActiveTab = 2
         frm1.vaTabPro1.SetFocus
         frm1.fpMaskNext.SetFocus
         GoTo ChangeTrue
      ElseIf EmpData2FileRec.EMPRDATE <> DateDiff("d", "12/31/1979", frm1.fpMaskNext.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 2
         frm1.vaTabPro1.SetFocus
         frm1.fpMaskNext.SetFocus
         GoTo ChangeTrue
      End If
NextDate0:
      If EmpData2FileRec.EMPTDATE < -21914 Then GoTo TermDate0 ' -21914 is 01/01/1920
      If EmpData2FileRec.EMPTDATE = 0 And frm1.fpMaskTerm.Text = "" Then
         GoTo TermDate0
      ElseIf EmpData2FileRec.EMPTDATE <> 0 And frm1.fpMaskTerm.Text = "" Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 2
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskTerm.SetFocus
          GoTo ChangeTrue
      ElseIf EmpData2FileRec.EMPTDATE <> DateDiff("d", "12/31/1979", frm1.fpMaskTerm.Text) Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 2
          frm1.vaTabPro1.SetFocus
          frm1.fpMaskTerm.SetFocus
          GoTo ChangeTrue
      End If
    End If
TermDate0:
    
    'added 9/1/04----------
    If QPTrim$(EmpData2FileRec.Comment) <> QPTrim$(frm1.fptxtComment.Text) Then
       changeFlag = True
       frm1.vaTabPro1.ActiveTab = 2
       frm1.vaTabPro1.SetFocus
       frm1.fptxtComment.SetFocus
       GoTo ChangeTrue
    End If
    'added 9/1/04^^^^^^^^^^^^

      If QPTrim$(EmpData2FileRec.EMPFEDX) <> QPTrim$(frm1.fpcomboFedX.Text) Then
      'if it has then set the changeFlag to 1 and reset focus to
      'where the change was made
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboFedX.SetFocus
         GoTo ChangeTrue
      End If
      If QPTrim$(EmpData2FileRec.EMPFEDO2) <> QPTrim$(frm1.fpcomboFedAmtPct.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboFedAmtPct.SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPFEDO1 <> Val(frm1.fptxtFedFig.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fptxtFedFig.SetFocus
         GoTo ChangeTrue
      End If
      If QPTrim$(EmpData2FileRec.EMPFEDS) <> QPTrim$(frm1.fpcomboFedStatus.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboFedStatus.SetFocus
         GoTo ChangeTrue
      End If
   
      If EmpData2FileRec.EMPFEDA <> Val(frm1.fptxtAllowNumFed.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAllowNumFed.SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPFEDAA <> frm1.fptxtAddWHFed.Text Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAddWHFed.SetFocus
         GoTo ChangeTrue
      End If
      If QPTrim$(EmpData2FileRec.EMPSTAX) <> QPTrim$(frm1.fpcomboStateX.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboStateX.SetFocus
         GoTo ChangeTrue
      End If
   
      If QPTrim$(EmpData2FileRec.EMPSTAO2) <> QPTrim$(frm1.fpcomboStateAmtPct.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboStateAmtPct.SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPSTAO1 <> Val(frm1.fptxtStateFig.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fptxtStateFig.SetFocus
         GoTo ChangeTrue
      End If
      If QPTrim$(EmpData2FileRec.EMPSTAS) <> QPTrim$(frm1.fpcomboStateStatus.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboStateStatus.SetFocus
         GoTo ChangeTrue
      End If
    
      If EmpData2FileRec.EMPSTAA <> Val(frm1.fptxtAllowNumState.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAllowNumState.SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPSTAAA <> frm1.fptxtAddWHState.Text Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAddWHState.SetFocus
         GoTo ChangeTrue
      End If
      If QPTrim$(EmpData2FileRec.EMPSOCX) <> QPTrim$(frm1.fpcomboSocX.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboSocX.SetFocus
         GoTo ChangeTrue
      End If
   
      If QPTrim$(EmpData2FileRec.EMPMEDX) <> QPTrim$(frm1.fpcomboMedX.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboMedX.SetFocus
         GoTo ChangeTrue
      End If
      
      frm1.fpcomboEIC.Col = 0
      If QPTrim$(EmpData2FileRec.EMPEIC) <> QPTrim$(frm1.fpcomboEIC.ColText) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 3
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboEIC.SetFocus
         GoTo ChangeTrue
      End If
   For x = 1 To 50
       frm1.vaSpreadMisc.Col = 2
       frm1.vaSpreadMisc.Row = x
       If QPTrim$(EmpData2FileRec.EmpDed(x).DPct) <> QPTrim$(frm1.vaSpreadMisc.Text) Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 4
          frm1.vaSpreadMisc.SetFocus
          frm1.vaSpreadMisc.SetActiveCell 2, x
          GoTo ChangeTrue
       End If
       frm1.vaSpreadMisc.Col = 3
       frm1.vaSpreadMisc.Row = x
       'the > 0 is used to get past the negative numbers
       'being loaded from the DOS files
       If EmpData2FileRec.EmpDed(x).DAmt <> Val(frm1.vaSpreadMisc.Text) Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 4
          frm1.vaSpreadMisc.SetFocus
          frm1.vaSpreadMisc.SetActiveCell 3, x
          GoTo ChangeTrue
       End If
       frm1.vaSpreadMisc.Col = 4
       frm1.vaSpreadMisc.Row = x
       If QPTrim$(frm1.vaSpreadMisc.Text) = "YES" Then frm1.vaSpreadMisc.Text = "Y"
       If QPTrim$(frm1.vaSpreadMisc.Text) = "NO" Then frm1.vaSpreadMisc.Text = "N"
       If QPTrim$(EmpData2FileRec.EmpDed(x).DOTI) <> QPTrim$(frm1.vaSpreadMisc.Text) Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 4
          frm1.vaSpreadMisc.SetFocus
          frm1.vaSpreadMisc.SetActiveCell 4, x
          GoTo ChangeTrue
       End If
   Next x

      
'need to add 2 more to the current 3
      If QPTrim$(EmpData2FileRec.EMPEACT1) <> QPTrim$(frm1.fptxtAN(1).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 5
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAN(1).SetFocus
         GoTo ChangeTrue
      End If
      
      If EmpData2FileRec.EMPEAMT1 <> frm1.fptxtE(1).Text Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 5
         frm1.vaTabPro1.SetFocus
         frm1.fptxtE(1).SetFocus
         GoTo ChangeTrue
      End If
      If QPTrim$(EmpData2FileRec.EMPEACT2) <> QPTrim$(frm1.fptxtAN(2).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 5
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAN(2).SetFocus
         GoTo ChangeTrue
      End If
      
      If EmpData2FileRec.EMPEAMT2 <> frm1.fptxtE(2).Text Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 5
         frm1.vaTabPro1.SetFocus
         frm1.frm1.fptxtE(2).SetFocus
         GoTo ChangeTrue
      End If
      
      If QPTrim$(EmpData2FileRec.EMPEACT3) <> QPTrim$(frm1.fptxtAN(3).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 5
         frm1.vaTabPro1.SetFocus
         frm1.fptxtAN(3).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPEAMT3 <> frm1.fptxtE(3).Text Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 5
         frm1.vaTabPro1.SetFocus
         frm1.fptxtE(3).SetFocus
         GoTo ChangeTrue
      End If

      EmpData2FileRec.EMPHP = ""
'need to add 2 more to the current 8
      For x = 1 To 8
         If QPTrim$(EmpData2FileRec.EDist(x).DAcct) <> QPTrim$(frm1.fptxtWDAN(x).Text) Then
            changeFlag = True
            frm1.vaTabPro1.ActiveTab = 6
            frm1.vaTabPro1.SetFocus
            frm1.fptxtWDAN(x).SetFocus
         End If
         If EmpData2FileRec.EDist(x).DAmt <> Val(frm1.fptxtWDDD(x).Text) Then
            changeFlag = True
            frm1.vaTabPro1.ActiveTab = 6
            frm1.vaTabPro1.SetFocus
            frm1.fptxtWDDD(x).SetFocus
         End If
      Next x
NoMoreDAcct:
      If EmpData2FileRec.EMPVACE <> Val(frm1.fptxtEarned(1).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtEarned(1).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPVUSED <> Val(frm1.fptxtUsed(1).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtUsed(1).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPSLE <> Val(frm1.fptxtEarned(2).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtEarned(2).SetFocus
         GoTo ChangeTrue
      End If
      
      If EmpData2FileRec.EMPSLUSE <> Val(frm1.fptxtUsed(2).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtUsed(2).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPCTE <> Val(frm1.fptxtEarned(3).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtEarned(3).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.EMPCTUSE <> Val(frm1.fptxtUsed(3).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtUsed(3).SetFocus
         GoTo ChangeTrue
      End If
      
      If EmpData2FileRec.PERERN <> Val(frm1.fptxtEarned(4).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtEarned(4).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.PerUsed <> Val(frm1.fptxtUsed(4).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtUsed(4).SetFocus
         GoTo ChangeTrue
      End If
      
      If EmpData2FileRec.HOLERN <> Val(frm1.fptxtEarned(5).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtEarned(5).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.HolUsed <> Val(frm1.fptxtUsed(5).Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fptxtUsed(5).SetFocus
         GoTo ChangeTrue
      End If
      If EmpData2FileRec.LeaveTbl <> Val(frm1.fpcomboLT.Text) Then
         changeFlag = True
         frm1.vaTabPro1.ActiveTab = 7
         frm1.vaTabPro1.SetFocus
         frm1.fpcomboLT.SetFocus
      End If
      
      '***********added 11/12/2002*************
      If newEmpFlag = True Then
        If QPTrim$(frm1.fpcombo401K.Text) <> "N" Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 7
          frm1.vaTabPro1.SetFocus
          frm1.fpcombo401K.SetFocus
        End If
      ElseIf newEmpFlag = False Then
        If QPTrim$(EmpData2FileRec.YN401K) <> QPTrim$(frm1.fpcombo401K.Text) Then
           changeFlag = True
           frm1.vaTabPro1.ActiveTab = 7
           frm1.vaTabPro1.SetFocus
           frm1.fpcombo401K.SetFocus
        End If
      End If
      '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
      
      If newEmpFlag = True Then
        If QPTrim$(frm1.fpcomboESC.Text) <> "N" Then
          changeFlag = True
          frm1.vaTabPro1.ActiveTab = 7
          frm1.vaTabPro1.SetFocus
          frm1.fpcomboESC.SetFocus
        End If
      ElseIf newEmpFlag = False Then
        If QPTrim$(EmpData2FileRec.ExcludeESC) <> QPTrim$(frm1.fpcomboESC.Text) Then
           changeFlag = True
           frm1.vaTabPro1.ActiveTab = 7
           frm1.vaTabPro1.SetFocus
           frm1.fpcomboESC.SetFocus
        End If
      End If
'Change handling procedure_______________________________________________
ChangeTrue:
      If changeFlag = False Then 'no changes detected
        If newEmpFlag = True Then '7/25/03
          frmEmployeeMaintMenu.Show
        Else
          Call frmEmployeeLookUp.ActivateControls '7/25/03
        End If
        DoEvents
         Unload frmEditEmpData
         GoTo endClick
      'if a change was made then bring up a warning window that forces
      'the user to decide whether to save, review or abandon changes
      Else
         DoWhatFlag = PromptSaveChanges(frm1)
         Select Case DoWhatFlag
         Case SaveChangeOptions1.scoSaveChanges 'save changes
            Call SaveEmpInfo(newEmpFlag, thisRecordNum, frm1)
         Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
         Case SaveChangeOptions1.scoAbandonChanges 'abandon
           If newEmpFlag = False Then
'              frmEmployeeLookUp.Show 'commented out 7/25/03
              Call frmEmployeeLookUp.ActivateControls
              DoEvents
              Unload frmEditEmpData
           Else
              frmEmployeeMaintMenu.Show
              DoEvents
              Unload frmEditEmpData
           End If
         Case Else:
           'Do nothing because we don't know about any options except
           'save, review or abandon...used as a placeholder for adding
           'other options at a later date
         End Select
         
      End If
      GoTo endClick
endClick:

End Sub




