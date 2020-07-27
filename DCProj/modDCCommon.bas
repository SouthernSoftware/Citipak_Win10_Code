Attribute VB_Name = "modDCCommon"
Option Explicit
'!!! Procedures below Needed for reports!!! Mark with!!!
'!!! Added Round on 4-17-03
Public Function Round#(ByVal n#)
  Round# = (Int(n# * 100 + 0.5000001)) / 100
End Function

'!!! from gl common for date check on report screens
Public Function CheckValDate(ValCheck As String)
  Dim Month As Integer, Day As Integer, year As Integer
  Month = Val(Mid(ValCheck, 1, 2))
  Day = Val(Mid(ValCheck, 4, 2))
  year = Val(Mid(ValCheck, 7, 4))
  'Checks date if Blank then won't check for valid date
  'and then checks each section, month, day and year
  'if any section wrong then returns false value
  If InStr(ValCheck, "_") <= 0 Then
    If ((Month > 0) And (Month < 13)) Then
      If Day > 0 And Day < 32 Then
        If year > 1979 And year < 2099 Then
          CheckValDate = True
        End If
      End If
    End If
  End If
End Function
Public Function FillCatCMBO(x As fpCombo)
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumOFDCCatRecs As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  ghandle = FreeFile
  Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(ghandle) \ DCCatCodeRecLen
  x.Row = 0
  For cnt = 1 To NumOFDCCatRecs
    Get #ghandle, cnt, DCCatCodeRec
    If DCCatCodeRec.InactiveFlag <> "Y" Then
      x.AddItem Str$(cnt) & Chr$(9) & QPTrim$(DCCatCodeRec.CATCODE) & Chr$(9) & DCCatCodeRec.CODEDESC
    Else
      x.AddItem Str$(cnt) & Chr$(9) & QPTrim$(DCCatCodeRec.CATCODE) & Chr$(9) & "Inactivated Code"
    End If
  Next
  Close
End Function
Public Function FillCatCMBOAll(x As fpCombo)
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumOFDCCatRecs As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  ghandle = FreeFile
  Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(ghandle) \ DCCatCodeRecLen
  x.Row = 0
  x.AddItem Str$(0) & Chr$(9) & "All" & Chr$(9) & "All Codes"
  x.Row = 1
  For cnt = 1 To NumOFDCCatRecs
    Get #ghandle, cnt, DCCatCodeRec
    If DCCatCodeRec.InactiveFlag <> "Y" Then
      x.AddItem Str$(cnt) & Chr$(9) & DCCatCodeRec.CATCODE & Chr$(9) & DCCatCodeRec.CODEDESC
    Else
      x.AddItem Str$(cnt) & Chr$(9) & DCCatCodeRec.CATCODE & Chr$(9) & "Inactivated Code"
    End If
  Next
  Close
End Function
Public Sub GetUBBankINfo()
  Dim CMBnkAcct As CMBankAcctRecType
  Dim CMBnkAcctLen As Integer, CMFile As Integer
  On Local Error GoTo ubb
  CMBnkAcctLen = Len(CMBnkAcct)
  CMFile = FreeFile
  Open DCPath + "CMBkAcct.DAT" For Random Shared As CMFile Len = CMBnkAcctLen
  Get CMFile, 1, CMBnkAcct
    BnkAcctNum$ = QPTrim$(CMBnkAcct.COMPACCT)
  Close
  Exit Sub
ubb:
   BnkAcctNum$ = " "
End Sub

Public Sub RelinkDCStuff(formname As Form)
  Dim DCCustLen As Integer, DCVehLen As Integer, DCTranLen As Integer, vcnt As Long
  Dim CFile As Integer, VFile As Integer, TFile As Integer, CustRec As Long
  Dim cnt As Long, NumOfCust As Long, NumOfTran As Long, NumOfVeh As Long, VTfile As Integer
    ReDim DCCustRec(1 To 2) As DCCustRecType
    ReDim DCVehRec(1 To 2) As DCVehType
    ReDim DCTrans(1 To 2) As DCTransRecType
    ReDim TempDCV(1 To 2) As DCVehType
   FrmShowPctComp.Label1 = "Clearing customer links"
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
    vcnt = 0
    DCCustLen = Len(DCCustRec(1))
    DCVehLen = Len(DCVehRec(1))
    DCTranLen = Len(DCTrans(1))

    CFile = FreeFile
    Open "DCCust.dat" For Random Shared As CFile Len = DCCustLen
    NumOfCust& = LOF(CFile) / DCCustLen

    VFile = FreeFile
    Open "DCVEH.dat" For Random Shared As VFile Len = DCVehLen
    NumOfVeh& = LOF(VFile) / DCVehLen
    
    VTfile = FreeFile
    Open "DCVTmp.dat" For Random Shared As VTfile Len = DCVehLen
    
    For cnt& = 1 To NumOfCust&
      FrmShowPctComp.ShowPctComp cnt, NumOfCust
      Get CFile, cnt&, DCCustRec(1)
   ''''If cnt& = 142 > NumOfCust& Then Stop
      DCCustRec(1).FirstCar = 0
      DCCustRec(1).LastCar = 0
      Put CFile, cnt&, DCCustRec(1)
 '    ShowPctCompL cnt&, NumOfCust&
    Next
   FrmShowPctComp.Label1 = "Clearing Vehicle links"
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
    For cnt& = 1 To NumOfVeh&
      FrmShowPctComp.ShowPctComp cnt, NumOfVeh
      Get VFile, cnt&, DCVehRec(1)
      If DCVehRec(1).Active <> "N" Then
        TempDCV(1).NextRec = 0
        TempDCV(1).DecalCat = DCVehRec(1).DecalCat
        TempDCV(1).makemodel = DCVehRec(1).makemodel
        TempDCV(1).StateTag = DCVehRec(1).StateTag
        TempDCV(1).ExpireDate = DCVehRec(1).ExpireDate
        TempDCV(1).Sticker = DCVehRec(1).Sticker
        TempDCV(1).Valid = DCVehRec(1).Valid
        TempDCV(1).Active = "Y"
        TempDCV(1).Notes = DCVehRec(1).Notes
        TempDCV(1).PBFlag = DCVehRec(1).PBFlag
        TempDCV(1).Desc = DCVehRec(1).Desc
        TempDCV(1).Fee = DCVehRec(1).Fee
        TempDCV(1).MasterRecord = DCVehRec(1).MasterRecord
        TempDCV(1).MoreRoom = ""
        vcnt = vcnt + 1
        Put VTfile, vcnt, TempDCV(1)
      End If
    Next
    Close
    'Close VFile
    'Close VTfile
    DoEvents
    If FileSize("DCVTmp.dat") > 0 Then
      KillFile "DCVEH.dat"
      Name "DCVTmp.dat" As "DCVEH.dat"
    End If
    CFile = FreeFile
    Open "DCCust.dat" For Random Shared As CFile Len = DCCustLen
    NumOfCust& = LOF(CFile) / DCCustLen


    VFile = FreeFile
    Open "DCVEH.dat" For Random Shared As VFile Len = DCVehLen
    NumOfVeh& = LOF(VFile) / DCVehLen
   FrmShowPctComp.Label1 = "Relinking Vehicles"
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
    For cnt& = 1 To NumOfVeh&
      FrmShowPctComp.ShowPctComp cnt, NumOfVeh
      Get VFile, cnt&, DCVehRec(1)
     ''' If cnt& = 194 Then Stop
      CustRec& = DCVehRec(1).MasterRecord
'''''      If CustRec& = 1033 Then Stop
'      If Len(QPTrim$(DCVehRec(1).DecalCat)) <= 0 Then
'        GoTo SkipVeh
'      End If   Or QPTrim$(DCVehRec(1).Active) = "N"
      If CustRec& <= 0 Or CustRec& > NumOfCust& Then
        GoTo SkipVeh
      End If
      If DCVehRec(1).Active = "N" Then
        GoTo SkipVeh
      End If
      Get CFile, CustRec&, DCCustRec(1)
   '''here dont forget to rem this out @)#*@)#*)#*)(@#*)#*)@(#*)#*
     '''If CustRec& = 142 Then Stop

      If DCCustRec(1).FirstCar = 0 Then  'if the first car
        DCCustRec(1).FirstCar = cnt&
        DCCustRec(1).LastCar = cnt&
        Put CFile, CustRec&, DCCustRec(1)
      Else                               'nope not first
        Get VFile, DCCustRec(1).LastCar, DCVehRec(2)  'get old last car
        DCVehRec(2).NextRec = cnt&       'set next car in old car
        Put VFile, DCCustRec(1).LastCar, DCVehRec(2) 'put old last car back
        DCCustRec(1).LastCar = cnt&      'set new last car in cust
        Put CFile, CustRec&, DCCustRec(1) 'put cust back
      End If
SkipVeh:
  Next
  Close VFile
    TFile = FreeFile
    Open "DCTRANS.dat" For Random Shared As TFile Len = DCTranLen
    NumOfTran& = LOF(TFile) / DCTranLen
   FrmShowPctComp.Label1 = "Clearing Transaction links"
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
    For cnt& = 1 To NumOfCust&
      FrmShowPctComp.ShowPctComp cnt, NumOfCust
      Get CFile, cnt&, DCCustRec(1)
      DCCustRec(1).FirstTrans = 0
      DCCustRec(1).LastTrans = 0
      Put CFile, cnt&, DCCustRec(1)
     ' ShowPctCompL cnt&, NumOfCust&
    Next
   FrmShowPctComp.Label1 = "Clearing Transaction links"
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
    For cnt& = 1 To NumOfTran&
     FrmShowPctComp.ShowPctComp cnt, NumOfTran
      Get TFile, cnt&, DCTrans(1)
      DCTrans(1).NextTrans = 0
      Put TFile, cnt&, DCTrans(1)
    '  ShowPctCompL cnt&, NumOfTran&
    Next
   FrmShowPctComp.Label1 = "Relinking Transactions"
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
    For cnt& = 1 To NumOfTran&
      FrmShowPctComp.ShowPctComp cnt, NumOfTran
      Get TFile, cnt&, DCTrans(1)
      CustRec& = Val(QPTrim$(DCTrans(1).CustomerNumber))
      If CustRec& <= 0 Or CustRec& > NumOfCust& Then
        GoTo SkipTrans
      End If
      Get CFile, CustRec&, DCCustRec(1)
      If DCCustRec(1).FirstTrans = 0 Then  'if the first trans
        DCCustRec(1).FirstTrans = cnt&
        DCCustRec(1).LastTrans = cnt&
           ''''''If CustRec& <= 0 Or CustRec& > NumOfTran& Then Stop
   
        Put CFile, CustRec&, DCCustRec(1)
      Else                                 'not first trans
        Get TFile, DCCustRec(1).LastTrans, DCTrans(2)  'get old last trans
        DCTrans(2).NextTrans = cnt&
        
        Put TFile, DCCustRec(1).LastTrans, DCTrans(2)  'put old last back

        DCCustRec(1).LastTrans = cnt&      'set new last
                  ''''' If CustRec& <= 0 Or CustRec& > NumOfTran& Then Stop

        Put CFile, CustRec&, DCCustRec(1) 'put cust back

      End If
     ' ShowPctCompL cnt&, NumOfTran&
SkipTrans:
    Next
 
    Close
  MsgBox "RELINK COMPLETE", vbOKOnly, "Procedure Complete"
ExitRelink:

End Sub
  
Public Function GetZipEDigit$(Zip$)
  Dim ZipLen As Integer, ZipVal As Integer, DashPos As Integer
  Dim cnt As Integer, Dif As Double
  ZipLen = Len(Zip$)
  ZipVal = 0

  DashPos = InStr(Zip$, "-")
  Do While DashPos
    Zip$ = Left$(Zip$, DashPos - 1) + Mid$(Zip$, DashPos + 1)
    DashPos = InStr(Zip$, "-")
  Loop

  For cnt = 1 To ZipLen
    ZipVal = ZipVal + Val(Mid$(Zip$, cnt, 1))
  Next

  If ZipVal Mod 10 > 0 Then
    Dif = 10 - (ZipVal Mod 10)
  Else
    Dif = 0
  End If
  GetZipEDigit$ = QPTrim$(Str$(Dif))

End Function

Public Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  If intDate% = -32767 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function

Public Function Date2Num%(txtDate$)
  On Error GoTo BadDate2Num
  If Len(QPTrim$(txtDate$)) = 10 Then
    Date2Num% = DateDiff("d", "12/31/1979", txtDate$)
  Else
    Date2Num% = -32767
  End If
  Exit Function

BadDate2Num:
  On Error GoTo 0
  Date2Num% = -32767
End Function

Public Sub LoadDCSetUpFile(dcSetUpRec() As DCSetupType, DCSetuplen)
  Dim Handle As Integer
  DCSetuplen = Len(dcSetUpRec(1))            'use the length as an error flag
  If Exist(DCPath$ + "DCSetUP.dat") Then
    Handle = FreeFile
    Open DCPath$ + "DCSetUP.dat" For Random Shared As Handle Len = DCSetuplen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, dcSetUpRec(1)
    End If
    Close Handle
  End If
End Sub
'Public Sub LoadUBBillSetUpFile(UBBillSetUpRec() As UBBillSetupType, UBBillSetuplen)
'  Dim Handle As Integer
'  UBBillSetuplen = Len(UBBillSetUpRec(1))            'use the length as an error flag
'  If Exist(UBPath$ + "UBBilSET.DAT") Then
'    Handle = FreeFile
'    Open UBPath$ + "UBBilSET.DAT" For Random Shared As Handle Len = UBBillSetuplen    'open data file
'    If LOF(Handle) > 0 Then
'      Get #Handle, 1, UBBillSetUpRec(1)
'    End If
'    Close Handle
'  End If
'End Sub
'Public Sub LoadUBBillLetterFile(UBBillLetterRec() As UBBillLetterType, UBBillLetterlen)
'  Dim Handle As Integer
'  UBBillLetterlen = Len(UBBillLetterRec(1))            'use the length as an error flag
'  If Exist(UBPath$ + "UBBilLtr.DAT") Then
'    Handle = FreeFile
'    Open UBPath$ + "UBBilLtr.DAT" For Random Shared As Handle Len = UBBillLetterlen    'open data file
'    If LOF(Handle) > 0 Then
'      Get #Handle, 1, UBBillLetterRec(1)
'    End If
'    Close Handle
'  End If
'End Sub
Public Function Exist(FileName$)
  On Local Error Resume Next
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
'  If Err Then
'    FileName$ = ""
'  End If
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    Kill FileName$
  End If
'  On Local Error GoTo 0
End Function

Public Sub KillFile(FileName$)
  If Exist(FileName) Then
    Kill FileName$
  End If
End Sub
Public Function FileSize(FileName$) As Long
  Dim FileHandle As Integer
  If Exist(FileName$) Then
    FileHandle = FreeFile
    Open FileName$ For Binary As FileHandle
    FileSize = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize = 0
  End If
End Function

Public Function RemNulls$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  RemNulls$ = Text
End Function

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function
Public Function QPStripCom$(Address$)
  Dim x As String, StrLen As Long, cnt As Long, thischar As String
  x$ = QPTrim$(Address$)
  StrLen = Len(x$)
  For cnt = 1 To StrLen
    thischar = Mid$(x$, cnt, 1)
    If thischar = "," Then
      Mid$(x$, cnt, 1) = " "
    End If
  Next

  QPStripCom$ = Trim$(x$)

End Function
Public Function QPStripLast$(NM$)
  Dim x As String, StrLen As Long, cnt As Long, thischar As String
  x$ = QPTrim$(NM$)
  StrLen = Len(x$)
  For cnt = 1 To StrLen
    thischar = Mid$(x$, (StrLen + 1) - cnt, 1)
    If thischar = " " Then
      x$ = Right$(x$, cnt)
      Exit For
    End If
  Next

  QPStripLast$ = Trim$(x$)

End Function

Public Function QPStripStuff$(Temp$)
  Dim x As String, StrLen As Long, cnt As Long, thischar As String, newcnt As Long
  Dim xx As String
  x$ = QPTrim$(Temp$)
  xx$ = ""
  newcnt = 0
  StrLen = Len(x$)
  For cnt = 1 To StrLen
    thischar = Mid$(x$, cnt, 1)
    If thischar = "(" Or thischar = ")" Or thischar = "-" Then
      Mid$(x$, cnt, 1) = " "
    End If
  Next
  For cnt = 1 To StrLen
    thischar = Mid$(x$, cnt, 1)
    If thischar = " " Then
      xx$ = xx$ + ""
    Else
      xx$ = xx$ + thischar
    End If
  Next
  QPStripStuff$ = Trim$(xx$)

End Function

Public Static Function Using$(ByVal fmt As String, ByVal Number As Double, Optional LeadZeroFlag As Boolean)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(Len(fmt))
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen = 0 Then
    TempNumber = "0"
    GoTo GotZero
  End If
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
GotZero:
  If LeadZeroFlag Then
    If TempNumber = ".00" Then
      TempNumber = "0.00"
    End If
  End If
  
  RSet FmtNumber = TempNumber
  
  Using = FmtNumber
  
'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function


Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String, Optional HideF7btn As Boolean)
  frmLoadingRpt.Show
  DoEvents
  frmViewPrint.ReportName = ReportFile$
  frmViewPrint.Caption = Title
  frmViewPrint.PgNum = PgNum
  If ForceSBar Then
    frmViewPrint.fpMemo1.ScrollBars = BothFixed
  Else
    frmViewPrint.fpMemo1.ScrollBars = BothAuto
  End If
  If Algn Then
    frmViewPrint.cmdAlignment.Enabled = True
    frmViewPrint.AlignRpt = AlgnRptfile$
  Else
    frmViewPrint.cmdAlignment.Enabled = False
  End If
  If HideF7btn Then
    frmViewPrint.cmdPrnScn.Enabled = False
  End If
  DoEvents
  Unload frmLoadingRpt
  DoEvents
  frmViewPrint.Show vbModal
End Sub
Public Sub ViewPrintM(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String, Optional HideF7btn As Boolean)
 ' frmLoadingRpt.Show 1
 'not using loadingrpt form only diff between this and regular viewprint
 'the problem was all modal forms
  DoEvents
  frmViewPrint.ReportName = ReportFile$
  frmViewPrint.Caption = Title
  frmViewPrint.PgNum = PgNum
  If ForceSBar Then
    frmViewPrint.fpMemo1.ScrollBars = BothFixed
  Else
    frmViewPrint.fpMemo1.ScrollBars = BothAuto
  End If
  If Algn Then
    frmViewPrint.cmdAlignment.Enabled = True
    frmViewPrint.AlignRpt = AlgnRptfile$
  Else
    frmViewPrint.cmdAlignment.Enabled = False
  End If
  If HideF7btn Then
    frmViewPrint.cmdPrnScn.Enabled = False
  End If
 ' DoEvents
 'Unload frmLoadingRpt
  DoEvents
  frmViewPrint.Show vbModal
End Sub


Public Sub DCTerminate()
  Dim DCFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  On Local Error Resume Next
  DCLog "DC Exited: "
  Ready4others PWcnt
  If DebugMode = False Then
    Shell "CitiPak.exe", vbMaximizedFocus
  End If
  DoTheTime
  DoEvents
  For DCFrmCnt = Forms.Count - 1 To 0 Step -1
    DoEvents
    Unload Forms(DCFrmCnt)
  Next
  End
End Sub

Public Sub CitiTerminate()
  Dim DCFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  ClearInUse PWcnt
  DoEvents
  For DCFrmCnt = Forms.Count - 1 To 0 Step -1
    Unload Forms(DCFrmCnt)
  Next
  DoEvents
  End
End Sub

Public Static Sub DCLog(Text$)
  Dim Today As String, TheTime As String
  Dim AmPm As String, Hour As String
  Dim ThisHour As Integer, LogFile As Integer
  
  Today$ = Date$
  Today$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  TheTime$ = Time$
  If Left$(TheTime$, 1) = "0" Then
    ThisHour = Val(Mid$(TheTime$, 2, 1))
  Else
    ThisHour = Val(Mid$(TheTime$, 1, 2))
  End If

  Select Case ThisHour
  Case Is > 11
    ThisHour = ThisHour - 12
    If ThisHour = 0 Then ThisHour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    Hour = 12
    AmPm$ = "am"
  End Select
  Select Case ThisHour
    Case 1 To 9
      Hour$ = "0" + QPTrim$(Str$(ThisHour))
    Case Else
      Hour$ = QPTrim$(Str$(ThisHour))
  End Select
  TheTime$ = Hour$ + ":" + Mid$(TheTime$, 4) + AmPm$
  LogFile = FreeFile
  Open DCPath$ + "DCLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "DC: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub
Public Function GetOKorNot%(MsgText() As String, Optional OKOnly As Boolean, Optional ByVal NoFlash As Boolean, Optional ByVal Add2Font As Integer)
  Dim zz As Integer, RetValue As Integer
  If OKOnly Then
    frmMsgDialog.RetLabel = "-2"
  End If
  frmMsgDialog.Caption = MsgText(0)
  For zz = 1 To 5
    frmMsgDialog.Label(zz - 1) = MsgText(zz)
    If Add2Font > 0 Then
      frmMsgDialog.Label(zz - 1).FontSize = frmMsgDialog.Label(zz - 1).FontSize + Add2Font
    End If
  Next
  If NoFlash Then
    frmMsgDialog.Timer1.Enabled = False
  End If
  frmMsgDialog.Show vbModal
  RetValue = Val(frmMsgDialog.RetLabel)
  Unload frmMsgDialog
  GetOKorNot% = RetValue
End Function
Public Function GetRPTName(Newrp As String)
  Dim Part As Double
  Part = Timer
  Newrp = Newrp + QPTrim(Str(CLng(Part)))
End Function
'Public Sub OpenCatCodeFile(NumOfCatRecs, CatFile)
'  ReDim DCCatCodeRec(1) As DCCatCodeRecType
'  CatCodeRecLen = Len(DCCatCodeRec(1))
'  CatFile = FreeFile
'  Open "CatCode.DAT" For Random As CatFile Len = CatCodeRecLen
'  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
'End Sub
Public Sub GetAcctStruct(GLUserName$, GLFundLen%, GLAcctLen%, GLDetLen%)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetupRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetupRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetupRec(1)
  Close SetupFile
  GLUserName = QPTrim$(GLSetupRec(1).UserName)
  GLFundLen = GLSetupRec(1).FundLen
  GLAcctLen = GLSetupRec(1).AcctLen
  GLDetLen = GLSetupRec(1).DetLen
  Erase GLSetupRec
End Sub
Public Static Function FillAcctNumName(txtField As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField.Row = -1
  txtField.InsertRow = Str$(0) & Chr$(9) & "Not Found" & Chr$(9) & "Invalid Account" & Chr$(9) & "0"
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
      If GLAcct.Deleted = 0 Then
        txtField.InsertRow = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.Title) & Chr$(9) & QPStrip(GLAcct.Num)
      End If
  Next
  Close AcctIdxFileNum
  Close AcctFile
  'Erase AcctIdxFileNum, NumAIdxRecs
  'Erase AcctFile, NumAccts, CntA
  End Function
Public Sub OpenAcctFile(AcctFileNum, Optional NumAccts As Integer)
  Dim GLAcctRecLen As Integer
  GLAcctRecLen = Len(GLAcct)
  AcctFileNum = FreeFile
  Open DCPath$ + "GLAcct.DAT" For Random Shared As AcctFileNum Len = GLAcctRecLen
  NumAccts = LOF(AcctFileNum) \ GLAcctRecLen
End Sub
Public Sub OpenAcctIdx(AcctIdxFileNum, NumAIdxRecs)
  Dim GLAcctIdxLen As Integer
  GLAcctIdxLen = Len(GLAcctidx)
  AcctIdxFileNum = FreeFile
  Open DCPath$ + "GLAcct.Idx" For Random Shared As AcctIdxFileNum Len = GLAcctIdxLen
  NumAIdxRecs = LOF(AcctIdxFileNum) \ GLAcctIdxLen
End Sub
Public Function GetNumCodeRecs%()
  Dim DCCodeRecLen As Integer
  ReDim DCCodeRec(1) As DCCatCodeRecType
  DCCodeRecLen = Len(DCCodeRec(1))
  GetNumCodeRecs = FileSize(DCPath + "DCCODE.DAT") \ DCCodeRecLen
  Erase DCCodeRec
End Function

Public Function QPStrip$(AcctNum$)
  Dim x As String, DashPos As Integer
   x$ = QPTrim$(AcctNum$)  '(Form$(AcctNum, 0))
   Do
      DashPos = InStr(x$, "-")
      If DashPos > 0 Then
         x$ = Left$(x$, DashPos - 1) + Mid$(x$, DashPos + 1)
      End If
    Loop While DashPos

    QPStrip$ = x$

End Function
Public Function DCCustCnt()
  Dim DCCustRecLen As Integer, DCFile As Integer, NumOfDCRecs As Long
  DCCustCnt = False

  ReDim tmpCustRec(1) As DCCustRecType
  DCCustRecLen = Len(tmpCustRec(1))

  DCFile = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As DCFile Len = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  Close DCFile

  DCCustCnt = NumOfDCRecs

  Erase tmpCustRec

End Function
'Function returns True if a customer has been deleted.
Public Function IsDeleted%(AcctNum&)
  ReDim DCCustRec(1) As DCCustRecType
  Dim Handle As Integer
  Dim DCCustRecLen As Integer
  
  DCCustRecLen = Len(DCCustRec(1))
  Handle = FreeFile
  Open "DCCUST.DAT" For Random Shared As Handle Len = DCCustRecLen
  Get Handle, AcctNum&, DCCustRec(1)
  Close Handle
  
  If UCase$(DCCustRec(1).Deleted) <> "Y" Then
    IsDeleted% = False
  Else
    IsDeleted% = True
  End If
  Erase DCCustRec

End Function

Public Sub OpenDCCustFile(NumOfDCRecs, DCFile)
  Dim DCCustRecLen As Integer
  Close DCFile
  ReDim DCCustRec(1) As DCCustRecType
  DCCustRecLen = Len(DCCustRec(1))
  DCFile = FreeFile
  Open "DCCUST.DAT" For Random Shared As DCFile Len = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  'FOR x = 1 TO NumOfDcRecs
  'GET DCFile, x, DCCust(1)
  'PRINT DCCust(1).Custnumb; TAB(15); DCCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
End Sub
Public Sub OpenDCCustIdxFile(NumOfDCIdxRecs, DCIdxFile)
  Dim DCCustIdxRecLen As Integer
  Close DCIdxFile
  ReDim DCCustIdxRec(1) As DCCustIDXRecType
  DCCustIdxRecLen = Len(DCCustIdxRec(1))
  DCIdxFile = FreeFile
  Open "DCCUST.IDX" For Random Access Read Write Shared As DCIdxFile Len = DCCustIdxRecLen
  NumOfDCIdxRecs = LOF(DCIdxFile) \ DCCustIdxRecLen
End Sub
 Public Sub PrintCustInfo(Rec As Long, RptType As Integer)
  Dim PageNo As Integer, Title As String, tb As Integer, dcnt As Long
  Dim Dash80 As String, ReportFile As String, Num1 As Long, Linecnt As Integer
  Dim DCRpt As Integer, ToPrint As String, TPDate As String
  Dim Msgflag As Boolean, RecNo As Long, NumOfVehs As Long, cnt As Long
  Dim NumOfDCRecs As Long, DCFile As Integer, GCode As String
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim Maxline As Integer
  Linecnt = 0
  ReDim DCCustRec(1) As DCCustRecType
  RecNo = Rec
  OpenDCCustFile NumOfDCRecs, DCFile
  Get DCFile, RecNo, DCCustRec(1)
  Close DCFile
  Title$ = "Customer Information Report"
  Dash80$ = String$(80, "-")
  TPDate$ = ""
  ReportFile$ = DCPath$ + "DCINFo.RPT"
  DCRpt = FreeFile
  Open ReportFile$ For Output As DCRpt
  ToPrint$ = ""
  MaxLines = 60
  If RptType = 1 Then 'do the graphics
  ToPrint$ = ""
  ToPrint$ = Str$(RecNo) + "~" + QPTrim$(DCCustRec(1).CUSTNUMB)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).SORTNAME)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).BILLNAME)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).ADDRESS1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).ADDRESS2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).City)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).State)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).ZIPCODE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).SOSEC)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).DRVLIC)
  ToPrint$ = ToPrint$ + "~" + Num2Date(DCCustRec(1).DATEOPED)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).HPHONE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustRec(1).WPHONE)
  Select Case DCCustRec(1).CASHONLY
    Case "N", " "
    ToPrint$ = ToPrint$ + "~" + "No"
  Case Else
    ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  If DCCustRec(1).FirstCar > 0 Then
    ToPrint$ = ToPrint$ + "~" + "Yes"
  Else
    ToPrint$ = ToPrint$ + "~" + "No"
  End If
  Select Case DCCustRec(1).resident
    Case "N", " "
      ToPrint$ = ToPrint$ + "~" + "No"
    Case Else
      ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  Select Case DCCustRec(1).Owner
    Case "N", " "
      ToPrint$ = ToPrint$ + "~" + "No"
    Case Else
      ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
   ToPrint$ = ToPrint$ + "~" + Using$("$###,###,###.##", DCCustRec(1).AcctBal)
  ReDim DCVRec(1) As DCVehType
  Num1 = DCCustRec(1).FirstCar
  If Num1 > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    cnt = Num1
    Do Until cnt = 0
    'For cnt = Num1 To Num2
    Get DCvFile, cnt, DCVRec(1)
    If DCVRec(1).Active = "Y" Then
      GCode$ = Str$(cnt)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).DecalCat)
      GCode$ = GCode$ + "~" + Using$("$###.##", DCVRec(1).Fee)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).makemodel)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).StateTag)
      GCode$ = GCode$ + "~" + Num2Date$(DCVRec(1).ExpireDate)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Sticker)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Valid)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Desc)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Notes)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).PBFlag)
      dcnt = dcnt + 1
      Print #DCRpt, ToPrint$ + "~" + GCode$
      GCode$ = ""
    
    End If
      cnt = DCVRec(1).NextRec
    Loop 'Next
    Close DCvFile
  End If
    If dcnt <= 0 Then
        GCode$ = "0~  ~ ~No Vehicles to Display ~No Vehicles ~ ~ ~ ~ ~ ~ "
        ToPrint$ = ToPrint$ + "~" + GCode$
        Print #DCRpt, ToPrint$
        
    End If
  
  
  Close
  Load frmLoadingRpt
  'frmLoadingRpt.setwherefrom frmUBCustMenu
  ARptCustInfo.txtDate = Now
  ARptCustInfo.txtTown = TOWNNAME$
  ARptCustInfo.GetName ReportFile$
  ARptCustInfo.startrpt
  Else
  Print #DCRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
  Print #DCRpt, Tab(30); Title$
  Print #DCRpt, Now
  Print #DCRpt, TOWNNAME$
  Print #DCRpt, Dash80$
  Print #DCRpt,
  Print #DCRpt, "Cust #: "; QPTrim$(DCCustRec(1).CUSTNUMB);
  Print #DCRpt, Tab(50); "Search Name: "; QPTrim$(DCCustRec(1).SORTNAME)
  Print #DCRpt, "Customer Name: "; QPTrim$(DCCustRec(1).BILLNAME);
  Print #DCRpt, Tab(50); "Date Opened: "; Num2Date(DCCustRec(1).DATEOPED)
  Print #DCRpt, "Address: "; QPTrim$(DCCustRec(1).ADDRESS1)
  Print #DCRpt, Tab(10); QPTrim$(DCCustRec(1).ADDRESS2);
  Print #DCRpt, Tab(50); "Account Balance: "; Using$("$###,###,###.##", DCCustRec(1).AcctBal)
  Print #DCRpt, Tab(10); QPTrim$(DCCustRec(1).City); " "; QPTrim$(DCCustRec(1).State); " "; QPTrim$(DCCustRec(1).ZIPCODE)
  Print #DCRpt,
  Print #DCRpt, "DriverLic#: "; QPTrim$(DCCustRec(1).DRVLIC);
  Select Case DCCustRec(1).CASHONLY
  Case "N", " "
    Print #DCRpt, Tab(50); "  Cash Only: "; "No"
  Case Else
    Print #DCRpt, Tab(50); "  Cash Only: "; "Yes"
  End Select
  Print #DCRpt, "SocSec#: "; QPTrim$(DCCustRec(1).SOSEC);
  Select Case DCCustRec(1).resident
  Case "N", " "
    Print #DCRpt, Tab(50); "Residential: "; "No"
  Case Else
    Print #DCRpt, Tab(50); "Residential: "; "Yes"
  End Select
  Print #DCRpt, "Home Phone: "; QPTrim$(DCCustRec(1).HPHONE);
  Select Case DCCustRec(1).Owner
  Case "N", " "
    Print #DCRpt, Tab(50); "      Owner: "; "No"
  Case Else
    Print #DCRpt, Tab(50); "      Owner: "; "Yes"
  End Select
  Print #DCRpt, "Work Phone: "; QPTrim$(DCCustRec(1).WPHONE);
  If DCCustRec(1).FirstCar > 0 Then
    Print #DCRpt, Tab(45); "Vehicles on File: "; "Yes"
  Else
    Print #DCRpt, Tab(45); "Vehicles on File: "; "No"
  End If
  
  Print #DCRpt,
  Print #DCRpt, "-------------------------- Vehicle Information ----------------------"
  Linecnt = 20
  ReDim DCVRec(1) As DCVehType
  Num1 = DCCustRec(1).FirstCar
  If Num1 > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    cnt = Num1
    Do Until cnt = 0
    Get DCvFile, cnt, DCVRec(1)
    If DCVRec(1).Active = "Y" Then
      If Linecnt >= MaxLines Then
          Print #DCRpt, FF$
          Print #DCRpt, Tab(30); Title$
          Print #DCRpt, Now
          Print #DCRpt, TOWNNAME$
          Print #DCRpt, Dash80$
          Print #DCRpt,
          Print #DCRpt, "Cust #: "; QPTrim$(DCCustRec(1).CUSTNUMB);
          Print #DCRpt, Tab(50); "Search Name: "; QPTrim$(DCCustRec(1).SORTNAME)
          Print #DCRpt, "Customer Name: "; QPTrim$(DCCustRec(1).BILLNAME)
          Print #DCRpt, "Continued  ---------------- Vehicle Information ----------------------"
          Linecnt = 8
        End If
      Print #DCRpt, "Category   Fee        Sticker#       Expires      Valid    P/B"
      Print #DCRpt, "--------   ---        --------       -------      -----    ---"
      Print #DCRpt, QPTrim$(DCVRec(1).DecalCat); Tab(12); Using$("$###.##", DCVRec(1).Fee);
      Print #DCRpt, Tab(22); QPTrim$(DCVRec(1).Sticker); Tab(36); Num2Date$(DCVRec(1).ExpireDate);
      Print #DCRpt, Tab(52); QPTrim$(DCVRec(1).Valid); Tab(60); QPTrim$(DCVRec(1).PBFlag)
      Print #DCRpt, "   Make/Model                            Vin#/Desc"
      Print #DCRpt, "   ----------                            ---------"
      Print #DCRpt, Tab(3); QPTrim$(DCVRec(1).makemodel); Tab(40); QPTrim$(DCVRec(1).Desc)
      Print #DCRpt, "   State Lic#                            Notes"
      Print #DCRpt, "   ----------                            -----"
      Print #DCRpt, Tab(3); QPTrim$(DCVRec(1).StateTag); Tab(40); QPTrim$(DCVRec(1).Notes)
      Linecnt = Linecnt + 9
      dcnt = dcnt + 1
    End If
      cnt = DCVRec(1).NextRec
    Loop 'Next
    Close DCvFile
  End If
    If dcnt <= 0 Then
        Print #DCRpt, "No Vehicles to Display**************"
        Linecnt = Linecnt + 1
    End If
 
  Print #DCRpt,
  Print #DCRpt, Dash80$
  Print #DCRpt, Chr$(12)

  Close

  ViewPrint ReportFile$, Title$
  KillFile ReportFile$
  End If
End Sub
Public Sub DisplayCustTransList(CustRec As Long)
  ReDim DCTranRec(1) As DCTransRecType
  ReDim DCCustRec(1) As DCCustRecType
  Dim DCCustRecLen As Integer, DCTranRecLen As Integer
  Dim PrevTranRec As Long
  Dim DCFile As Integer, dcnt As Integer
  Dim Build As String * 80
  Dim TType As String, TDesc As String
  Dim CurBal As Double
  
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents

  DCCustRecLen = Len(DCCustRec(1))
  DCTranRecLen = Len(DCTranRec(1))
  
  DCFile = FreeFile
  Open DCPath + "DCCust.dat" For Random Shared As DCFile Len = DCCustRecLen
  Get DCFile, CustRec&, DCCustRec(1)
  Close DCFile

  CurBal# = DCCustRec(1).AcctBal
'
Top:
'
  DCFile = FreeFile
  Open DCPath + "DCTRANS.DAT" For Random Shared As DCFile Len = DCTranRecLen
  
  PrevTranRec& = DCCustRec(1).FirstTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      dcnt = dcnt + 1
      Get DCFile, PrevTranRec&, DCTranRec(1)
      LSet Build = Str(DCTranRec(1).TransDate) + Chr9$ + " " + Num2Date(DCTranRec(1).TransDate)
      GoSub GetTransType
      Mid$(Build, 22) = TType$
      Mid$(Build, 53) = Using("#####.##", DCTranRec(1).TransAmount, True)
'      'this will show the actual trans number in the list
'      Mid$(Build, 55) = Str$(PrevTranRec&)
      Mid$(Build, 65) = Using("#####.##", DCTranRec(1).BalanceAfterTrans, True)
      Mid$(Build$, 73) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      frmTRDispListDC.fpTRList.AddItem Build$
      PrevTranRec& = DCTranRec(1).NextTrans
    Loop
  End If
  Close DCFile
  frmTRDispListDC.Label5.Caption = QPTrim(DCCustRec(1).BILLNAME)
  frmTRDispListDC.Label7 = "Acct: " & Str$(CustRec&)
  frmTRDispListDC.Label2 = "Balance: " + Using("#####.##", CurBal#, True)
  'frmTRDispListDC.Label3 = "Current: " + Using("#####.##", CurBal#, True)
  'frmTRDispListDC.Label4 = "Previous: " + Using("#####.##", PreBal#, True)
  Unload frmInfo
  DoEvents
  frmTRDispListDC.Show vbModal
  Erase DCTranRec, DCCustRec

Exit Sub

GetTransType:
'
  Select Case DCTranRec(1).TransType
  Case 1 'Charge
    TType$ = "Decal Charge"
  Case 2 'Payment
    TType$ = "Decal Payment"
  Case 3  'Charge Void
    TType$ = "Void Charge"
  Case 4  'Payment Void
    TType$ = "Void Payment"
  Case Else
    TType$ = Str$(DCTranRec(1).TransType) + " ???"
  End Select
  TDesc$ = QPTrim$(DCTranRec(1).TRVinDesc)
Return

End Sub
Public Sub PrintTRListScreen()
  Unload frmTRDispListDC
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "DCTRlist.rpt", "Customer Transaction List"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "DCTRlist.rpt"
    ARptLineRpt.startrpt
  End If
End Sub
Public Sub PrintTRDetlScreen()
  Unload frmTRDetailDC
  Unload frmTRDispListDC
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "DCTRDetl.RPT", "Customer Detail Transaction"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "DCTRDetl.RPT"
    ARptLineRpt.startrpt
  End If
End Sub

''Shenendoah
'   ExpDate$ = "06-30-0" + QPTrim$(Str$(Val(Right$(Date$, 1)) + 1))
'   IsuDate$ = "05-15-0" + QPTrim$(Str$(Val(Right$(Date$, 1))))
'
'   Print #RptHandle, Str$(DCCustIdxRec(1).IDXRECORD); Tab(35); DCCustRec(1).SOSEC
'   Print #RptHandle,
'   Print #RptHandle, Tab(65); IsuDate$
'   Print #RptHandle, Tab(65); ExpDate$
'   Print #RptHandle,
'   Print #RptHandle,
'
'  'PRINT #RptHandle, QPTrim$(LEFT$(DCVRec(1).Notes, 4));
'  Print #RptHandle, Tab(2); Left$(DCVRec(1).makemodel, 19);
'  Print #RptHandle, Tab(22); Left$(DCVRec(1).Desc, 30)
'  'PRINT #RptHandle,
'  'PRINT #RptHandle, TAB(70); USING "#######"; ControlNumber!
'
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(69); "15.00"
'  Print #RptHandle, Tab(69); " 8.00"
'  Print #RptHandle,
'  Print #RptHandle, Tab(69); " 1.00"
'  Print #RptHandle, ""
'  Print #RptHandle, Tab(44); Right$(IsuDate$, 2) + "-" + Right$(ExpDate$, 2)
'  Print #RptHandle, Tab(8); DCCustRec(1).BILLNAME; Tab(64); Left$(DCVRec(1).StateTag, 12)
'  Print #RptHandle, Tab(8); DCCustRec(1).ADDRESS1 '; TAB(64); "1146"
'  Print #RptHandle, Tab(8); DCCustRec(1).ADDRESS2
'  Print #RptHandle, Tab(8); RTrim$(DCCustRec(1).city); " "; RTrim$(DCCustRec(1).STATE); "  "; RTrim$(DCCustRec(1).ZIPCODE)
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, "!"
''Stand appl
'  Print #RptHandle, Tab(2); Left$(DCVRec(1).Notes, 4);
'  Print #RptHandle, Tab(13); Left$(DCVRec(1).makemodel, 11);
'  Print #RptHandle, Tab(28); Left$(DCVRec(1).Desc, 11);
'  Print #RptHandle, Tab(43); Left$(DCVRec(1).StateTag, 12);
'  Print #RptHandle, Tab(70); Using; "#######"; ControlNumber!
'
'  For LCnt = 1 To 11
'    Print #RptHandle,
'  Next LCnt
'  Print #RptHandle, Tab(8); DCCustRec(1).BILLNAME
'  Print #RptHandle, Tab(8); DCCustRec(1).ADDRESS1
'  Print #RptHandle, Tab(8); DCCustRec(1).ADDRESS2
'  Print #RptHandle, Tab(8); RTrim$(DCCustRec(1).city); " "; RTrim$(DCCustRec(1).STATE); "  "; RTrim$(DCCustRec(1).ZIPCODE)
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); DCCustRec(1).SOSEC;
'  Print #RptHandle, Tab(27); DCCustRec(1).SocSec1
'  For LCnt = 1 To 4
'    Print #RptHandle, ""
'  Next LCnt
''
''Stephen
'  ThisYear = Val(Right$(Date$, 4))
'  TYear$ = Str$(ThisYear + 1)
'  LYear$ = Str$(ThisYear)
'
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(31); "TOWN OF STEPHENS CITY"
'  Print #RptHandle, Tab(31); " Post Office Box 250"
'  Print #RptHandle, Tab(27); "STEPHENS CITY, VIRGINIA 22655"
'  Print #RptHandle, Tab(34); "(540) 869-3087"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Acct:"; DCCustIdxRec(1).IDXRECORD
'  Print #RptHandle, Tab(5); DCCustRec(1).BILLNAME
'  Print #RptHandle, Tab(5); DCCustRec(1).ADDRESS1
'  Print #RptHandle, Tab(5); DCCustRec(1).ADDRESS2
'  Print #RptHandle, Tab(5); RTrim$(DCCustRec(1).city); " "; RTrim$(DCCustRec(1).STATE); "  "; RTrim$(DCCustRec(1).ZIPCODE)
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Dear Stephens City resident:"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "The town will be selling the"; TYear$; " decal beginning January 3,"; TYear$; " and"
'  Print #RptHandle, Tab(5); "the expiration date of the"; LYear$; " decal is Febuary 15,"; TYear$; "."
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "According to our records you own the following vehicle garaged in town:"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
''  PRINT #RptHandle, TAB(2); LEFT$(DCVRec(1).Notes, 4);
'  Print #RptHandle, Tab(5); "Vehicle make & model: "; QPTrim$(DCVRec(1).makemodel)
'  Print #RptHandle, Tab(5); "      State license#: "; QPTrim$(DCVRec(1).StateTag)
'  Print #RptHandle, Tab(5); "                VIN#: "; QPTrim$(DCVRec(1).Desc)
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Please verify this information, or correct and send $20.00 for each"
'  Print #RptHandle, Tab(5); "renewal to the Town Office or stop by between 8:30 a.m. and 5:00 p.m."
'  Print #RptHandle, Tab(5); "to renew in person."
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "All personal property taxes must be paid in order to renew decals."
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Thank You,"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Joyce E. Zalar"
'  Print #RptHandle, Tab(5); "Town Clerk"
'  Print #RptHandle, FF$
'
''Stuart
'  ThisYear = (Val(Right$(Date$, 4)) - 1)
'  TYear$ = Str$(ThisYear + 1)
'  LYear$ = Str$(ThisYear)
'
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(35); "TOWN OF STUART"
'  Print #RptHandle, Tab(36); "P.O. BOX 422"
'  Print #RptHandle, Tab(34); "STUART, VA 24171"
'  Print #RptHandle, Tab(35); "(276) 694-3811"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Acct:"; DCCustIdxRec(1).IDXRECORD
'  Print #RptHandle, Tab(5); DCCustRec(1).BILLNAME
'  Print #RptHandle, Tab(5); DCCustRec(1).ADDRESS1
'  Print #RptHandle, Tab(5); DCCustRec(1).ADDRESS2
'  Print #RptHandle, Tab(5); RTrim$(DCCustRec(1).city); " "; RTrim$(DCCustRec(1).STATE); "  "; RTrim$(DCCustRec(1).ZIPCODE)
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Dear resident:"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "The town will be selling the"; TYear$; " decal beginning March 1,"; TYear$; " and"
'  Print #RptHandle, Tab(5); "the expiration date of the"; LYear$; " decal is April 30,"; TYear$; "."
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "According to our records you own the following vehicle garaged in town:"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
''  PRINT #RptHandle, TAB(2); LEFT$(DCVRec(1).Notes, 4);
'  Print #RptHandle, Tab(5); "Vehicle make & model: "; QPTrim$(DCVRec(1).makemodel)
'  Print #RptHandle, Tab(5); "      State license#: "; QPTrim$(DCVRec(1).StateTag)
'  Print #RptHandle, Tab(5); "                VIN#: "; QPTrim$(DCVRec(1).Desc)
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Please verify this information, or correct and send renewal for"
'  Print #RptHandle, Tab(5); "each to the Town Office or stop by between 9:00 t"
'  Print #RptHandle, Tab(5); "or 2:00 to 5:00 p.m to renew in person."
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "All personal property taxes must be paid in order"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Thank You,"
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle,
'  Print #RptHandle, Tab(5); "Town Clerk"
'  Print #RptHandle, FF$
'********************************************
'Tax Stuff
Public Function VAGetCustBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    VAGetCustBalance = 0
    Exit Function
  End If
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenVATaxTransFile THandle, NumOfTRecs

  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
'      If TaxTran.Amount = 0.58 Then Stop
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
'      TaxTran.BelongTo = TaxTran.BelongTo
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 5    'penalty
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10    'adjust pay down affecting credit balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 11    'adjust prepay down
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 12    'refund total prepay balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 30    'PPTRA removal transaction
        GTOwed# = OldRound(GTOwed# + TaxTran.PPTRARmvl)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    VAGetCustBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    VAGetCustBalance# = 0
  End If

  Close THandle

End Function

Public Sub OpenVATaxCustFile(TaxCustHandle As Integer, NumOfTaxCustRec As Long)
  Dim TaxCustLen As Integer
  Dim TaxCustRec As VATaxCustType
  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open "TaxCust.dat" For Random Shared As TaxCustHandle Len = TaxCustLen
  NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
End Sub
'Public Sub OpenTaxSetUpFile(TaxSetUpHandle As Integer)
'  Dim TaxSetUpLen As Integer
'  Dim TaxSetUp As TaxMasterType
'  TaxSetUpLen = Len(TaxSetUp)
'  TaxSetUpHandle = FreeFile
'  Open "TAXSETUP.DAT" For Random Shared As TaxSetUpHandle Len = TaxSetUpLen
'End Sub
Public Sub OpenVATaxTransFile(TaxTransHandle As Integer, NumOfTaxTransRecs As Long)
  Dim TaxTransLen As Integer
  Dim TaxTransRate As VATaxTransactionType
  TaxTransLen = Len(TaxTransRate)
  TaxTransHandle = FreeFile
  Open "TaxTrans.dat" For Random Shared As TaxTransHandle Len = TaxTransLen
  NumOfTaxTransRecs = LOF(TaxTransHandle) / Len(TaxTransRate)
End Sub
Public Function OldRound#(n As Double)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function
Public Function GetDefaultDCLookUP%()
  ReDim dcSetUpRec(1) As DCSetupType
  Dim RecLen As Integer
  LoadDCSetUpFile dcSetUpRec(), RecLen
  GetDefaultDCLookUP = Val(dcSetUpRec(1).DefLook)
  Erase dcSetUpRec
End Function

