Attribute VB_Name = "TCCommon"
Option Explicit
  Public ComputerName As String
  Public CurrCitiPath As String
  Public StartPath As String
  Public ScreenW As Long
  Public GCustNum As Long
  Public WhichOne As String * 1
  Public FileVers As String
  Public MyClip As String
  Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
      
      Public Const ConvErrors = "CNVRERRS.DAT"
      Public Const ConvResults = "CNVRSLTS.DAT"
      Public Const CustOptSearch = "TXCOPTSH.DAT"
      Public Const CustNameIdxFile = "TAXNMIDX.DAT"
      Public Const SrchNameIdxFile = "SRCHNMIDX.DAT"
      Public Const SocSecIdxFile = "TXSSIDX.DAT"
      Public Const TaxCustFile = "TAXCUST.DAT"
      Public Const CustPinFile = "TAXCPIN.DAT"
      Public Const TaxPropFile = "TAXPROP.DAT"
      Public Const OldTaxPropFile = "OLDTAXPROP.DAT"
      Public Const TaxPersFile = "TAXPERS.DAT"
      Public Const OldTaxPersFile = "OLDTAXPERS.DAT"
      Public Const TaxPersPINFile = "TAXPPIN.DAT"
      Public Const TaxRealPINFile = "TAXRPIN.DAT"
      Public Const TaxBillOPFile = "TAXOPBL.DAT"
      Public Const TaxTransFile = "TAXTRANS.DAT"
      Public Const PPTaxBillFile = "TAXPBILL.DAT"
      Public Const RealTaxBillFile = "TAXRBILL.DAT"
      Public Const PPTaxPreRptFile = "TXPPREBL.RPT"
      Public Const RETaxPreRptFile = "TXRPREBL.RPT"
'--------------------------------------------------
      Public Const TaxSetupName = "TAXSETUP.DAT"
      Public Const PerTaxName = "TAXPERS.DAT"
      Public Const TaxPropName = "TAXPROP.DAT"
      Public Const InternalPinFile = "TAXINPIN.DAT"
      Public Const MessageName = "TAXMESS.DAT"
      Public Const PersTaxBillFile = "TAXPBILL.DAT"
      Public Const PersTaxBillOPFile = "TAXPERSOPBL.DAT"
      Public Const RealTaxBillOPFile = "TAXREALOPBL.DAT"
      Public Const PersTempTaxBillAddOn = "TMPPERSBLADD.DAT"
      Public Const PersTaxBillInfoFile = "TAXPERSBINFO.DAT"
      Public Const TaxBillPostDateFile = "TXBLPSTDTE.DAT"
      Public Const ConversionFile = "TXCNVDAT.DAT"
      Public Const ConvSpreadFile = "TXSPREAD.DAT"
Public Sub OpenConvErrorsFile(ConvErrorsHandle As Integer, NumOfConvErrorsFiles As Long)
  Dim ConvErrorsRecLen As Integer
  Dim ConvErrorsRec As ConvErrorType
  ConvErrorsRecLen = Len(ConvErrorsRec)
  ConvErrorsHandle = FreeFile
  Open ConvErrors For Random Shared As ConvErrorsHandle Len = ConvErrorsRecLen
  NumOfConvErrorsFiles = LOF(ConvErrorsHandle) / ConvErrorsRecLen
End Sub
Public Sub OpenConvResultsFile(ConvResultsHandle As Integer, NumOfConvResultsFiles As Long)
  Dim ConvResultsRecLen As Integer
  Dim ConvResultsRec As ConvResultsType
  ConvResultsRecLen = Len(ConvResultsRec)
  ConvResultsHandle = FreeFile
  Open ConvResults For Random Shared As ConvResultsHandle Len = ConvResultsRecLen
  NumOfConvResultsFiles = LOF(ConvResultsHandle) / ConvResultsRecLen
End Sub
Public Sub OpenCustOptSearchFile(COSHandle As Integer, NumOfCOSFiles As Long)
  Dim COSRecLen As Integer
  Dim COSRec As OptCustIdxType
  COSRecLen = Len(COSRec)
  COSHandle = FreeFile
  Open CustOptSearch For Random Shared As COSHandle Len = COSRecLen
  NumOfCOSFiles = LOF(COSHandle) / COSRecLen
End Sub
Public Sub OpenSocSecIdxFile(SSHandle As Integer, NumOfSSFiles As Long)
  Dim SSRecLen As Integer
  Dim SSRec As SocSecIdxType
  SSRecLen = Len(SSRec)
  SSHandle = FreeFile
  Open SocSecIdxFile For Random Shared As SSHandle Len = SSRecLen
  NumOfSSFiles = LOF(SSHandle) / SSRecLen
End Sub
Public Sub OpenSrchNameIdxFile(SrchNameIdxHandle As Integer, NumOfNameIdxRec As Long)
  Dim SrchNameIdxLen As Integer
  Dim SrchNameIdxRec As SrchNameIdxType
  SrchNameIdxLen = Len(SrchNameIdxRec)
  SrchNameIdxHandle = FreeFile
  Open SrchNameIdxFile For Random Shared As SrchNameIdxHandle Len = SrchNameIdxLen
  NumOfNameIdxRec = LOF(SrchNameIdxHandle) / Len(SrchNameIdxRec)
End Sub
Public Sub OpenNameIdxFile(NameIdxHandle As Integer, NumOfNameIdxRec As Long)
  Dim NameIdxLen As Integer
  Dim NameIdxRec As CustNameIdxType
  NameIdxLen = Len(NameIdxRec)
  NameIdxHandle = FreeFile
  Open CustNameIdxFile For Random Shared As NameIdxHandle Len = NameIdxLen
  NumOfNameIdxRec = LOF(NameIdxHandle) / Len(NameIdxRec)
End Sub
Public Sub OpenOldPersPropFile(PersPropHandle As Integer, NumOfPersProp As Long)
  Dim PersPropLen As Integer
  Dim PersPropRec As PersonalRecType
  PersPropLen = Len(PersPropRec)
  PersPropHandle = FreeFile
  Open OldTaxPersFile For Random Shared As PersPropHandle Len = PersPropLen
  NumOfPersProp = LOF(PersPropHandle) / Len(PersPropRec)
End Sub
Public Sub OpenOldRealPropFile(RealPropHandle As Integer, NumOfRealProp As Long)
  Dim RealPropLen As Integer
  Dim RealPropRec As PropertyRecType
  RealPropLen = Len(RealPropRec)
  RealPropHandle = FreeFile
  Open OldTaxPropFile For Random Shared As RealPropHandle Len = RealPropLen
  NumOfRealProp = LOF(RealPropHandle) / Len(RealPropRec)
End Sub
Public Sub OpenConvSpreadFile(SpreadHandle As Integer, NumOfSpreadFiles As Integer)
  Dim SpreadLen As Integer
  Dim SpreadRec As ConvSpreadsheet
  SpreadLen = Len(SpreadRec)
  SpreadHandle = FreeFile
  Open ConvSpreadFile For Random Shared As SpreadHandle Len = SpreadLen
  NumOfSpreadFiles = LOF(SpreadHandle) / SpreadLen
End Sub
Public Sub OpenTempConvFile(TCHandle As Integer, NumOfTCFiles As Long)
  Dim TCLen As Integer
  Dim TCRec As TempConversionData
  TCLen = Len(TCRec)
  TCHandle = FreeFile
  Open ConversionFile For Random Shared As TCHandle Len = TCLen
  NumOfTCFiles = LOF(TCHandle) / TCLen
End Sub
      
Public Sub OpenPersPinFile(PersPinHandle As Integer, NumOfPersPins As Long)
  Dim PersPinLen As Integer
  Dim PersPinRec As PINSearchType
  PersPinLen = Len(PersPinRec)
  PersPinHandle = FreeFile
  Open TaxPersPINFile For Random Shared As PersPinHandle Len = PersPinLen
  NumOfPersPins = LOF(PersPinHandle) / Len(PersPinRec)
End Sub
Public Sub OpenRealPinFile(RealPinHandle As Integer, NumOfRealPins As Long)
  Dim RealPinLen As Integer
  Dim RealPinRec As PINSearchType
  RealPinLen = Len(RealPinRec)
  RealPinHandle = FreeFile
  Open TaxRealPINFile For Random Shared As RealPinHandle Len = RealPinLen
  NumOfRealPins = LOF(RealPinHandle) / Len(RealPinRec)
End Sub
Public Sub OpenRealPropFile(RealPropHandle As Integer, NumOfRealProp As Long)
  Dim RealPropLen As Integer
  Dim RealPropRec As PropertyRecType
  RealPropLen = Len(RealPropRec)
  RealPropHandle = FreeFile
  Open TaxPropFile For Random Shared As RealPropHandle Len = RealPropLen
  NumOfRealProp = LOF(RealPropHandle) / Len(RealPropRec)
End Sub
Public Sub OpenIntPinFile(IntPinHandle As Integer, NumOfIntPins As Long)
  Dim IntPinLen As Integer
  Dim IntPinRec As InternalPinType
  IntPinLen = Len(IntPinRec)
  IntPinHandle = FreeFile
  Open InternalPinFile For Random Shared As IntPinHandle Len = IntPinLen
  NumOfIntPins = LOF(IntPinHandle) / Len(IntPinRec)
End Sub
     
Public Sub OpenCustPinFile(CustPinHandle As Integer, NumOfCustPins As Long)
  Dim CustPinLen As Integer
  Dim CustPinRec As PINRecType
  CustPinLen = Len(CustPinRec)
  CustPinHandle = FreeFile
  Open CustPinFile For Random Shared As CustPinHandle Len = CustPinLen
  NumOfCustPins = LOF(CustPinHandle) / Len(CustPinRec)
End Sub
Public Sub OpenPersPropFile(PersPropHandle As Integer, NumOfPersProp As Long)
  Dim PersPropLen As Integer
  Dim PersPropRec As PersonalRecType
  PersPropLen = Len(PersPropRec)
  PersPropHandle = FreeFile
  Open TaxPersFile For Random Shared As PersPropHandle Len = PersPropLen
  NumOfPersProp = LOF(PersPropHandle) / Len(PersPropRec)
End Sub
Public Sub OpenTaxPropFile(TaxPropHandle As Integer, NumOfTaxProps As Long)
  Dim TaxPropLen As Integer
  Dim TaxPropRec As PropertyRecType
  TaxPropLen = Len(TaxPropRec)
  TaxPropHandle = FreeFile
  Open TaxPropName For Random Shared As TaxPropHandle Len = Len(TaxPropRec)
  NumOfTaxProps = LOF(TaxPropHandle) / Len(TaxPropRec)
End Sub
      
Public Sub OpenTaxPersFile(PersTaxHandle As Integer, NumOfPersRecs As Long)
  Dim PersTaxLen As Integer
  Dim PersTaxRec As PersonalRecType
  PersTaxLen = Len(PersTaxRec)
  PersTaxHandle = FreeFile
  Open PerTaxName For Random Shared As PersTaxHandle Len = PersTaxLen
  NumOfPersRecs = LOF(PersTaxHandle) / Len(PersTaxRec)
End Sub
      
Public Sub OpenTaxCustFile(TaxCustHandle As Integer, NumOfTaxCustRec As Long)
  Dim TaxCustLen As Integer
  Dim TaxCustRec As TaxCustType
  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open TaxCustFile For Random Shared As TaxCustHandle Len = TaxCustLen
  NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
End Sub

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

Public Function Exist(FileName$) As Boolean
  Dim FileHandle As Integer
  Dim TempSize As Long
  On Local Error Resume Next
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  TempSize = LOF(FileHandle)
  Close FileHandle
  If TempSize <= 0 Then
    Kill FileName$
    Exist = False
  Else
    Exist = True
  End If

End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
  
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function

Public Function Date2Num%(TheDate$)
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function
Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

Public Function OldRound#(n As Double)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function

Public Sub KillFile(FileName As String)
  On Local Error Resume Next
  If Exist(FileName$) Then
    Kill FileName$
  End If
End Sub

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmTCViewPrint.ReportName = ReportFile$
   frmTCViewPrint.Caption = Title
   frmTCViewPrint.PgNum = PgNum
   frmTCViewPrint.cmdAlignment.Visible = False
   If ForceSBar Then
     frmTCViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmTCViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   frmTCViewPrint.Show 1
   Unload frmTCLoadingRpt
End Sub

Public Static Function Using$(ByVal fmt As String, ByVal Number As Double)
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
  FmtNumber = Space$(FmtLen)
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
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
  RSet FmtNumber = TempNumber
  Using = FmtNumber
  
End Function

Public Function ReplaceString$(Text As String, ChangeThis As String, ToThis As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim NewText As String
  Dim thischar$
  Dim CTChar$
  Dim TTChar$
  Dim CTLen As Integer
  Dim TTLen As Integer
  Dim BigLen As Integer
  'this function takes the incoming text and rebuilds it one
  'letter at a time until it encounters the text to change
  'at which time it replaces the text to change with the
  'new text
  StrLen = Len(Text)
  CTLen = Len(ChangeThis$)
  TTLen = Len(ToThis$)
  If CTLen > TTLen Then
    BigLen = CTLen
  ElseIf TTLen > CTLen Then
    BigLen = TTLen
  Else
    BigLen = CTLen
  End If
  
  For cnt = 1 To StrLen 'set up loop to iterate thru entire text
    thischar = Mid$(Text, cnt, 1) 'step thru text a letter at a time
    CTChar = Mid$(Text, cnt, CTLen) 'starting with the current letter
    'read ahead the length of the text "change this"
    If CTChar = ChangeThis Then 'if we find the "change this" in the
    'text
      NewText = NewText + ToThis 'assign the length of CTChar to "ToThis"
      'inside the rebuilt new text
      cnt = cnt + BigLen - 1 'advance count to compensate for the addition of
      'CTChar
    Else
      NewText = NewText + thischar 'build new text one letter at a time
    End If
  Next
  ReplaceString$ = Trim$(NewText) 'rim out the new text
  Text = ReplaceString$ 'old text is now new text
End Function

Public Sub InsertSSNDashes(ByRef SSN As String)
  Dim ThisLen As Integer
  Dim x As Integer
  Dim NewSSN As String
  
  If InStr(1, SSN, "-") = 4 And InStr(1, SSN, "-") = 7 Then
    Exit Sub
  End If
  ThisLen = Len(SSN)
  ReDim thischar(1 To 9) As String
  For x = 1 To 9
    thischar(x) = Mid(SSN, x, 1)
    If Not IsNumeric(thischar(x)) Or thischar(x) = "" Then
      thischar(x) = " "
    End If
  Next x
  For x = 1 To 9
    NewSSN = NewSSN + thischar(x)
    If x = 3 Or x = 5 Then
      NewSSN = NewSSN + "-"
    End If
  Next x
  
  SSN = NewSSN
  
End Sub

Public Sub MakeRealPINFile()
  Dim RealPINS As PINSearchType
  Dim RPHandle As Integer
  Dim NumOfRealPins As Long
  Dim RealRec As PropertyRecType
  Dim PHandle As Integer
  Dim NumOfPropRecs As Long
  Dim cnt&
  
  KillFile TaxRealPINFile
  
  OpenRealPropFile PHandle, NumOfPropRecs
  
  OpenRealPinFile RPHandle, NumOfRealPins
  
  For cnt& = 1 To NumOfPropRecs&
    Get PHandle, cnt&, RealRec
    RealPINS.PIN = RealRec.RealPin
    RealPINS.Cust = cnt&
    Put RPHandle, cnt&, RealPINS
  Next
  
  Close

End Sub

Public Sub MakePersPINFile()
  Dim PersPINS As PINSearchType
  Dim PPHandle As Integer
  Dim NumOfPersPins As Long
  Dim PersRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPropRecs As Long
  Dim cnt&
  
  KillFile TaxPersPINFile
  
  OpenPersPropFile PPHandle, NumOfPropRecs
  
  OpenPersPinFile PRHandle, NumOfPersPins
  
  For cnt& = 1 To NumOfPropRecs&
    Get PPHandle, cnt&, PersRec
    PersPINS.PIN = PersRec.PropPin
    PersPINS.Cust = cnt&
    Put PRHandle, cnt&, PersPINS
  Next
  
  Close

End Sub

Public Sub TCMsg(Top As Integer, Message As String)
  frmTCMsg.Label1.Caption = Message
  frmTCMsg.Label1.Top = Top
  frmTCMsg.Show vbModal
End Sub
Public Sub Savemsg(Top As Integer, Message As String)
  frmTCSave.Label1.Caption = Message
  frmTCSave.Label1.Top = Top
  frmTCSave.Show vbModal
End Sub

Public Function TCMsgWOpts(Top As Integer, Message As String, CmdF10 As String, CmdESC As String) As String
  frmTCMsgWOpts.Label1.Caption = Message
  frmTCMsgWOpts.Label1.Top = Top
  frmTCMsgWOpts.cmdCont.Text = CmdF10
  frmTCMsgWOpts.cmdExit.Text = CmdESC
  frmTCMsgWOpts.Show vbModal
  TCMsgWOpts = frmTCMsgWOpts.fptxtChoice.Text
End Function

Public Function GetPhoneNum(PhoneNum$) As String
  Dim ThisPhone$
  Dim NewPhone$
  Dim ThisLen As Integer
  Dim x As Integer
  
  ThisPhone$ = ReplaceString(PhoneNum$, "-", "")
  ThisPhone$ = ReplaceString(ThisPhone$, "(", "")
  ThisPhone$ = ReplaceString(ThisPhone$, ")", "")
  ThisPhone$ = ReplaceString(ThisPhone$, " ", "")
  
  NewPhone = ""
  ThisLen = Len(ThisPhone)
  If ThisLen = 10 Then
    For x = 1 To 12
      If x = 4 Or x = 8 Then
        NewPhone = NewPhone + "-"
      ElseIf x < 4 Then
        NewPhone = NewPhone + Mid(ThisPhone, x, 1)
      ElseIf x < 8 And x > 4 Then
        NewPhone = NewPhone + Mid(ThisPhone, x - 1, 1)
      ElseIf x > 8 Then
        NewPhone = NewPhone + Mid(ThisPhone, x - 2, 1)
      End If
    Next x
  ElseIf ThisLen = 7 Then
    For x = 1 To 12
      If x <= 3 Then
        NewPhone = NewPhone + "0"
      ElseIf x = 4 Or x = 8 Then
        NewPhone = NewPhone + "-"
      ElseIf x <= 7 Then
        NewPhone = NewPhone + Mid(ThisPhone, x - 4, 1)
      Else
        NewPhone = NewPhone + Mid(ThisPhone, x - 5, 1)
      End If
    Next x
  End If
    
  GetPhoneNum = NewPhone

End Function

Public Function InsertZipDash(Zip$) As String
  Dim ZipLen As Integer
  Dim Thisch$
  Dim x As Integer
  Dim ThisZip$
  
  If Mid(Zip$, 6, 1) = "-" Then
    InsertZipDash = Zip$
    Exit Function
  End If
  
  ZipLen = Len(QPTrim$(Zip$))
  If ZipLen <= 5 Then
    InsertZipDash = Zip$
    Exit Function
  End If
  
  For x = 1 To ZipLen
    If x = 6 Then
      Thisch = "-" + Mid(Zip, x, 1)
    Else
      Thisch = Mid(Zip, x, 1)
    End If
    If x <> 6 Then
      If Not IsNumeric(Thisch) Then
        InsertZipDash = Zip$
        Exit Function
      End If
    Else
      If Not IsNumeric(Mid(Thisch, 2, 1)) Then
        InsertZipDash = Zip$
        Exit Function
      End If
    End If
    ThisZip = ThisZip + Thisch
  Next x
  InsertZipDash = ThisZip
End Function

Public Sub CreateCustNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigName$
  Dim ThisName$
  Dim Thisx As Long
  Dim SmallName$
  Dim TempName As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As CustNameIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).CustName = QPTrim$(CustRec.CustName)
    ThisName = QPTrim$(CustRec.CustName)
    If ThisName > BigName Then
      BigName = ThisName
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigName = BigName + "A"
  SmallName = BigName
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisName = TempCustIdx(x).CustName
      If ThisName < SmallName Then
        SmallName = ThisName
        Thisx = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(Thisx)
    TempCustIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do ' NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallName = BigName
  Loop
  
  KillFile "TAXNMIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateCustNameIdx", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Sub CreateSrchNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigName$
  Dim ThisName$
  Dim Thisx As Long
  Dim SmallName$
  Dim TempName As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As SrchNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As SrchNameIdxType
  Dim ThisCnt As Integer
  Dim NumOfIdxRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.SName) <> "" Then
      Exit For
    End If
  Next x
  
  If x > NumOfCustRecs Then
    KillFile "SRCHNMIDX.DAT"
    Close CustHandle
    Exit Sub
  End If
  
  ReDim TempCustIdx(1 To NumOfCustRecs) As SrchNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).SearchName = QPTrim$(CustRec.SName)
    ThisName = QPTrim$(CustRec.SName)
    If ThisName > BigName Then
      BigName = ThisName
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigName = BigName + "A"
  SmallName = BigName
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisName = TempCustIdx(x).SearchName
      If ThisName < SmallName Then
        SmallName = ThisName
        Thisx = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(Thisx)
    TempCustIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do ' NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallName = BigName
  Loop
  
  KillFile "SRCHNMIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenSrchNameIdxFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateSrchNameIdx", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub

Public Sub CreateSSIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim Thisx As Long
  Dim SmallNum As Double
  Dim TempName As Long
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim SSIdx As SocSecIdxType
  Dim SSIdxHandle As Integer
  Dim SSIdxRecLen As Long
  Dim NumOfSSIdxRecs As Long
  Dim RecNum As Long
  Dim HoldThis As SocSecIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  Dim SSN$
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  
  ReDim TempSSIdx(1 To NumOfCustRecs) As SocSecIdxType
  
  BigNum = 0
  DoEvents
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    If QPTrim$(CustRec.CSSN) = "" Then CustRec.CSSN = "111111111"
    SSN = ReplaceString(CustRec.CSSN, "-", "")
    SSN = ReplaceString(SSN, " ", "")
    SSN = QPTrim(SSN)
    If SSN = "" Then GoTo BadNum
    If Not IsNumeric(SSN) Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempSSIdx(ThisCnt).CustRec = x
    TempSSIdx(ThisCnt).SSNum = CDbl(SSN)
    ThisNum = CDbl(SSN)
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close CustHandle
  
  DoEvents
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisNum = TempSSIdx(x).SSNum
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        Thisx = x
      End If
    Next x
    HoldThis = TempSSIdx(Nextx)
    TempSSIdx(Nextx) = TempSSIdx(Thisx)
    TempSSIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  KillFile "TXSSIDX.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenSocSecIdxFile SSIdxHandle, NumOfSSIdxRecs
  For x = 1 To ThisCnt
    SSIdx = TempSSIdx(x)
    Put SSIdxHandle, x, SSIdx
  Next x
  
  Close SSIdxHandle
  
  DoEvents
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateSSIdx", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    End
  
End Sub

'Public Sub CreateOptCustIdx()
'  Dim CHandle As Integer
'  Dim TotalAccts As Long
'  Dim x As Long
'  Dim n As Long
'  Dim Nextx As Long
'  Dim y As Long, cnt As Long
'  Dim ThisText$, CustRecNo As Long
'  Dim CustCnt As Long
'  Dim BigDesc$
'  Dim ThisDesc$
'  Dim Thisx As Long
'  Dim SmallDesc$
'  Dim CustRec As TaxCustType
'  Dim CustHandle As Integer
'  Dim NumOfCustRecs As Long
'  Dim CustIdx As OptCustIdxType
'  Dim CustIdxHandle As Integer
'  Dim CustIdxRecLen As Long
'  Dim RecNum As Long
'  Dim HoldThis As OptCustIdxType
'  Dim ThisCnt As Long
'  Dim NumOfIdxRecs As Long
'  Dim First As Integer
'  Dim Second As Integer
'  Dim Third As Integer
'  Dim Fourth As Integer
'  Dim Fifth As Integer
'  Dim Sixth As Integer
'  Dim Seventh As Integer
'  Dim Eighth As Integer
'  Dim Ninth As Integer
'  Dim First1 As Integer
'  Dim Second1 As Integer
'  Dim Third1 As Integer
'  Dim Fourth1 As Integer
'  Dim Fifth1 As Integer
'  Dim Sixth1 As Integer
'  Dim Seventh1 As Integer
'  Dim Eighth1 As Integer
'  Dim Ninth1 As Integer
'
'  On Error GoTo ERRORSTUFF
'
'  OpenTaxCustFile CustHandle, NumOfCustRecs
'
'  For x = 1 To NumOfCustRecs
'    Get CustHandle, x, CustRec
'    If QPTrim$(CustRec.OptSrchDesc) <> "" Then
'      Exit For
'    End If
'  Next x
'
'  If x > NumOfCustRecs Then
'    KillFile "TXCOPTSH.DAT"
'    Close CustHandle
'    Exit Sub
'  End If
'
'  ReDim TempCustIdx(1 To NumOfCustRecs) As OptCustIdxType
'
'  BigDesc = "A"
'  For x = 1 To NumOfCustRecs
'    Get CustHandle, x, CustRec
'    If CustRec.Deleted <> 0 Then GoTo BadNum
'    ThisCnt = ThisCnt + 1
'    TempCustIdx(ThisCnt).CustRec = x
'    TempCustIdx(ThisCnt).OptDesc = QPTrim$(CustRec.OptSrchDesc)
'    TempCustIdx(ThisCnt).CustPin = CustRec.PIN
'    ThisDesc = QPTrim$(CustRec.OptSrchDesc)
'    If ThisDesc > BigDesc Then
'      BigDesc = ThisDesc
'    End If
'BadNum:
'  Next x
'  Close CustHandle
'
'  BigDesc = BigDesc + "A"
'  SmallDesc = BigDesc
'  Nextx = 1
'
'  Do
'    For x = Nextx To ThisCnt
'      ThisDesc = QPTrim$(TempCustIdx(x).OptDesc)
'      If ThisDesc <= SmallDesc Then
'        SmallDesc = ThisDesc
'        Thisx = x
'      End If
'    Next x
'    HoldThis = TempCustIdx(Nextx)
'    TempCustIdx(Nextx) = TempCustIdx(Thisx)
'    TempCustIdx(Thisx) = HoldThis
'    If Nextx = ThisCnt Then Exit Do
'    Nextx = Nextx + 1
'    SmallDesc = BigDesc
'  Loop
'
'  KillFile "TXCOPTSH.DAT"
'  'must kill the old file because if a customer is deleted
'  'it still remains as a record...not deleting causes multiple
'  'repeats of the last customer depending on how many customers
'  'have been deleted
'
'  OpenCustOptSearchFile CustIdxHandle, NumOfIdxRecs
'  For x = 1 To ThisCnt
'    CustIdx = TempCustIdx(x)
'    Put CustIdxHandle, x, CustIdx
'  Next x
'
'  Close CustIdxHandle
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateOptCustIdx", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'    End
'
'
'End Sub
'
Public Sub CreateOptCustIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Long
  Dim x As Long
  Dim n As Long
  Dim Nextx As Long
  Dim y As Long, cnt As Long
  Dim ThisText$, CustRecNo As Long
  Dim CustCnt As Long
  Dim BigDesc$
  Dim ThisDesc$
  Dim Thisx As Long
  Dim SmallDesc$
  Dim CustRec As TaxCustType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Long
  Dim CustIdx As OptCustIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Long
  Dim RecNum As Long
  Dim HoldThis As OptCustIdxType
  Dim ThisCnt As Long
  Dim NumOfIdxRecs As Long
  Dim EmptyCnt As Long
  Dim NotEmptyCnt As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxCustFile CustHandle, NumOfCustRecs
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.OptSrchDesc) <> "" Then
      Exit For
    End If
  Next x
  
  If x > NumOfCustRecs Then
    KillFile "TXCOPTSH.DAT"
    Close CustHandle
    Exit Sub
  End If
      
  ReDim TempCustIdx(1 To NumOfCustRecs) As OptCustIdxType
  ReDim TempNotEmptyIdx(1 To 1) As OptCustIdxType
  ReDim TempEmptyIdx(1 To 1) As OptCustIdxType
  BigDesc = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).OptDesc = QPTrim$(CustRec.OptSrchDesc)
    TempCustIdx(ThisCnt).CustPin = CustRec.PIN
    ThisDesc = QPTrim$(CustRec.OptSrchDesc)
    If ThisDesc > BigDesc Then
      BigDesc = ThisDesc
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigDesc = BigDesc + "A"
  SmallDesc = BigDesc
  
  Nextx = 1
  EmptyCnt = 0
  
  Do While Nextx <= ThisCnt
    ThisDesc = QPTrim$(TempCustIdx(Nextx).OptDesc)
    If ThisDesc = "" Then
      EmptyCnt = EmptyCnt + 1
      ReDim Preserve TempEmptyIdx(1 To EmptyCnt) As OptCustIdxType
      TempEmptyIdx(EmptyCnt) = TempCustIdx(Nextx)
    Else
      NotEmptyCnt = NotEmptyCnt + 1
      ReDim Preserve TempNotEmptyIdx(1 To NotEmptyCnt) As OptCustIdxType
      TempNotEmptyIdx(NotEmptyCnt) = TempCustIdx(Nextx)
    End If
    Nextx = Nextx + 1
  Loop
  Nextx = 1
  
  Do
    For x = Nextx To NotEmptyCnt
      ThisDesc = QPTrim$(TempNotEmptyIdx(x).OptDesc)
      If ThisDesc <= SmallDesc Then
        SmallDesc = ThisDesc
        Thisx = x
      End If
    Next x
EmptyStr: 'added 3/27/06
    HoldThis = TempNotEmptyIdx(Nextx)
    TempNotEmptyIdx(Nextx) = TempNotEmptyIdx(Thisx)
    TempNotEmptyIdx(Thisx) = HoldThis
    If Nextx = NotEmptyCnt Then Exit Do
    Nextx = Nextx + 1
    SmallDesc = BigDesc
  Loop
  
  KillFile "TXCOPTSH.DAT"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  
  OpenCustOptSearchFile CustIdxHandle, NumOfIdxRecs
  For x = 1 To NotEmptyCnt 'ThisCnt
    CustIdx = TempNotEmptyIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  
  For x = NotEmptyCnt + 1 To EmptyCnt 'ThisCnt
    CustIdx = TempEmptyIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  
  Close CustIdxHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "TaxCommon", "CreateOptCustIdx", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    End
  
  
End Sub


