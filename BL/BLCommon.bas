Attribute VB_Name = "BLCommon"
Option Explicit
  Public ScreenW As Long
  Public coladj As Double
  Public doAlign As Boolean
  Public alnRpt$
  Public BadMaskFlag As Boolean
  Public NumOfAligns As Integer
  Public ComputerName As String
  Public CurrCitiPath As String
  Public StartPath As String
  Public RptOpt As Integer 'used to determine the type of reports; graphic or text
  Public ToPrint1(1 To 10) As Integer
  Public ToPrint2(1 To 10) As Integer
  Public GCatNum As Integer
  Public GCustNum As Integer
  Public GPayNum As Integer
  Public ItemChangeFlag As Boolean
  Public OPERNUM As Integer
  Public EditFlag As Boolean 'used in entering/editing payment transactions
  Public FromCustEdit As Boolean
  Public Twiddle As String
  Public RecpDef As Integer
  Public RecpPort As String
  Public OmitList() As Long
  Public InPayOmit() As Long
  Public OmitCnt As Long
  Public PayOmitCnt As Long
  Public DidPrint As Integer
  Public PrintSign As Boolean 'used in printing the laser license
  Public PayDate As String
  Public ThisCustXNum As Integer 'used when opening frmBLTransHistJr modally
  Public FromBL As Boolean
  Public CntrlDef As Integer
  
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
       Public Const BLCatCodeName = "ARCODE.DAT"
       Public Const BLCustFileName = "ARCUST.DAT"
       Public Const BLTransFileName = "ARTRANS.DAT"
       Public Const JGLAcctIdxFile = "GLACCT.IDX"
       Public Const AcctFileName = "GLACCT.DAT"
       Public Const CatCodeIdxName = "arcatcodeidx.dat"
       Public Const CustNameIdx = "arcustnameidx.dat"
       Public Const LicNumIdx = "arlicnumidx.dat"
       Public Const CustNumIdx = "arcustnumidx.dat"
       Public Const CustSearchNameIdx = "arsrhidx.dat"
       Public Const BLTransTempPost = "artmppst.dat"
       Public Const BLOperRecName = "CMOPER.DAT"
       Public Const BLPayFileName = "AREDPY"
       Public Const BLTownSetUpName = "artownsu.dat"
       Public Const BLTempCustRecName = "artmpcus.dat"
       Public Const BLTempPrintLicName = "artmplic.dat"
       Public Const BLTempPenaltyCharges = "artmppen.dat"
       Public Const BLLaserLetterName1 = "arlaser1.dat"
       Public Const BLLaserLetterName2 = "arlaser2.dat"
       Public Const BLLaserLetterName3 = "arlaser3.dat"
       Public Const BLLaserLetterName4 = "arlaser4.dat"
       Public Const BLLaserLetterName5 = "arlaser5.dat"
Public Sub OpenLaserFile5(LHandle As Integer)
  Dim LRec As LaserLetterType5
  Dim LRecLen As Integer
  LRecLen = Len(LRec)
  LHandle = FreeFile
  Open BLLaserLetterName5 For Random Shared As LHandle Len = LRecLen
End Sub
Public Sub OpenLaserFile1(LHandle As Integer)
  Dim LRec As LaserLetterType1
  Dim LRecLen As Integer
  LRecLen = Len(LRec)
  LHandle = FreeFile
  Open BLLaserLetterName1 For Random Shared As LHandle Len = LRecLen
End Sub
Public Sub OpenLaserFile2(LHandle As Integer)
  Dim LRec As LaserLetterType2
  Dim LRecLen As Integer
  LRecLen = Len(LRec)
  LHandle = FreeFile
  Open BLLaserLetterName2 For Random Shared As LHandle Len = LRecLen
End Sub
Public Sub OpenLaserFile3(LHandle As Integer)
  Dim LRec As LaserLetterType3
  Dim LRecLen As Integer
  LRecLen = Len(LRec)
  LHandle = FreeFile
  Open BLLaserLetterName3 For Random Shared As LHandle Len = LRecLen
End Sub
Public Sub OpenLaserFile4(LHandle As Integer)
  Dim LRec As LaserLetterType4
  Dim LRecLen As Integer
  LRecLen = Len(LRec)
  LHandle = FreeFile
  Open BLLaserLetterName4 For Random Shared As LHandle Len = LRecLen
End Sub

Public Sub OpenGLIdxFile(GLHandle As Integer)
  Dim GLRec As JGLAcctIdxType
  Dim GLRecLen As Integer
  GLRecLen = Len(GLRec)
  GLHandle = FreeFile
  Open JGLAcctIdxFile For Random Shared As GLHandle Len = GLRecLen
End Sub
Public Sub OpenGLAcctFile(GLHandle As Integer)
  Dim GLRec As GLAcctRecType
  Dim GLRecLen As Integer
  GLRecLen = Len(GLRec)
  GLHandle = FreeFile
  Open AcctFileName For Random Shared As GLHandle Len = GLRecLen
End Sub
Public Sub OpenTempLicPrint(TempLicPrintHandle As Integer)
  Dim TempLicPrintRec As TempLicPrintType
  Dim TempLicPrintLen As Integer
  TempLicPrintLen = Len(TempLicPrintRec)
  TempLicPrintHandle = FreeFile
  Open BLTempPrintLicName For Random Shared As TempLicPrintHandle Len = TempLicPrintLen
End Sub
'Public Sub OpenTempCharge(TempChargeHandle As Integer)
'  Dim TempChargeRec As TempChargesType
'  Dim TempChargeLen As Integer
'  TempChargeLen = Len(TempChargeRec)
'  TempChargeHandle = FreeFile
'  Open BLTempChargesName For Random Shared As TempChargeHandle Len = TempChargeLen
'End Sub
Public Sub OpenTempCustRec(TempCustHandle As Integer)
  Dim TempCustRec As TempCustRecType
  Dim TempCustLen As Integer
  TempCustLen = Len(TempCustRec)
  TempCustHandle = FreeFile
  Open BLTempCustRecName For Random Shared As TempCustHandle Len = TempCustLen
End Sub

Public Sub OpenTownFile(TownRecHandle As Integer)
  Dim TownRec As TownSetUpType
  Dim TownRecLen As Integer
  TownRecLen = Len(TownRec)
  TownRecHandle = FreeFile
  Open BLTownSetUpName For Random Shared As TownRecHandle Len = TownRecLen
End Sub
       
Public Sub OpenPayFile(PayHandle As Integer, Oper As Integer)
  Dim PayRec As AREditPaymentRecType
  Dim PayRecLen As Integer
  Dim Operator$
  
  Operator$ = Str(Oper)
  PayRecLen = Len(PayRec)
  PayHandle = FreeFile
  Open BLPayFileName + Operator$ + ".DAT" For Random Shared As PayHandle Len = PayRecLen
End Sub
Public Sub OpenOperRecFile(OperRecHandle As Integer)
  Dim OperRec As CMOperRecType
  Dim OperRecLen As Integer
  OperRecLen = Len(OperRec)
  OperRecHandle = FreeFile
  Open BLOperRecName For Random Shared As OperRecHandle Len = OperRecLen
End Sub

Public Sub OpenTempPostFile(TempPostHandle As Integer)
  Dim TempPostRec As TempTransPostType
  Dim TempPostLen As Integer
  TempPostLen = Len(TempPostRec)
  TempPostHandle = FreeFile
  Open BLTransTempPost For Random Shared As TempPostHandle Len = TempPostLen
End Sub

Public Sub OpenPenTransFile(PenTransHandle As Integer)
  Dim PenTransRec As TempPenaltyCharges
  Dim PenTransLen As Integer
  PenTransLen = Len(PenTransRec)
  PenTransHandle = FreeFile
  Open BLTempPenaltyCharges For Random Shared As PenTransHandle Len = PenTransLen
End Sub

Public Sub OpenSrchNameIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustSearchNameIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustSearchNameIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub

Public Sub OpenCustNumIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustNumIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustNumIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub
Public Sub OpenLicNumIdxFile(LicIdxHandle As Integer)
  Dim LicIdx As CustLicNumIdxType
  Dim LicIdxLen As Integer
  LicIdxLen = Len(LicIdx)
  LicIdxHandle = FreeFile
  Open LicNumIdx For Random Shared As LicIdxHandle Len = LicIdxLen
End Sub
       
Public Sub OpenTransFile(TransHandle As Integer)
  Dim TransRec As ARTransRecType
  Dim TransLen As Integer
  TransLen = Len(TransRec)
  TransHandle = FreeFile
  Open BLTransFileName For Random Shared As TransHandle Len = TransLen
End Sub

Public Sub OpenCustNameIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustNameIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustNameIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub
Public Sub OpenCustFile(CustHandle As Integer)
  Dim CustRec As ARCustRecType
  Dim CustLen As Integer
  CustLen = Len(CustRec)
  CustHandle = FreeFile
  Open BLCustFileName For Random Shared As CustHandle Len = CustLen
End Sub

Public Sub OpenCatCodeFile(CatCodeHandle As Integer)
  Dim CatCodeRec As ARNewCatCodeRecType
  Dim CatCodeLen As Integer
  CatCodeLen = Len(CatCodeRec)
  CatCodeHandle = FreeFile
  Open BLCatCodeName For Random Shared As CatCodeHandle Len = CatCodeLen
End Sub
Public Sub OpenCatCodeIdxFile(CatCodeIdxHandle As Integer)
  Dim CatCodeIdx As CatCodeIdxType
  Dim CatCodeIdxLen As Integer
  
  CatCodeIdxLen = Len(CatCodeIdx)
  CatCodeIdxHandle = FreeFile
  Open CatCodeIdxName For Random Shared As CatCodeIdxHandle Len = CatCodeIdxLen
End Sub
'Public Sub OpenSetupFile(SetUpFileNum As Integer)
'  Dim GLRec As GLSetupRecType
'  Dim GLSetupRecLen As Integer
'
'  GLSetupRecLen = Len(GLRec)
'  SetUpFileNum = FreeFile
'  Open GLSetupName For Random As SetUpFileNum Len = GLSetupRecLen
'End Sub
Public Function PromptSaveChanges(frm As Form) As SaveChangeOptions1
  frmBLWarnChangeMade.Show vbModal, frm
  PromptSaveChanges = frmBLWarnChangeMade.Selection
'  Unload frmBLWarnChangeMade
End Function
       
Public Sub Terminate()
   Dim UBFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
   Close
   
   If Exist("custlistopen.dat") Then KillFile "custlistopen.dat"
   If Exist("catlistopen.dat") Then KillFile "catlistopen.dat"
   If Exist("categoryedit.dat") Then KillFile "categoryedit.dat"
   If Exist("customeredit.dat") Then KillFile "customeredit.dat"
   If Exist("adjustbalance.dat") Then KillFile "adjustbalance.dat"
   If Exist("custbalList.dat") Then KillFile "custbalList.dat"
   If Exist("custlistRpt.dat") Then KillFile "custlistRpt.dat"
   If Exist("custlicList.dat") Then KillFile "custlicList.dat"
   If Exist("custXlicList.dat") Then KillFile "custXlicList.dat"
   If Exist("custappList.dat") Then KillFile "custappList.dat"
   If Exist("custquickList.dat") Then KillFile "custquickList.dat"
   If Exist("custappsRenews.dat") Then KillFile "custappsRenews.dat"
   If Exist("custappIssue.dat") Then KillFile "custappIssue.dat"
   If Exist("transentry.dat") Then KillFile "transentry.dat"
   If Exist("pencalc.dat") Then KillFile "pencalc.dat"
   If Exist("townsetup.dat") Then KillFile "townsetup.dat"
   If Exist("issueappslics.dat") Then KillFile "issueappslics.dat"
   If Exist("advanceltrprint.dat") Then KillFile "advanceltrprint.dat"
   If Exist("pencalcscr.dat") Then KillFile "pencalcscr.dat"
   If Exist("dlnqnotice.dat") Then KillFile "dlnqnotice.dat"
   If Exist("dlnqmllbls.dat") Then KillFile "dlnqmllbls.dat"
   If Exist("setstatus.dat") Then KillFile "setstatus.dat"
   If Exist("mllbls.dat") Then KillFile "mllbls.dat"
   If Exist("changeaccmeth.dat") Then KillFile "changeaccmeth.dat"
   If Exist("XlistInactiveY.dat") Then KillFile "XlistInactiveY.dat"
   If Exist("inoutrpt.dat") Then KillFile "inoutrpt.dat"
   If Exist("custinfomodal.dat") Then KillFile "custinfomodal.dat"
   If Exist("transhistjr.dat") Then KillFile "transhistjr.dat"
   If Exist("custlookup.dat") Then KillFile "custlookup.dat"
   If Exist("custByCat.dat") Then KillFile "custByCat.dat"
   If Exist("custrptsmenu.dat") Then KillFile "custrptsmenu.dat"

   For UBFrmCnt = Forms.Count - 1 To 0 Step -1
      Unload Forms(UBFrmCnt)
   Next
   DoEvents
   End
End Sub

Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
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

Public Sub KillFile(FileName As String)
'  On Local Error Resume Next
  If Exist(FileName$) Then 'added 7/24
    Kill FileName$
  End If
End Sub

Public Function CheckValDate(ValCheck As String)
  Dim Month As Integer, Day As Integer, Year As Integer
  Month = Val(Mid(ValCheck, 1, 2))
  Day = Val(Mid(ValCheck, 4, 2))
  Year = Val(Mid(ValCheck, 7, 4))
  'Checks date if Blank then won't check for valid date
  'and then checks each section, month, day and year
  'if any section wrong then returns false value
      If InStr(ValCheck, "_") <= 0 Then
          If ((Month > 0) And (Month < 13)) Then
              If Day > 0 And Day < 32 Then
                  If Year > 1919 And Year < 2099 Then
                      CheckValDate = True
                  End If
              End If
          End If
      End If
End Function

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

Public Function Exist(FileName$) As Boolean
  Dim FileHandle As Integer
  Dim TempSize As Long
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
Public Function CheckCitiDir(CityPakDir$) As Integer
  On Local Error GoTo Oh_Shucks
  Dim Handle As Integer
  'function verifies a valid path to the Citipak directory
  CityPakDir$ = QPTrim$(CityPakDir$)
  
  If Len(QPTrim$(CityPakDir$)) = 0 Then 'path not saved yet
    CheckCitiDir = 1
    Exit Function
  End If
  
  If Right$(CityPakDir$, 1) <> "\" Then 'adds a back slash if absent
    CityPakDir$ = CityPakDir$ + "\"
  End If
  
  Handle = FreeFile
  'If the next open statement cannot occur without an error then
  'the On Local Error statement above sends the function back as false
  Open CityPakDir$ + "TestCitiDir.TXT" For Binary As Handle
  'if we get here then all is well
  Close
  CheckCitiDir = True
  Exit Function
  
Oh_Shucks:
  CheckCitiDir = False
  Close
  
End Function

Function CheckFor2ManyDecimals(Text As String) As Boolean
  Dim cnt As Integer
  Dim DecCnt As Integer
  Dim StrLen As Long
  Dim thischar$
  'this function traps errors created when a user keys in a
  'decimal value and inadvertantly keys in more than 1 decimal
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Mid$(Text, cnt, 1)
    If thischar = "." Then DecCnt = DecCnt + 1 'counts decimals
  Next cnt
  If DecCnt > 1 Then 'if decimal count is more than 1 then process
  'accordingly
    CheckFor2ManyDecimals = True
  Else
    CheckFor2ManyDecimals = False
  End If
End Function
'Sub RPTSetupPRN(RPTNum, Handle)
'  Dim RPTPitch As Integer
'  Dim PrinterSetUpFile As Integer
''  Dim PrntType As PRNSetupRecType
'  Dim x As Integer
'  Dim PHandle As Integer
'  Dim DefPrinter As String
'  Dim PrnDef As String
'  Dim LineLen As Integer
'  Dim TextLine$
'  Dim Y As Integer
'  Dim z As Integer
'  Dim NextCommaPOS As Integer
'  Dim CodeStartPOS As Integer
'  Dim Codeline1$
'  Dim Codeline2$
'  'this sub coordinates the printing procedure so that any
'  'pitch data saved in the Printer setup screen for a
'  'particular report gets sent to the printer
'  For z = 1 To 10 'clear all existing codes
'    ToPrint1(z) = 0
'    ToPrint2(z) = 0
'  Next z
'  OpenPrinterSetupFile PrinterSetUpFile
'  Get PrinterSetUpFile, 1, PrntType
'  Close PrinterSetUpFile
'  DefPrinter = QPTrim$(PrntType.Printer)
'  'if a pitch isn't saved for this print job then by
'  'default the pitch becomes 10
'
'  If RPTNum = 123 Then GoTo SkipThis '123 is an arbitrary
'  'number used to signify the end of a report that tells
'  'this program to look for the reset codes
'
'  RPTPitch = PrntType.RPT(RPTNum) 'pitch is specified in the
'  'printer setup screen
'
'SkipThis:
'
'  GoSub GetPrinterCodes
'  If Len(Codeline1) Then 'CodeLine1 represents the reset codes
'  'because in the prprndf.dat file the reset codes come
'  'before the pitch codes
'  'at this point the proper codes have been determined and
'  'the select statement tells the printer which codes to use
'    Select Case Y
'      Case 1:
'        Print #Handle, Chr(ToPrint1(1));
'      Case 2:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2));
'      Case 3:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3));
'      Case 4:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4));
'      Case 5:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5));
'      Case 6:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6));
'      Case 7:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7));
'      Case 8:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8));
'      Case 9:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9));
'      Case 10:
'        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9)); Chr(ToPrint1(10));
'      Case Else:
'    End Select
'  ElseIf Len(Codeline2) Then 'CodeLine2 represents the pitch codes
'    Select Case Y 'Y = the number of codes
'      Case 1:
'        Print #Handle, Chr(ToPrint2(1));
'      Case 2:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2));
'      Case 3:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3));
'      Case 4:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4));
'      Case 5:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5));
'      Case 6:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6));
'      Case 7:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7));
'      Case 8:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8));
'      Case 9:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9));
'      Case 10:
'        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9)); Chr(ToPrint2(10));
'      Case Else:
'    End Select
'  End If
'
'  Exit Sub
'
'GetPrinterCodes:
'  PHandle = FreeFile
'  Open "PRData\Prprndf.dat" For Input As #PHandle  ' Open file.
'  Line Input #PHandle, TextLine   ' Read first line into TextLine.
'   'the second line is where individual printers start their codes
'   NextCommaPOS = 1
'
'   Do While Not EOF(PHandle) And NextCommaPOS <> 0  ' Loop until end of file.
'     Line Input #PHandle, TextLine   ' Read next line into Textline.
'     If TextLine = "@" + DefPrinter$ Then 'locate the default printer
'
'         If EOF(PHandle) Then Exit Do 'if for some reason we get to the end of the file
'         'then exit
'         If RPTNum = 123 Then '123 tells this code that we want the
'         'reset codes
'           Line Input #PHandle, TextLine 'read next line which by convention
'           'will always be the reset code line
'             LineLen = Len(TextLine)
'             Codeline1 = Mid(TextLine, 11, LineLen) 'by convention
'             '11 is where the first reset code begins in this line
'             CodeStartPOS = 1
'             Y = 1
'             Do
'               NextCommaPOS = InStr(CodeStartPOS, Codeline1, ",") 'look for comma
'               If NextCommaPOS = 0 Then 'if comma pos = 0 then we have no more commas
'                 LineLen = Len(Codeline1)
'                 ToPrint1(Y) = CInt(Mid(Codeline1, CodeStartPOS, 3))
'                 Exit Do 'we're at the end so exit loop
'               End If
'               ToPrint1(Y) = CInt(Mid(Codeline1, CodeStartPOS, 3)) 'look for a comma
'               CodeStartPOS = NextCommaPOS + 1 'start just behind the last comma
'               Y = Y + 1 ' advance y until no more commas found
'             Loop Until NextCommaPOS = 0
'             GoTo XIsOne 'jump to outer loop
'         End If
'       Do
'         Line Input #PHandle, TextLine 'look for pitch codes
'         If Mid(TextLine, 1, 2) = RPTPitch Then 'we found the proper pitch
'           LineLen = Len(TextLine)
'           Codeline2 = Mid(TextLine, 11, LineLen) 'read this line into
'           'Codeline2
'           Y = 1
'           CodeStartPOS = 1
'           Do
'             NextCommaPOS = InStr(CodeStartPOS, Codeline2, ",")
'             If NextCommaPOS = 0 Then 'no more commas in line
'               LineLen = Len(Codeline2)
'               ToPrint2(Y) = CInt(Mid(Codeline2, CodeStartPOS, 3))
'               Exit Do
'             End If
'             ToPrint2(Y) = CInt(Mid(Codeline2, CodeStartPOS, 3)) 'keep looking
'             'for commas
'             CodeStartPOS = NextCommaPOS + 1
'             Y = Y + 1
'           Loop Until NextCommaPOS = 0
'
'           Exit Do
'         End If
'        Loop Until NextCommaPOS = 0
'XIsOne:
'     End If 'ends if TextLine = @ + DefPrinter
'   Loop
'   Close #PHandle   ' Close file.
'   Return
'
'
'End Sub

'Public Function PromptSaveChanges(frm As Form) As SaveChangeOptions1
'  frmBLChangedWarning.Show vbModal, frm
'  PromptSaveChanges = frmBLChangedWarning.Selection
'  Unload frmBLChangedWarning
'End Function

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmBLViewPrint.ReportName = ReportFile$
   frmBLViewPrint.Caption = Title
   frmBLViewPrint.PgNum = PgNum
   frmBLViewPrint.cmdAlignment.Visible = False
   If ForceSBar Then
     frmBLViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmBLViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmBLViewPrint.cmdAlignment.Enabled = True
     frmBLViewPrint.AlignRpt = AlgnRptfile$
    Else
      frmBLViewPrint.cmdAlignment.Enabled = False
    End If
   frmBLViewPrint.Show 1
   doAlign = False
End Sub

Public Sub GetAcctStruct(GLFundLen%, GLAcctLen%, GLDetLen%)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetupRec(1) As GLSetupRecType
  
  SetUpRecLen = Len(GLSetupRec(1))
  If Exist("GLSETUP.DAT") Then
    SetupFile = FreeFile
    Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Else
    Exit Sub
  End If
  Get SetupFile, 1, GLSetupRec(1)
  Close SetupFile
  GLFundLen = GLSetupRec(1).FundLen
  GLAcctLen = GLSetupRec(1).AcctLen
  GLDetLen = GLSetupRec(1).DetLen
End Sub

Public Sub CreateCatCodeIdx()
  Dim BigNum$ ' As Double
  Dim ThisNum$ ' As Double
  Dim Thisx As Integer
  Dim SmallNum$ ' As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempNum As Integer
  Dim CodeHandle As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeRecLen As Integer
  Dim NumOfCodeRecs As Integer
  Dim CodeIdx As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CatCodeIdxType
  
  On Error GoTo ERRORSTUFF
  
  OpenCatCodeFile CodeHandle
  
  NumOfCodeRecs = LOF(CodeHandle) \ Len(CodeRec)
  ReDim TempCodeIdx(1 To NumOfCodeRecs) As CatCodeIdxType
  
  BigNum = 0
  For x = 1 To NumOfCodeRecs
    Get CodeHandle, x, CodeRec
    TempCodeIdx(x).CatCodeRec = x
    TempCodeIdx(x).CatCodeNum = QPTrim$(CodeRec.CatCode)
    ThisNum = QPTrim$(CodeRec.CatCode)
    If Val(ThisNum) > Val(BigNum) Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close CodeHandle
  
  BigNum = BigNum + "1"
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfCodeRecs
      ThisNum = TempCodeIdx(x).CatCodeNum
      If Val(ThisNum) < Val(SmallNum) Then
        SmallNum = ThisNum
        Thisx = x
      End If
    Next x
    HoldThis = TempCodeIdx(Nextx)
    TempCodeIdx(Nextx) = TempCodeIdx(Thisx)
    TempCodeIdx(Thisx) = HoldThis
    If Nextx = NumOfCodeRecs Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  KillFile "arcatcodeidx.dat"
  
  OpenCatCodeIdxFile CodeIdxHandle
  For x = 1 To NumOfCodeRecs
    CodeIdx = TempCodeIdx(x)
    Put CodeIdxHandle, x, CodeIdx
  Next x
  
  Close CodeIdxHandle
  
  Exit Sub
  
ERRORSTUFF:

   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "BLCommon", "CreateCatCodeIdx", Erl)
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
    ClearInUse PWcnt
    Terminate
  
End Sub
Public Sub CreateCustNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigName$
  Dim ThisName$
  Dim Thisx As Integer
  Dim SmallName$
  Dim TempName As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustNameIdxType
  Dim ThisCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).BillingName = QPTrim$(CustRec.BillName)
    ThisName = QPTrim$(CustRec.BillName)
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
      ThisName = TempCustIdx(x).BillingName
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
  
  KillFile "arcustnameidx.dat"
  'must kill the old file because if a customer is deleted
  'it still remains as a record...not deleting causes multiple
  'repeats of the last customer depending on how many customers
  'have been deleted
  OpenCustNameIdxFile CustIdxHandle
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "BLCommon", "CreateCustNameIdx", Erl)
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
    ClearInUse PWcnt
    Terminate
  
  
End Sub
Public Sub CreateLicNumIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigLic As Double
  Dim ThisLic As Double
  Dim Thisx As Integer
  Dim SmallLic As Double
  Dim TempLic As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim LicIdx As CustLicNumIdxType
  Dim LicIdxHandle As Integer
  Dim LicIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustLicNumIdxType
  Dim ThisCnt As Integer
  
  On Error GoTo ERRORSTUFF
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempLicIdx(1 To NumOfCustRecs) As CustLicNumIdxType
  
  BigLic = 999999999999#
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempLicIdx(ThisCnt).CustRec = x
    TempLicIdx(ThisCnt).LicNum = QPTrim$(CustRec.LICENSE)
    ThisLic = Val(CustRec.LICENSE)
    If ThisLic > BigLic Then
      BigLic = ThisLic
    End If
BadNum:
  Next x
  Close CustHandle
  
  SmallLic = BigLic
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt 'NumOfCustRecs
      ThisLic = Val(TempLicIdx(x).LicNum)
      If ThisLic < SmallLic Then
        SmallLic = ThisLic
        Thisx = x
      End If
    Next x
    HoldThis = TempLicIdx(Nextx)
    TempLicIdx(Nextx) = TempLicIdx(Thisx)
    TempLicIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do 'NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallLic = BigLic
  Loop
  
  KillFile "arlicnumidx.dat"
  OpenLicNumIdxFile LicIdxHandle
  For x = 1 To ThisCnt ' NumOfCustRecs
    LicIdx = TempLicIdx(x)
    Put LicIdxHandle, x, LicIdx
  Next x
  Close LicIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "BLCommon", "CreateLicNumIdx", Erl)
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
    ClearInUse PWcnt
    Terminate
  
  
End Sub

Public Sub CreateCustNumIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigNum As Integer
  Dim ThisNum As Integer
  Dim Thisx As Integer
  Dim SmallNum As Integer
  Dim TempNum As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNumIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustNumIdxType
  Dim ThisCnt As Integer
  
  On Error GoTo ERRORSTUFF
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustNumIdxType
  
  BigNum = 0
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo ItsDeleted
    If Not IsNumeric(CustRec.CustNumb) Then Exit For
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).CustNumb = CInt(CustRec.CustNumb)
    ThisNum = CInt(CustRec.CustNumb)
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
ItsDeleted:
  Next x
  Close CustHandle
  
  If x < NumOfCustRecs Then
    frmBLMessageBoxJr.Label1.Caption = "Customer number " + QPTrim$(CustRec.CustNumb) + " (" + QPTrim$(CustRec.CustName) + ") contains a non-numeric character which prevents proper sorting. Please re-issue this number as a numeric before continuing."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    DoEvents
    Close
    Exit Sub
  End If
  
  SmallNum = BigNum + 1
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt ' NumOfCustRecs
      ThisNum = CInt(TempCustIdx(x).CustNumb)
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        Thisx = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(Thisx)
    TempCustIdx(Thisx) = HoldThis
    If Nextx = ThisCnt Then Exit Do 'NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  
  KillFile "arcustnumidx.dat"
  'must delete the old file because if a customer is deleted then
  'it still holds a record...not deleting the old file causes
  'the last customer to be repeated as many times as there have
  'been deleted
  OpenCustNumIdxFile CustIdxHandle
  For x = 1 To ThisCnt ' NumOfCustRecs
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "BLCommon", "CreateCustNumIdx", Erl)
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
    ClearInUse PWcnt
    Terminate
  
End Sub

Public Sub CreateCustSearchNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigName$
  Dim ThisName$
  Dim Thisx As Integer
  Dim SmallName$
  Dim TempName As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustSearchNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustSearchNameIdxType
  Dim ThisCnt As Integer
  
  On Error GoTo ERRORSTUFF
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustSearchNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).SortName = QPTrim$(CustRec.SortName)
    ThisName = QPTrim$(CustRec.SortName)
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
    For x = Nextx To ThisCnt ' NumOfCustRecs
      ThisName = TempCustIdx(x).SortName
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
  
  KillFile "arsrhidx.dat"
  'must delete the old file because if a customer is deleted
  'then the deleted customer still holds a record...if the old
  'file is not deleted then it causes the last customer on the
  'index to be repeated the same number of times a customer has been
  'deleted
  OpenSrchNameIdxFile CustIdxHandle
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "BLCommon", "CreateCustSearchNameIdx", Erl)
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
    ClearInUse PWcnt
    Terminate
  
End Sub

Public Sub ReLinkTransactions(frm As Form)
  Dim TransRec As ARTransRecType
  Dim TranFile As Integer
  Dim NumCRec&
  Dim CustFile As Integer
  Dim ARCust As ARCustRecType
  Dim NumTRec&
  Dim Ccnt&
  Dim TCnt&
  Dim CustRec&
  Dim BadTran As Integer
  Dim AllTrans&
  
  ReDim ARTran(1 To 2) As ARTransRecType
  
  OpenCustFile CustFile
  
  NumCRec& = LOF(CustFile) / Len(ARCust)

  OpenTransFile TranFile
  NumTRec& = LOF(TranFile) / Len(TransRec)
  
  AllTrans = NumCRec + NumTRec
  
  frmBLShowPctComp.Label1 = "Relinking Transaction Records"
  frmBLShowPctComp.Show
  frmBLShowPctComp.cmdCancel.Visible = False
  EnableCloseButton frm.hwnd, False

  For Ccnt& = 1 To NumCRec&
    Get CustFile, Ccnt&, ARCust
    ARCust.FirstTrans = 0
    ARCust.LastTrans = 0
    Put CustFile, Ccnt&, ARCust
    frmBLShowPctComp.ShowPctComp Ccnt, AllTrans 'NumCRec
  Next
  
  For TCnt& = 1 To NumTRec&
    Get TranFile, TCnt&, ARTran(1)
    CustRec = Val(ARTran(1).CustomerNumber)
    If (CustRec& > 0) And (CustRec& <= NumCRec&) Then
      Get CustFile, CustRec, ARCust
      If ARCust.LastTrans = 0 Then
        ARCust.FirstTrans = TCnt&
        ARCust.LastTrans = TCnt&
        Put CustFile, CustRec&, ARCust
        ARTran(1).NextTrans = 0
        Put TranFile, TCnt&, ARTran(1)
      Else
        Get TranFile, ARCust.LastTrans, ARTran(2)  'get old last tr
        ARTran(2).NextTrans = TCnt&                    'point it to next tr
        Put TranFile, ARCust.LastTrans, ARTran(2)  'put it back
        ARCust.LastTrans = TCnt&                    'set new cust last TR
        Put CustFile, CustRec&, ARCust          'put it back
        ARTran(1).NextTrans = 0
        Put TranFile, TCnt&, ARTran(1)
      End If
      frmBLShowPctComp.ShowPctComp TCnt + NumCRec&, AllTrans 'NumTRec
    End If
NoGood:
  Next
  
  Unload frmBLShowPctComp
  EnableCloseButton frm.hwnd, True

  Close
  frmBLSucSave.Label1.Caption = "Transactions have been relinked successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal

End Sub

Public Function TrimPhone(PhoneNumber$) As String
  Dim x As Integer
  Dim PhoneLen As Integer
  Dim thischar$
  
  PhoneLen = Len(PhoneNumber)
  For x = 1 To PhoneLen
    thischar = Mid(PhoneNumber, x, 1)
    If IsNumeric(thischar) = True Then
      TrimPhone = TrimPhone + thischar
    End If
  Next x
  
End Function

Public Function FirstLicenseNum() As String
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim BigNum As Double
  Dim SmallNum As String
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  Dim ThisNum As Double
  
  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  If NumOfCustRecs = 0 Then
    Close
    FirstLicenseNum = "0"
    Exit Function
  End If
  BigNum = "0"
  For x = 1 To NumOfCustRecs
    Get CHandle, x, CustRec
    If Len(QPTrim$(CustRec.LICENSE)) = 0 Or QPTrim$(CustRec.LICENSE) = "AUTO" Then GoTo Invalid
    If Not IsNumeric(CustRec.LICENSE) Then GoTo Invalid
    ThisNum = CDbl(CustRec.LICENSE)
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
Invalid:
  Next x
  
  FirstLicenseNum = CStr(BigNum)
  
End Function

Public Function MakeDay$(DayNum As Integer)
  
  Select Case DayNum
  Case 1
    MakeDay$ = "Sunday"
  Case 2
    MakeDay$ = "Monday"
  Case 3
    MakeDay$ = "Tuesday"
  Case 4
    MakeDay$ = "Wednesday"
  Case 5
    MakeDay$ = "Thursday"
  Case 6
    MakeDay$ = "Friday"
  Case 7
    MakeDay$ = "Saturday"
  End Select
  
End Function

Public Function MakeLongDate$(PDate$)
  Dim DayNum As Integer
  Dim DayName$
  Dim MonthName$
  
  DayName$ = FindDayName(PDate$)
  MonthName$ = MakeMonth$(PDate$)
  MakeLongDate$ = DayName$ + ", " + MonthName$ + " " + Mid$(PDate$, 4, 2) + ", " + Right$(PDate$, 4)
End Function

Public Function MakeMonth$(TDate$)
  Dim Month As Integer
  
  Month = Val(Left$(TDate$, 2))
  Select Case Month
  Case 1
    MakeMonth$ = "January"
  Case 2
    MakeMonth$ = "February"
  Case 3
    MakeMonth$ = "March"
  Case 4
    MakeMonth$ = "April"
  Case 5
    MakeMonth$ = "May"
  Case 6
    MakeMonth$ = "June"
  Case 7
    MakeMonth$ = "July"
  Case 8
    MakeMonth$ = "August"
  Case 9
    MakeMonth$ = "September"
  Case 10
    MakeMonth$ = "October"
  Case 11
    MakeMonth$ = "November"
  Case 12
    MakeMonth$ = "December"
  End Select
End Function

Public Function FindDayName(ThisDate$) As String
   Dim FoundDay As Integer
   
   FoundDay = Date2Num(ThisDate)
   FoundDay = FoundDay Mod 7 'find number of days from
   'starting point
   FoundDay = FoundDay + 2 'starting point was a Monday so
   'add 2
   If FoundDay = 8 Then FoundDay = 1 'the highest mod would be
   '6 so 6 + 2 = 8 which would indicate 6 days from the
   'starting point (Monday) = Sunday
   Select Case FoundDay
     Case 1
       FindDayName = "Sunday"
     Case 2
       FindDayName = "Monday"
     Case 3
       FindDayName = "Tuesday"
     Case 4
       FindDayName = "Wednesday"
     Case 5
       FindDayName = "Thursday"
     Case 6
       FindDayName = "Friday"
     Case 7
       FindDayName = "Saturday"
     Case Else
       FindDayName = "Unknown"
  End Select
  
End Function

Public Function BegBalCheck(CustNum$, ByVal ONum$, ByRef ThisRec As Integer) As Integer
  Dim OHandle As Integer
  Dim OperRec As CitiPassType 'CMOperRecType
  Dim NumOperRecs As Integer
  Dim x As Integer
  Dim Operator$
  Dim y As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  
  BegBalCheck = 1
  OpenCitiPassFile OHandle, NumOperRecs
'  OpenOperRecFile OHandle
'  NumOperRecs = LOF(OHandle) / Len(OperRec)
  
  If NumOperRecs = 0 Then
    Close
    Exit Function
  End If
  
  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
'      OpIdx(x) = OperRec.OperatorNumber
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle
  
  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    If Exist(BLPayFileName + Operator$ + ".DAT") Then
      OpenPayFile PayHandle, OpIdx(x) 'look thru all operator files
      NumOfPayRecs = LOF(PayHandle) / Len(EditPayRec)
      For y = 1 To NumOfPayRecs 'if you find this customer already
      'has
        Get PayHandle, y, EditPayRec
        If QPTrim$(CustNum$) = QPTrim$(EditPayRec.CustNumber) Then
          If QPTrim$(Operator$) = QPTrim$(Str(ONum)) Then
            frmBLMessageBoxJrWOpts.Label1.Caption = "An unposted transaction is in progress for this customer. Do you want to edit this transaction?"
            frmBLMessageBoxJrWOpts.Label1.Top = 800
            frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Edit"
            frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC No"
            frmBLMessageBoxJrWOpts.Show vbModal
            If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
              Unload frmBLMessageBoxJrWOpts
              MainLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + QPTrim$(CustNum$) + " on " + MakeRegDate(EditPayRec.TranDate) + " and opted to continue with the payment edit.")
              BegBalCheck = 2
              ONum = "Operator"
              ThisRec = y
              Close
            Else
              Unload frmBLMessageBoxJrWOpts
              MainLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + QPTrim$(CustNum$) + " on " + MakeRegDate(EditPayRec.TranDate) + " and opted to exit the payment edit.")
              BegBalCheck = 4
            End If
            x = NumOperRecs
            Exit For
          Else
            frmBLMessageBoxJr.Label1.Caption = "An unposted transaction is in progress by operator number " + Operator$ + " on " + MakeRegDate(EditPayRec.TranDate) + ". Edit attempt is aborted."
            frmBLMessageBoxJr.Label1.Top = 800
            frmBLMessageBoxJr.Show vbModal
            BegBalCheck = 4
            MainLog ("Operator # " + QPTrim$(Str(ONum)) + " warned that a beginning balance transaction existed for customer # " + QPTrim$(CustNum$) + " by operator #" + QPTrim$(Operator$) + " on " + MakeRegDate(EditPayRec.TranDate) + " and edit attempt was aborted.")
            Exit For
          End If
        End If
      Next y
    End If
  Next x
  Close
End Function

Public Function GetCatRecNum(BillCat$) As Integer
  Dim x As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  
  GetCatRecNum = 0
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
    
  For x = 1 To CatRecNums
    Get CHandle, x, CatRec
      If QPTrim$(CatRec.CatCode) = QPTrim$(BillCat) Then
        GetCatRecNum = x
        Exit For
      End If
  Next x
  Close CHandle
  
End Function

Public Function GetGLRecNum(GLNum$) As Long
  Dim NumOfGLRecs As Integer
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  Dim x As Long
  
  GetGLRecNum = 0
  OpenGLAcctFile AcctHandle
  NumOfGLRecs = LOF(AcctHandle) / Len(GLAcctRec)
  For x = 1 To NumOfGLRecs
    Get AcctHandle, x, GLAcctRec
    If QPTrim$(GLAcctRec.Num) = GLNum Then
      If GLAcctRec.Deleted = 0 Then
        GetGLRecNum = x
        Exit For
      End If
    End If
  Next x
  Close AcctHandle

End Function

Public Function GetGLNum(GLNum As Long) As String
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  
  GetGLNum = ""
  If GLNum = 0 Then
    Exit Function
  End If
  OpenGLAcctFile AcctHandle
    Get AcctHandle, GLNum, GLAcctRec
    If GLAcctRec.Deleted = 0 Then
      GetGLNum = QPTrim$(GLAcctRec.Num)
    End If
  Close AcctHandle

End Function

Public Function EmpInLicProcess(EmpNum$) As Boolean
  Dim x As Long
  Dim TempRec As TempTransPostType
  Dim TempHandle As Integer
  Dim NumOfTempRecs As Long
  
  EmpInLicProcess = False
  OpenTempPostFile TempHandle
  NumOfTempRecs = LOF(TempHandle) / Len(TempRec)
  
  If NumOfTempRecs = 0 Then
    Close TempHandle
    Exit Function
  End If
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
    If QPTrim$(EmpNum$) = QPTrim$(TempRec.CustomerNumber) Then
      EmpInLicProcess = True
      Exit For
    End If
  Next x
  Close TempHandle
      
End Function

Public Function EmpInPenProcess(EmpNum$) As Boolean
  Dim PenTrans As TempPenaltyCharges
  Dim TPHandle As Integer
  Dim NumOfPen As Integer
  Dim x As Integer
  
  EmpInPenProcess = False
  OpenPenTransFile TPHandle
  NumOfPen = LOF(TPHandle) \ Len(PenTrans)
  For x = 1 To NumOfPen
    Get TPHandle, x, PenTrans
    If QPTrim$(EmpNum$) = QPTrim$(PenTrans.CustomerNumber) Then
      EmpInPenProcess = True
      Exit For
    End If
  Next x
  Close TPHandle
  
End Function
Public Function EmpInPayProcess(EmpNum$) As Boolean
  Dim OHandle As Integer
  Dim OperRec As CitiPassType 'CMOperRecType
  Dim NumOperRecs As Integer
  Dim Operator$
  Dim y As Integer, x As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  Dim OCnt As Integer
  
  EmpInPayProcess = False
  OpenCitiPassFile OHandle, NumOperRecs
  If NumOperRecs = 0 Then
    Close OHandle
    Return
  End If

  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
      'load an array with the operator numbers
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle
  
  ReDim InPayCnt(1 To 1) As String
  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    If Exist(BLPayFileName + Operator$ + ".DAT") Then
      'if the file above exists then this operator has
      'saved at least one transaction
      OpenPayFile PayHandle, OpIdx(x) 'look thru all operator files
      NumOfPayRecs = LOF(PayHandle) / Len(EditPayRec)
      For y = 1 To NumOfPayRecs
        Get PayHandle, y, EditPayRec
        If QPTrim$(EditPayRec.CustNumber) = "" Then GoTo Deleted
        If QPTrim$(EditPayRec.CustNumber) = QPTrim$(EmpNum$) Then
          EmpInPayProcess = True
          Exit For
        End If
Deleted:
      Next y
    End If
SkipIt:
    Close PayHandle
    If EmpInPayProcess = True Then Exit For
  Next x
  
End Function

Public Function GetCatDesc(CatNum$) As String
  Dim x As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  
  GetCatDesc = ""
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
  If CatRecNums = 0 Then Exit Function
  For x = 1 To CatRecNums
    Get CHandle, x, CatRec
      If QPTrim$(CatRec.CatCode) = QPTrim$(CatNum$) Then
        GetCatDesc = QPTrim$(CatRec.CODEDESC)
        Exit For
      End If
  Next x
  Close CHandle
  
End Function
  
Public Function CatInTempFile(ByVal ThisCat As String, ByRef CustCnt As Integer) As Boolean
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  
  CustCnt = 0
  CatInTempFile = False
  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  For x = 1 To NumOfCustRecs
    Get CHandle, x, CustRec
    If QPTrim$(CustRec.BILLCAT1) = QPTrim(ThisCat) Then
      If EmpInLicProcess(QPTrim$(CustRec.CustNumb)) Then
        CustCnt = CustCnt + 1
      End If
    ElseIf QPTrim$(CustRec.BILLCAT2) = QPTrim(ThisCat) Then
      If EmpInLicProcess(QPTrim$(CustRec.CustNumb)) Then
        CustCnt = CustCnt + 1
      End If
    ElseIf QPTrim$(CustRec.BILLCAT3) = QPTrim(ThisCat) Then
      If EmpInLicProcess(QPTrim$(CustRec.CustNumb)) Then
        CustCnt = CustCnt + 1
      End If
    ElseIf QPTrim$(CustRec.BILLCAT4) = QPTrim(ThisCat) Then
      If EmpInLicProcess(QPTrim$(CustRec.CustNumb)) Then
        CustCnt = CustCnt + 1
      End If
    ElseIf QPTrim$(CustRec.BILLCAT5) = QPTrim(ThisCat) Then
      If EmpInLicProcess(QPTrim$(CustRec.CustNumb)) Then
        CustCnt = CustCnt + 1
      End If
    End If
  Next x
  
  Close CHandle
  
  If CustCnt > 0 Then
    CatInTempFile = True
  End If
End Function

Public Function GetCodeDesc(CatCode$) As String
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim x As Integer
  
  GetCodeDesc = ""
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  If NumOfARCatRecs = 0 Then Exit Function
  For x = 1 To NumOfARCatRecs
    Get CodeHandle, x, CodeRec
    If QPTrim$(CatCode$) = QPTrim$(CodeRec.CatCode) Then
      GetCodeDesc = QPTrim$(CodeRec.CODEDESC)
      Close CodeHandle
      Exit Function
    End If
  Next x
  Close CodeHandle
End Function

Public Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP = FreeFile
  lenRP = Len(RcptPrnFile)
'  If Exist("C:\RcptPrn.dat") Then
'    Open "c:\RcptPrn.dat" For Random Shared As RP Len = lenRP
  If Exist(RcptFileName$) Then '2/14/08
    Open RcptFileName$ For Random Shared As RP Len = lenRP '2/14/08
    Get RP, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
      RecpDef = 1
    End If
    Close
  Else
    frmBLMessageBoxJr.Label1.Caption = "RECEIPT SETUP FILE NOT FOUND. Payment receipts will not be able to print. Receipt setup can be found on the Citipak main menu."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.cmdExit.Text = "ESC OK"
    frmBLMessageBoxJr.Show vbModal
    Close
    RecpDef = 99
  End If
End Sub

Public Function Check4ValidCatNum(ThisCat$) As Boolean
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim CatCodeCnt As Integer
  Dim x As Integer
  Dim FoundIt As Boolean
  
  Check4ValidCatNum = True
  FoundIt = False
  
  If Not IsNumeric(ThisCat) Then Exit Function
  
  OpenCatCodeIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then
    Close CodeIdxHandle
    Exit Function
  End If
  
  ReDim IdxRec(1 To CodeIdxRecNum) As Integer

  ReDim CodeIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle
  
  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)
  
  If CatCodeCnt = 0 Then
    Close CHandle
    Exit Function
  End If
  
  For x = 1 To CodeIdxRecNum
    Get CHandle, CodeIdx(x), CodeRec
      If QPTrim$(CodeRec.CatCode) = ThisCat Then
        FoundIt = True
        Exit For
      End If
  Next x
  
  Close CHandle
  
  If FoundIt = False Then Check4ValidCatNum = False
  
End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
  
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function

Public Function DirExists(ByVal strDirName As String) As Boolean
  On Error Resume Next
  
  Dim strFileName As String

  strFileName = strDirName & "\Nul"

  If (FileExists(strFileName)) Then
    DirExists = True
  Else
    DirExists = False
  End If
End Function

Public Sub Terminate2Shell()
   Dim UBFrmCnt As Integer
   ' Loop through the forms collection and unload each form.
   Close
   For UBFrmCnt = Forms.Count - 1 To 0 Step -1
       Unload Forms(UBFrmCnt)
   Next
   DoEvents
   End
End Sub

Public Sub GetTemp()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  
  'lentemp = Len(Tempfile)
  Tempfile = FreeFile
'  Open "c:\PassTemp.dat" For Random Shared As Tempfile ' Len = lentemp
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp 2/14/08
  Get Tempfile, 1, PassTemp
  PWUser = QPTrim(PassTemp.UserName)
  PWcnt = PassTemp.usernum
  Close

End Sub

Public Sub SetToGo()
  Dim Tempfile As Integer, lentemp As Integer
  Dim PassTemp As CitiPassTempType
  
  Tempfile = FreeFile
'  Open "c:\PassTemp.dat" For Random Shared As Tempfile ' Len = lentemp
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp '2/14/08
  If PWcnt = -3 Then PWcnt = 0
  PassTemp.usernum = PWcnt
  PassTemp.UserName = PWUser
  PassTemp.frommdl = 1
  Put Tempfile, 1, PassTemp
  Close
End Sub

