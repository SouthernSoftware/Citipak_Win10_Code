Attribute VB_Name = "FACommon"
Option Explicit
  
  Public Const gstrcProgName As String = "FixedAssets.exe"
  
  Public GRecNum As Long
  Public GCodeNum As Long
  Public GDeptNum As Long
  Public GFundNum As Long
  Public GTagNum As String
  Public ThisTag As String
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
  Public DeptList() As String
  Public NumOfDepts As Integer
  Public ItemChangeFlag As Boolean
  Public VhclTempDsplFlag As Boolean
  Public FocusOn As Integer
  Public AddItemFlag As Boolean
  Public FromFA As Boolean
  
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
            Public Const PRData = "prdata\"
   Public Const FASetUpFileName = "FASETUP.DAT"
    Public Const FAItemFileName = "FAITEMS.DAT"
   Public Const FAAssetCodeName = "FACODES.DAT"
    Public Const FADeptCodeName = "FADEPTCD.DAT"
    Public Const FAFundCodeName = "FAFUNDCD.DAT"
     Public Const FAYearEndName = "FAYEAR.DAT"
    Public Const FADeprEditName = "FADPREDT.DAT"
'Public Const FATempVhclDataName = "FATMPVHC.DAT"
'  Public Const PrinterSetUpFile = "PRPRNSET.DAT"
'     Public Const GLAcctIdxFile = "BAACCTDX.DAT"
'    Public Const JGLAcctIdxFile = "GLACCT.IDX"
'      Public Const AcctFileName = "GLACCT.DAT"
'     Public Const TransFileName = "GLTRANS.DAT"
        Public Const TagIdxName = "FATAGIDX.DAT"
        Public Const AssIdxName = "FAASSIDX.DAT"
       Public Const DeptIdxName = "FADEPIDX.DAT"
       Public Const FundIdxName = "FAFNDIDX.DAT"
       Public Const DprHistFileName = "FADPRHIST.DAT"
       Public Const TempDprFileName = "FATEMPDPR.DAT"
       Public Const TempDispDateName = "FATEMPDISPDATE.DAT"
       Public Const PrepostDsplName = "FAPREPOSTDSPL"
Public Sub Terminate()
  Dim UBFrmCnt As Integer
  Close
  KillFile "itemmaintmenu.dat"
  KillFile ("dprhistbyitemrpt.dat")
  KillFile ("dprhistrpt.dat")
  KillFile ("valrpt.dat")
  KillFile ("itemchecklist.dat")
  KillFile (TempDprFileName)
  KillFile ("masteritemlistopen.dat")
  KillFile ("editdeptopen.dat")
  KillFile ("edititemopen.dat")
  KillFile ("Wrntyrpt.dat")
  KillFile ("taglistopen.dat")
  KillFile ("assetbycoderpt.dat")
  KillFile ("fromItemMaintMenu.dat")
  KillFile "fromBuildDep.dat"
  KillFile ("assetbyfundrpt.dat")
' Loop through the forms collection and unload each form.
'   ClearInUsePRReg PWcnt 'we want this intact so if another user
 'gets in payroll the "inuse" warning will pop up
  Close
  For UBFrmCnt = Forms.Count - 1 To 0 Step -1
    Unload Forms(UBFrmCnt)
  Next
  DoEvents
  End
  
End Sub
Public Sub OpenPrePostDsplData(PrePostDsplHandle As Integer, ThisDate As Integer)
  Dim PrePostDsplRec As PrePostDsplType
  Dim PrePostDsplLen As Integer
  Dim StrDate As String
  
  StrDate = CStr(ThisDate)
  PrePostDsplLen = Len(PrePostDsplRec)
  PrePostDsplHandle = FreeFile
  Open PrepostDsplName + StrDate$ + ".DAT" For Random Shared As PrePostDsplHandle Len = PrePostDsplLen
End Sub

Public Sub OpenTempDisposedDate(TempDisposedDateHandle As Integer)
  Dim TempDisposedDateRec As TempDisposedOfDate
  Dim TempDisposedDateLen As Integer
  TempDisposedDateLen = Len(TempDisposedDateRec)
  TempDisposedDateHandle = FreeFile
  Open TempDispDateName For Random Shared As TempDisposedDateHandle Len = TempDisposedDateLen
End Sub
Public Sub OpenTempDprFile(TempDprHandle As Integer)
  Dim TempDprRec As DprSortIdxType
  Dim TempDprLen As Integer
  TempDprLen = Len(TempDprRec)
  TempDprHandle = FreeFile
  Open TempDprFileName For Random Shared As TempDprHandle Len = TempDprLen
End Sub
Public Sub OpenDprHistFile(DprHistHandle As Integer)
  Dim DprHistRec As DprHistType
  Dim DprHistLen As Integer
  DprHistLen = Len(DprHistRec)
  DprHistHandle = FreeFile
  Open DprHistFileName For Random Shared As DprHistHandle Len = DprHistLen
End Sub
'Public Sub OpenTempVhclFile(TempVhclHandle As Integer)
'  Dim TempVhclRec As TempVHCLDataType
'  Dim TempVhclLen As Integer
'  TempVhclLen = Len(TempVhclRec)
'  TempVhclHandle = FreeFile
'  Open FATempVhclDataName For Random Shared As TempVhclHandle Len = TempVhclLen
'End Sub
Public Sub OpenDeprEditFile(DeprEditHandle As Integer)
  Dim DeprEditRec As FADepFileType
  Dim DeprEditLen As Integer
  DeprEditLen = Len(DeprEditRec)
  DeprEditHandle = FreeFile
  Open FADeprEditName For Random Shared As DeprEditHandle Len = DeprEditLen
End Sub

Public Sub OpenYearFile(YearHandle As Integer)
  Dim YearRec As FAYearEndType
  Dim YearLen As Integer
  YearLen = Len(YearRec)
  YearHandle = FreeFile
  Open FAYearEndName For Random Shared As YearHandle Len = YearLen
End Sub

Public Sub OpenFundIdxFile(FundIdxHandle As Integer)
  Dim FundIdxRec As FundNumbSortIdxType
  Dim FundIdxLen As Integer
  FundIdxLen = Len(FundIdxRec)
  FundIdxHandle = FreeFile
  Open FundIdxName For Random Shared As FundIdxHandle Len = FundIdxLen
End Sub
      
Public Sub OpenDeptIdxFile(DeptIdxHandle As Integer)
  Dim DeptIdxRec As DeptNumbSortIdxType
  Dim DeptIdxLen As Integer
  DeptIdxLen = Len(DeptIdxRec)
  DeptIdxHandle = FreeFile
  Open DeptIdxName For Random Shared As DeptIdxHandle Len = DeptIdxLen
End Sub
Public Sub OpenAssIdxFile(AssIdxHandle As Integer)
  Dim AssIdxRec As ACNumbSortIdxType
  Dim AssIdxLen As Integer
  AssIdxLen = Len(AssIdxRec)
  AssIdxHandle = FreeFile
  Open AssIdxName For Random Shared As AssIdxHandle Len = AssIdxLen
End Sub
Public Sub OpenTagIdxFile(TagIdxHandle As Integer)
  Dim TagIdxRec As TagNumbSortIdxType
  Dim TagIdxLen As Integer
  TagIdxLen = Len(TagIdxRec)
  TagIdxHandle = FreeFile
  Open TagIdxName For Random Shared As TagIdxHandle Len = TagIdxLen
End Sub

Public Sub OpenFAFundCodeFile(FAFundCodeHandle As Integer)
  Dim FAFundCodeRec As FAFundCodeType
  Dim FAFundCodeRecLen As Integer
  FAFundCodeRecLen = Len(FAFundCodeRec)
  FAFundCodeHandle = FreeFile
  Open FAFundCodeName For Random Shared As FAFundCodeHandle Len = FAFundCodeRecLen
End Sub
Public Sub OpenFADeptCodeFile(FADeptCodeHandle As Integer)
  Dim FADeptCodeRec As FADeptCodeType
  Dim FADeptCodeRecLen As Integer
  FADeptCodeRecLen = Len(FADeptCodeRec)
  FADeptCodeHandle = FreeFile
  Open FADeptCodeName For Random Shared As FADeptCodeHandle Len = FADeptCodeRecLen
End Sub
Public Sub OpenFACodeNameFile(FACodeNameHandle As Integer)
  Dim FACodeNameRec As FAAssetCodeRecType
  Dim FACodeNameRecLen As Integer
  FACodeNameRecLen = Len(FACodeNameRec)
  FACodeNameHandle = FreeFile
  Open FAAssetCodeName For Random Shared As FACodeNameHandle Len = FACodeNameRecLen
End Sub
Public Sub OpenFASetUpFile(FASetUpHandle As Integer)
  Dim FASetUpRec As FASetupRecType
  Dim FASetUpRecLen As Integer
  FASetUpRecLen = Len(FASetUpRec)
  FASetUpHandle = FreeFile
  Open FASetUpFileName For Random Shared As FASetUpHandle Len = FASetUpRecLen
End Sub
Public Sub OpenFAItemFile(FAItemHandle As Integer)
  Dim FAItemRec As FAItemRecType
  Dim FAItemRecLen As Integer
  FAItemRecLen = Len(FAItemRec)
  FAItemHandle = FreeFile
  Open FAItemFileName For Random Shared As FAItemHandle Len = FAItemRecLen
End Sub
'Public Sub OpenPrinterSetupFile(PrinterSUFHandle As Integer)
'  Dim PrinterSUFRec As PRNSetupRecType
'  Dim PrinterSUFRecLen As Integer
'  PrinterSUFRecLen = Len(PrinterSUFRec)
'  PrinterSUFHandle = FreeFile
'  Open PrinterSetUpFile For Random Shared As PrinterSUFHandle Len = PrinterSUFRecLen
'End Sub

Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim ThisChar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    ThisChar = Asc(Mid$(Text, cnt, 1))
    If ThisChar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
  End Function

Public Function ReplaceString$(Text As String, ChangeThis As String, ToThis As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim NewText As String
  Dim ThisChar$
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
    ThisChar = Mid$(Text, cnt, 1) 'step thru text a letter at a time
    CTChar = Mid$(Text, cnt, CTLen) 'starting with the current letter
    'read ahead the length of the text "change this"
    If CTChar = ChangeThis Then 'if we find the "change this" in the
    'text
      NewText = NewText + ToThis 'assign the length of CTChar to "ToThis"
      'inside the rebuilt new text
      cnt = cnt + BigLen - 1 'advance count to compensate for the addition of
      'CTChar
    Else
      NewText = NewText + ThisChar 'build new text one letter at a time
    End If
  Next
  ReplaceString$ = Trim$(NewText) 'rim out the new text
  Text = ReplaceString$ 'old text is now new text
End Function

Public Sub KillFile(FileName As String)
  On Local Error Resume Next
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
  Dim ThisChar$
  'this function traps errors created when a user keys in a
  'decimal value and inadvertantly keys in more than 1 decimal
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    ThisChar = Mid$(Text, cnt, 1)
    If ThisChar = "." Then DecCnt = DecCnt + 1 'counts decimals
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
'  Dim PrntType As PRNSetupRecType
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
'   Do While Not eof(PHandle) And NextCommaPOS <> 0  ' Loop until end of file.
'     Line Input #PHandle, TextLine   ' Read next line into Textline.
'     If TextLine = "@" + DefPrinter$ Then 'locate the default printer
'
'         If eof(PHandle) Then Exit Do 'if for some reason we get to the end of the file
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

Public Function PromptCompletelyDelete(frm As Form) As CompletelyDeleteOption
  frmWarnCompletelyDelete.Show vbModal, frm
  PromptCompletelyDelete = frmWarnCompletelyDelete.Selection
  Unload frmWarnCompletelyDelete

End Function
Public Function PromptNotSubsequent(frm As Form) As NotSubsequentOption
  frmWarnYrNotSubsequent.Show vbModal, frm
  PromptNotSubsequent = frmWarnYrNotSubsequent.Selection
  Unload frmWarnYrNotSubsequent

End Function

Public Function PromptBadAssetCodeNum(frm As Form) As BadFACodeNumOption
  frmFABadAssetCodeNum.Show vbModal, frm
  PromptBadAssetCodeNum = frmFABadAssetCodeNum.Selection
  Unload frmFABadAssetCodeNum

End Function

Public Function PromptBadTagNum(frm As Form) As BadFATagNumOption
  frmFABadTagNum.Show vbModal, frm
  PromptBadTagNum = frmFABadTagNum.Selection
  Unload frmFABadTagNum

End Function
Public Function PromptWarnOverWrite(frm As Form) As WarnOption
  frmFAWarnOverWriteCode.Show vbModal, frm
  PromptWarnOverWrite = frmFAWarnOverWriteCode.Selection
  Unload frmFAWarnOverWriteCode
End Function

Public Function PromptSaveChanges(frm As Form) As SaveChangeOptions1
  frmFAChangedWarning.Show vbModal, frm
  PromptSaveChanges = frmFAChangedWarning.Selection
'  Unload frmFAChangedWarning
End Function

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmFAViewPrint.ReportName = ReportFile$
   frmFAViewPrint.Caption = Title
   frmFAViewPrint.PgNum = PgNum
   frmFAViewPrint.cmdAlignment.Visible = False
   If ForceSBar Then
     frmFAViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmFAViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmFAViewPrint.cmdAlignment.Enabled = True
     frmFAViewPrint.AlignRpt = AlgnRptfile$
    Else
      frmFAViewPrint.cmdAlignment.Enabled = False
    End If
   frmFAViewPrint.Show 1
   doAlign = False
End Sub

Public Sub CreateDprIdx(ThisYear$, ByRef ValidCnt As Long)
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempNum As Integer
  Dim DprRec As DprHistType
  Dim DprRecLen As Integer
  Dim DPRHandle As Integer
  Dim NumOfRecs As Long
  Dim DprIdx As DprSortIdxType
  Dim DprIdxHandle As Integer
  Dim DprIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As DprSortIdxType
  
  'used in depreciation history by year to print
  'only data from the requested year
  OpenDprHistFile DPRHandle
  NumOfRecs = LOF(DPRHandle) \ Len(DprRec)
  If NumOfRecs = 0 Then
    ValidCnt = -1
    Close
    Exit Sub
  End If

  ValidCnt = 0
  ReDim TempDprIdx(1 To 1) As DprSortIdxType
  BigNum = 0
  'this for loop builds a depreciation record array and it
  'also finds the largest tag number so a sort can start
  'below
  frmFAShowPctComp.Label1 = "Indexing Item Data"
  frmFAShowPctComp.cmdCancel.Visible = False
  frmFAShowPctComp.Show
  EnableCloseButton frmFADprHistRpt.hwnd, False
  frmFADprHistRpt.cmdExit.Enabled = False
  frmFADprHistRpt.cmdProcess.Enabled = False
  DoEvents
  For x = 1 To NumOfRecs
    Get DPRHandle, x, DprRec
    If QPTrim$(DprRec.DprYear) <> ThisYear$ Then GoTo BadNum
    ValidCnt = ValidCnt + 1 'keep track of valid assets
    TempDprIdx(ValidCnt).DprRecNum = x 'load array
    TempDprIdx(ValidCnt).DprNumb = QPTrim$(DprRec.ItemTag) 'load array
    ThisNum = Val(ReplaceString(DprRec.ItemTag, "-", ""))
    ReDim Preserve TempDprIdx(1 To ValidCnt + 1)
    If ThisNum > BigNum Then
      BigNum = ThisNum 'find largest item tag number
    End If
BadNum:
    frmFAShowPctComp.ShowPctComp x, NumOfRecs
  Next x
  Close DPRHandle
  
  If ValidCnt = 0 Then
    ValidCnt = -1
    Close
    Exit Sub
  End If
  Unload frmFAShowPctComp
  EnableCloseButton frmFADprHistRpt.hwnd, True
  frmFADprHistRpt.cmdExit.Enabled = True
  frmFADprHistRpt.cmdProcess.Enabled = True
  
  frmFAShowPctComp.Show
  frmFAShowPctComp.cmdCancel.Visible = False
  frmFAShowPctComp.Label1 = "Sorting Item Data"
  EnableCloseButton frmFADprHistRpt.hwnd, False
  frmFADprHistRpt.cmdExit.Enabled = False
  frmFADprHistRpt.cmdProcess.Enabled = False
  DoEvents
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  'now sort in tag number numerical order
  Do
    For x = Nextx To ValidCnt
      ThisNum = Val(ReplaceString(TempDprIdx(x).DprNumb, "-", ""))
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempDprIdx(Nextx)
    TempDprIdx(Nextx) = TempDprIdx(ThisX)
    TempDprIdx(ThisX) = HoldThis
    TempDprIdx(Nextx).DprRecNum = TempDprIdx(Nextx).DprRecNum
    If Nextx = ValidCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
    frmFAShowPctComp.ShowPctComp Nextx, ValidCnt
  Loop
  Unload frmFAShowPctComp
  EnableCloseButton frmFADprHistRpt.hwnd, True
  frmFADprHistRpt.cmdExit.Enabled = True
  frmFADprHistRpt.cmdProcess.Enabled = True
  OpenTempDprFile DprIdxHandle
  For x = 1 To ValidCnt
    DprIdx = TempDprIdx(x) 'now finalize the array
    Put DprIdxHandle, x, DprIdx
  Next x
  Close

End Sub

Public Sub CreateTagIdx()
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempNum As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAItemRecLen As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As TagNumbSortIdxType
  
  'this indexing routine sorts all tag numbers from
  'smallest to largest each time a new tag number
  'is saved
  
  FAItemRecLen = Len(FAItemRec)
  FAHandle = FreeFile
  Open "FAITEMS.DAT" For Random Shared As FAHandle Len = FAItemRecLen
  
  NumOfFARecs = LOF(FAHandle) \ Len(FAItemRec)
  
  ReDim TempTagIdx(1 To NumOfFARecs) As TagNumbSortIdxType
  If NumOfFARecs = 1 Then
    Get FAHandle, 1, FAItemRec
    TagIdxRecLen = Len(TagIdx)
    TagIdxHandle = FreeFile
    Open "FATAGIDX.DAT" For Random Shared As TagIdxHandle Len = TagIdxRecLen
    TempTagIdx(1).DataRecNum = 1
    TempTagIdx(1).TagNumb = QPTrim$(FAItemRec.ItemTag)
    Put TagIdxHandle, 1, TempTagIdx(1)
    Close
    Exit Sub
  End If
  
  
  frmFAShowPctComp.Label1 = "Adding and Indexing This Item"
  frmFAShowPctComp.cmdCancel.Visible = False
  frmFAShowPctComp.Show
  BigNum = 0
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    TempTagIdx(x).DataRecNum = x
    TempTagIdx(x).TagNumb = QPTrim$(FAItemRec.ItemTag)
    ThisNum = Val(ReplaceString(FAItemRec.ItemTag, "-", ""))
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close FAHandle
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfFARecs
      ThisNum = Val(ReplaceString(TempTagIdx(x).TagNumb, "-", ""))
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempTagIdx(Nextx)
    TempTagIdx(Nextx) = TempTagIdx(ThisX)
    TempTagIdx(ThisX) = HoldThis
    TempTagIdx(Nextx).DataRecNum = TempTagIdx(Nextx).DataRecNum
'    If Nextx = NumOfFARecs Then Exit Do
    
    If Nextx = NumOfFARecs - 1 Then Exit Do 'Or NumOfFARecs = 1 Then Exit Do
    frmFAShowPctComp.ShowPctComp Nextx, NumOfFARecs
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  Unload frmFAShowPctComp
  TagIdxRecLen = Len(TagIdx)
  TagIdxHandle = FreeFile
  Open "FATAGIDX.DAT" For Random Shared As TagIdxHandle Len = TagIdxRecLen
  For x = 1 To NumOfFARecs
    TagIdx = TempTagIdx(x)
    Put TagIdxHandle, x, TagIdx
  Next x
  
  Close

End Sub

Public Sub CreateAssetIdx()
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempNum As Integer
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim ACItemRecLen As Integer
  Dim NumOfACRecs As Integer
  Dim AssIdx As ACNumbSortIdxType
  Dim AssIdxHandle As Integer
  Dim AssIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As ACNumbSortIdxType
  
  'this indexing routine sorts asset codes in numeric order
  'each time a new asset code is saved
  OpenFACodeNameFile CodeHandle
  
  NumOfACRecs = LOF(CodeHandle) \ Len(CodeRec)
  ReDim TempAssIdx(1 To NumOfACRecs) As ACNumbSortIdxType
  If NumOfACRecs = 1 Then
    OpenAssIdxFile AssIdxHandle
      Get CodeHandle, 1, CodeRec
      TempAssIdx(1).AssNumb = QPTrim$(CodeRec.ASSETCODE)
      TempAssIdx(1).AssRecNum = 1
      Put AssIdxHandle, 1, TempAssIdx(1)
      Close
      Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To NumOfACRecs
    Get CodeHandle, x, CodeRec
    TempAssIdx(x).AssRecNum = x
    TempAssIdx(x).AssNumb = QPTrim$(CodeRec.ASSETCODE)
    ThisNum = CDbl(QPTrim$(CodeRec.ASSETCODE))
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close CodeHandle
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfACRecs
      ThisNum = TempAssIdx(x).AssNumb
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempAssIdx(Nextx)
    TempAssIdx(Nextx) = TempAssIdx(ThisX)
    TempAssIdx(ThisX) = HoldThis
'    If Nextx = NumOfACRecs Then Exit Do
    If Nextx = NumOfACRecs - 1 Then Exit Do 'Or NumOfACRecs = 1 Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  OpenAssIdxFile AssIdxHandle
  For x = 1 To NumOfACRecs
    AssIdx = TempAssIdx(x)
    Put AssIdxHandle, x, AssIdx
  Next x
  
  Close

End Sub

Public Sub CreateDeptIdx()
  Dim BigNum As Integer
  Dim ThisNum As Integer
  Dim ThisX As Integer
  Dim SmallNum As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim DeptHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim DeptItemRecLen As Integer
  Dim NumOfDeptRecs As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DeptIdxHandle As Integer
  Dim DeptIdxRecNum As Integer
  Dim RecNum As Integer
  Dim HoldThis As DeptNumbSortIdxType
  
  'this indexing routine sorts department numbers
  'numerically everytime new dept number is saved
  OpenFADeptCodeFile DeptHandle
  
  NumOfDeptRecs = LOF(DeptHandle) \ Len(DeptRec)
  ReDim TempDeptIdx(1 To NumOfDeptRecs) As DeptNumbSortIdxType
  If NumOfDeptRecs = 1 Then
    OpenDeptIdxFile DeptIdxHandle
      Get DeptHandle, 1, DeptRec
      TempDeptIdx(1).DeptIdxDesc = QPTrim$(DeptRec.DeptDesc)
      TempDeptIdx(1).DeptNumb = DeptRec.DeptNum
      TempDeptIdx(1).DeptRecNum = 1
      Put DeptIdxHandle, 1, TempDeptIdx(1)
      Close
      Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To NumOfDeptRecs
    Get DeptHandle, x, DeptRec
    TempDeptIdx(x).DeptRecNum = x
    TempDeptIdx(x).DeptNumb = DeptRec.DeptNum
    TempDeptIdx(x).DeptIdxDesc = QPTrim$(DeptRec.DeptDesc)
    ThisNum = DeptRec.DeptNum
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close DeptHandle
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfDeptRecs
      ThisNum = TempDeptIdx(x).DeptNumb
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempDeptIdx(Nextx)
    TempDeptIdx(Nextx) = TempDeptIdx(ThisX)
    TempDeptIdx(ThisX) = HoldThis
'    If Nextx = NumOfDeptRecs Then Exit Do
    If Nextx = NumOfDeptRecs - 1 Then Exit Do 'Or NumOfDeptRecs = 1 Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  OpenDeptIdxFile DeptIdxHandle
  For x = 1 To NumOfDeptRecs
    DeptIdx = TempDeptIdx(x)
    Put DeptIdxHandle, x, DeptIdx
  Next x
  
  Close

End Sub

Public Sub CreateFundIdx()
  Dim BigNum As Integer
  Dim ThisNum As Integer
  Dim ThisX As Integer
  Dim SmallNum As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim FundHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim FundItemRecLen As Integer
  Dim NumOfFundRecs As Integer
  Dim FundIdx As FundNumbSortIdxType
  Dim FundIdxHandle As Integer
  Dim FundIdxRecNum As Integer
  Dim RecNum As Integer
  Dim HoldThis As FundNumbSortIdxType
  
  'this indexing routine sorts funds in numeric order
  'each time a new fund number is saved
  OpenFAFundCodeFile FundHandle
  
  NumOfFundRecs = LOF(FundHandle) \ Len(FundRec)
  ReDim TempFundIdx(1 To NumOfFundRecs) As FundNumbSortIdxType
  If NumOfFundRecs = 1 Then
    OpenFundIdxFile FundIdxHandle
      Get FundHandle, 1, FundRec
      TempFundIdx(1).FundIdxDesc = QPTrim$(FundRec.FundDesc)
      TempFundIdx(1).FundNumb = FundRec.FundNum
      TempFundIdx(1).FundRecNum = 1
      Put FundIdxHandle, 1, TempFundIdx(1)
      Close
      Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To NumOfFundRecs
    Get FundHandle, x, FundRec
    TempFundIdx(x).FundRecNum = x
    TempFundIdx(x).FundNumb = FundRec.FundNum
    TempFundIdx(x).FundIdxDesc = FundRec.FundDesc
    ThisNum = FundRec.FundNum
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close FundHandle
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfFundRecs
      ThisNum = TempFundIdx(x).FundNumb
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempFundIdx(Nextx)
    TempFundIdx(Nextx) = TempFundIdx(ThisX)
    TempFundIdx(ThisX) = HoldThis
    If Nextx = NumOfFundRecs - 1 Then Exit Do 'Or NumOfFundRecs = 1 Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  OpenFundIdxFile FundIdxHandle
  For x = 1 To NumOfFundRecs
    FundIdx = TempFundIdx(x)
    Put FundIdxHandle, x, FundIdx
  Next x
  
  Close

End Sub

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
  Open PassP$ For Random Shared As Tempfile ' Len = lentemp '2/14/08
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
  PassTemp.frommdl = 5
  Put Tempfile, 1, PassTemp
  Close
End Sub
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

Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
  
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function

