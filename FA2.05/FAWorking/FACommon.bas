Attribute VB_Name = "FACommon"
Option Explicit
  Public RecNum As Long
  Public CodeNum As Long
  Public ScreenW As Long
  Public coladj As Double
  Public doAlign As Boolean
  Public alnRpt$
  Public BadMaskFlag As Boolean
  Public NumOfAligns As Integer
  Public OutFileNames(1 To 20) As String
  Public InFileNames(1 To 20) As String
  Public ComputerName As String
  Public CurrCitiPath As String
  Public StartPath As String
  Public RptOpt As Integer 'used to determine the type of reports; graphic or text
  Public ToPrint1(1 To 10) As Integer
  Public ToPrint2(1 To 10) As Integer
  Public DeptList() As String
  Public NumOfDepts As Integer
  
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
            Public Const PRData = "prdata\"
            Public Const FASetUpFileName = "FASETUP.DAT"
            Public Const FAItemFileName = "FAITEMS.DAT"
            Public Const FAAssetCodeName = "FACODES.DAT"
      Public Const UnitFileName = "PRUNIT.DAT"
       Public Const SysFileName = "PRSYS.DAT"
 Public Const TransWorkFileName = "PRTRANST.DAT"
 Public Const TransHistFileName = "PRTRANSH.DAT"
  Public Const PrinterSetUpFile = "PRPRNSET.DAT"
    Public Const GLAcctIdxFile = "BAACCTDX.DAT"
    Public Const JGLAcctIdxFile = "GLACCT.IDX"
      Public Const AcctFileName = "GLACCT.DAT"
      Public Const TransFileName = "GLTRANS.DAT"
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
Public Sub OpenPrinterSetupFile(PrinterSUFHandle As Integer)
  Dim PrinterSUFRec As PRNSetupRecType
  Dim PrinterSUFRecLen As Integer
  PrinterSUFRecLen = Len(PrinterSUFRec)
  PrinterSUFHandle = FreeFile
  Open PRData + PrinterSetUpFile For Random Shared As PrinterSUFHandle Len = PrinterSUFRecLen
End Sub

Public Sub OpenGLTransFile(GLTransHandle As Integer)
  Dim GLTransRec As GLTransRecType
  Dim GLTransRecLen As Integer
  GLTransRecLen = Len(GLTransRec)
  GLTransHandle = FreeFile
  Open GetCitiDirFolder + TransFileName For Random Shared As GLTransHandle Len = GLTransRecLen
End Sub

Public Sub OpenGLAcctFile(GLHandle As Integer)
  Dim GLRec As GLAcctRecType
  Dim GLRecLen As Integer
  GLRecLen = Len(GLRec)
  GLHandle = FreeFile
  Open GetCitiDirFolder + AcctFileName For Random Shared As GLHandle Len = GLRecLen
End Sub
Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function

Public Sub OpenSysFile(SysFileHandle As Integer)
  Dim SysFileRec As RegDSysFileRecType
  Dim SysRecLen As Integer
  SysRecLen = Len(SysFileRec)
  SysFileHandle = FreeFile
  Open PRData + SysFileName For Random Shared As SysFileHandle Len = SysRecLen
End Sub

Public Sub OpenUnitFile(FileHandle As Integer)
  Dim UnitFileRec As UnitFileRecType
  Dim UnitRecLen As Integer
  UnitRecLen = Len(UnitFileRec)
  FileHandle = FreeFile
  Open PRData + UnitFileName For Random Shared As FileHandle Len = UnitRecLen
End Sub

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim Cnt As Long
  Dim ThisChar As Integer
  StrLen = Len(Text)
  For Cnt = 1 To StrLen
    ThisChar = Asc(Mid$(Text, Cnt, 1))
    If ThisChar = 0 Then
      Mid$(Text$, Cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
  End Function

Public Function ReplaceString$(Text As String, ChangeThis As String, ToThis As String)
  Dim StrLen As Long
  Dim Cnt As Long
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
  
  For Cnt = 1 To StrLen 'set up loop to iterate thru entire text
    ThisChar = Mid$(Text, Cnt, 1) 'step thru text a letter at a time
    CTChar = Mid$(Text, Cnt, CTLen) 'starting with the current letter
    'read ahead the length of the text "change this"
    If CTChar = ChangeThis Then 'if we find the "change this" in the
    'text
      NewText = NewText + ToThis 'assign the length of CTChar to "ToThis"
      'inside the rebuilt new text
      Cnt = Cnt + BigLen - 1 'advance count to compensate for the addition of
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

'Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
'   frmLoadingRpt.Show
'   frmViewPrint.ReportName = ReportFile$
'   frmViewPrint.Caption = Title
'   frmViewPrint.PgNum = PgNum
'   frmViewPrint.cmdAlignment.Visible = False
'   If ForceSBar Then
'     frmViewPrint.fpMemo1.ScrollBars = BothFixed
'   Else
'     frmViewPrint.fpMemo1.ScrollBars = BothAuto
'   End If
'   If Algn Then
'     frmViewPrint.cmdAlignment.Enabled = True
'     frmViewPrint.AlignRpt = AlgnRptfile$
'    Else
'      frmViewPrint.cmdAlignment.Enabled = False
'    End If
'   frmViewPrint.Show 1
'   Unload frmLoadingRpt
'   doAlign = False
'End Sub

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
  Dim Cnt As Integer
  Dim DecCnt As Integer
  Dim StrLen As Long
  Dim ThisChar$
  'this function traps errors created when a user keys in a
  'decimal value and inadvertantly keys in more than 1 decimal
  StrLen = Len(Text)
  For Cnt = 1 To StrLen
    ThisChar = Mid$(Text, Cnt, 1)
    If ThisChar = "." Then DecCnt = DecCnt + 1 'counts decimals
  Next Cnt
  If DecCnt > 1 Then 'if decimal count is more than 1 then process
  'accordingly
    CheckFor2ManyDecimals = True
  Else
    CheckFor2ManyDecimals = False
  End If
End Function

Function GetCitiDirFolder()
  ReDim SysRec(1) As RegDSysFileRecType
  Dim SysFileHandle As Integer
  Dim TmpChr As String
  Dim TmpDir As String
  
  On Local Error Resume Next
  OpenSysFile SysFileHandle
  Get SysFileHandle, 1, SysRec(1)
  Close SysFileHandle
  'In the System Interface screen you cannot access the needed
  'GL list if nothing is saved...this is an effort at allowing a gl
  'search to occur if there is at least an entry in the Citipak
  'field
  'this function is also used to access gl files that are located
  'only in the Citipak directory
  TmpDir = QPTrim$(SysRec(1).CITIDIR)
  If Len(TmpDir) = 0 Then
    GoTo PathOK
  End If
  
  TmpChr = Right$(TmpDir, 1)
  If TmpChr = ":" Then
    GetCitiDirFolder = TmpDir
    GoTo PathOK
  ElseIf TmpChr <> "\" Then
    GetCitiDirFolder = TmpDir + "\"
    GoTo PathOK
  Else
    GetCitiDirFolder = TmpDir
  End If

PathOK:
  
End Function

Public Function FilesROK(frm As Form, InFileNames() As String, OutFileNames() As String, ThisMany As Integer) As Boolean
  Dim NextName As Integer
  Dim x As Integer
  'this function scans for files necessary to run a particular part
  'of the program and looks in the PRData folder for them...if they
  'are missing then a warning screen pops up telling the user what
  'the problem is and how to fix it (located in frmWarnFilesMissing)
  FilesROK = True
  NextName = 1
  For x = 1 To ThisMany 'for loop takes incoming files needing checking
  'and looks in PRData for them...if they are missing they are added
  'to OutFileNames and if they are OK then they are skipped
    If Not Exist(InFileNames(x)) Then
      OutFileNames(NextName) = InFileNames(x)
      NextName = NextName + 1
      FilesROK = False
    End If
  Next x
  If FilesROK = False Then
    frmFAWarnFilesMissing.Show vbModal, frm
    For x = 1 To ThisMany
      InFileNames(x) = ""
      OutFileNames(x) = ""
    Next x
  End If
End Function

Public Sub UnloadAllFormsAndOpn()
  Unload frmFAAssetsCodesmenu
  Unload frmFAEditAssetCode
  Unload frmFAEditItem
  Unload frmFAItemCheckList
  Unload frmFAItemLookUp
  Unload frmFAMainMenu
  Unload frmFAMasterItemListing
  Unload frmFAReportMenu
  Unload frmFAWarnFilesMissing
  Unload frmFAYearEndMenu
  Unload frmFASystemSetup
End Sub
Sub RPTSetupPRN(RPTNum, Handle)
  Dim RPTPitch As Integer
  Dim PrinterSetUpFile As Integer
  Dim PrntType As PRNSetupRecType
  Dim x As Integer
  Dim PHandle As Integer
  Dim DefPrinter As String
  Dim PrnDef As String
  Dim LineLen As Integer
  Dim TextLine$
  Dim y As Integer
  Dim z As Integer
  Dim NextCommaPOS As Integer
  Dim CodeStartPOS As Integer
  Dim Codeline1$
  Dim Codeline2$
  'this sub coordinates the printing procedure so that any
  'pitch data saved in the Printer setup screen for a
  'particular report gets sent to the printer
  For z = 1 To 10 'clear all existing codes
    ToPrint1(z) = 0
    ToPrint2(z) = 0
  Next z
  OpenPrinterSetupFile PrinterSetUpFile
  Get PrinterSetUpFile, 1, PrntType
  Close PrinterSetUpFile
  DefPrinter = QPTrim$(PrntType.Printer)
  'if a pitch isn't saved for this print job then by
  'default the pitch becomes 10
  
  If RPTNum = 123 Then GoTo SkipThis '123 is an arbitrary
  'number used to signify the end of a report that tells
  'this program to look for the reset codes
  
  RPTPitch = PrntType.RPT(RPTNum) 'pitch is specified in the
  'printer setup screen

SkipThis:

  GoSub GetPrinterCodes
  If Len(Codeline1) Then 'CodeLine1 represents the reset codes
  'because in the prprndf.dat file the reset codes come
  'before the pitch codes
  'at this point the proper codes have been determined and
  'the select statement tells the printer which codes to use
    Select Case y
      Case 1:
        Print #Handle, Chr(ToPrint1(1));
      Case 2:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2));
      Case 3:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3));
      Case 4:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4));
      Case 5:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5));
      Case 6:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6));
      Case 7:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7));
      Case 8:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8));
      Case 9:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9));
      Case 10:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9)); Chr(ToPrint1(10));
      Case Else:
    End Select
  ElseIf Len(Codeline2) Then 'CodeLine2 represents the pitch codes
    Select Case y 'Y = the number of codes
      Case 1:
        Print #Handle, Chr(ToPrint2(1));
      Case 2:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2));
      Case 3:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3));
      Case 4:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4));
      Case 5:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5));
      Case 6:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6));
      Case 7:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7));
      Case 8:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8));
      Case 9:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9));
      Case 10:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9)); Chr(ToPrint2(10));
      Case Else:
    End Select
  End If
  
  Exit Sub
  
GetPrinterCodes:
  PHandle = FreeFile
  Open "PRData\Prprndf.dat" For Input As #PHandle  ' Open file.
  Line Input #PHandle, TextLine   ' Read first line into TextLine.
   'the second line is where individual printers start their codes
   NextCommaPOS = 1
   
   Do While Not EOF(PHandle) And NextCommaPOS <> 0  ' Loop until end of file.
     Line Input #PHandle, TextLine   ' Read next line into Textline.
     If TextLine = "@" + DefPrinter$ Then 'locate the default printer
     
         If EOF(PHandle) Then Exit Do 'if for some reason we get to the end of the file
         'then exit
         If RPTNum = 123 Then '123 tells this code that we want the
         'reset codes
           Line Input #PHandle, TextLine 'read next line which by convention
           'will always be the reset code line
             LineLen = Len(TextLine)
             Codeline1 = Mid(TextLine, 11, LineLen) 'by convention
             '11 is where the first reset code begins in this line
             CodeStartPOS = 1
             y = 1
             Do
               NextCommaPOS = InStr(CodeStartPOS, Codeline1, ",") 'look for comma
               If NextCommaPOS = 0 Then 'if comma pos = 0 then we have no more commas
                 LineLen = Len(Codeline1)
                 ToPrint1(y) = CInt(Mid(Codeline1, CodeStartPOS, 3))
                 Exit Do 'we're at the end so exit loop
               End If
               ToPrint1(y) = CInt(Mid(Codeline1, CodeStartPOS, 3)) 'look for a comma
               CodeStartPOS = NextCommaPOS + 1 'start just behind the last comma
               y = y + 1 ' advance y until no more commas found
             Loop Until NextCommaPOS = 0
             GoTo XIsOne 'jump to outer loop
         End If
       Do
         Line Input #PHandle, TextLine 'look for pitch codes
         If Mid(TextLine, 1, 2) = RPTPitch Then 'we found the proper pitch
           LineLen = Len(TextLine)
           Codeline2 = Mid(TextLine, 11, LineLen) 'read this line into
           'Codeline2
           y = 1
           CodeStartPOS = 1
           Do
             NextCommaPOS = InStr(CodeStartPOS, Codeline2, ",")
             If NextCommaPOS = 0 Then 'no more commas in line
               LineLen = Len(Codeline2)
               ToPrint2(y) = CInt(Mid(Codeline2, CodeStartPOS, 3))
               Exit Do
             End If
             ToPrint2(y) = CInt(Mid(Codeline2, CodeStartPOS, 3)) 'keep looking
             'for commas
             CodeStartPOS = NextCommaPOS + 1
             y = y + 1
           Loop Until NextCommaPOS = 0
           
           Exit Do
         End If
        Loop Until NextCommaPOS = 0
XIsOne:
     End If 'ends if TextLine = @ + DefPrinter
   Loop
   Close #PHandle   ' Close file.
   Return
   
   
End Sub

Public Function AddDashesToGLNumber(ByVal GLNum$, Fund As Integer, Dept As Integer, Detail As Integer)
  
  Dim NewGLNum As String
  
  If Mid(GLNum, Fund + 1, 1) <> "-" And Mid(GLNum, Fund + Dept + 2, 1) <> "-" Then
      NewGLNum = Mid(GLNum, 1, Fund) + "-" + Mid(GLNum, Fund + 1, Dept) + "-" + Mid(GLNum, Fund + Dept + 1, Detail)
      AddDashesToGLNumber = NewGLNum
  Else
      AddDashesToGLNumber = GLNum
  End If
  

End Function

Public Sub GetAcctStruct(GLFundLen%, GLAcctLen%, GLDetLen%)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetupRec(1) As GLSetupRecType
  'this sub determines the lengths of each piece of the gl number...
  '(ie. 12-345-6789 breaks down to GLFundLen = 2, GLAcctLen = 3
  'and GLDetLen (Dept) = 4)...this data is used in validating
  'GL numbers before they are saved
'  StartPath = StartPath
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
  Erase GLSetupRec
End Sub

Public Function PromptBadGLNum(frm As Form) As BadGLNUMOption
  frmFABadGLNum.Show vbModal, frm
  PromptBadGLNum = frmFABadGLNum.Selection
  Unload frmFABadGLNum
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
  Unload frmFAChangedWarning
End Function

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
'   frmLoadingRpt.Show
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
'   Unload frmLoadingRpt
   doAlign = False
End Sub

Public Sub SortAssetCodes(Arr() As Struct, NumOfFARecs As Integer, ByRef PCnt As Integer, ByRef PNumOFFARecs As Integer, PBar As Boolean)
  
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempArr As Struct
  Dim OddRecNums As Integer
  
  On Error GoTo ERRORSTUFF
  
  ReDim OddNums(1 To NumOfFARecs) As Struct
  BigNum = 0
  OddRecNums = NumOfFARecs
  For x = 1 To NumOfFARecs
  PCnt = PCnt + 1
    If Len(QPTrim$(Arr(x).who)) <> 4 Then GoTo BadNum
    ThisNum = CDbl(QPTrim$(Arr(x).who))
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
    If PBar = True Then frmFAShowPctComp.ShowPctComp PCnt, PNumOFFARecs
  Next x

  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  Do
    For x = Nextx To OddRecNums
      If Len(QPTrim$(Arr(x).who)) <> 4 Then
        TempArr = Arr(OddRecNums)
        Arr(OddRecNums) = Arr(x)
        Arr(x) = TempArr
        OddRecNums = OddRecNums - 1
        GoTo BadNum1
      End If
      ThisNum = CDbl(QPTrim$(Arr(x).who))
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    TempArr = Arr(Nextx)
    Arr(Nextx) = Arr(ThisX)
    Arr(ThisX) = TempArr
BadNum1:
    Nextx = Nextx + 1
    PCnt = PCnt + 1
    If Nextx = NumOfFARecs Then Exit Do
    SmallNum = BigNum
    If PBar = True Then frmFAShowPctComp.ShowPctComp PCnt, PNumOFFARecs
  Loop
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "SortAssetCodes", "modFACommon", Erl)
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
End Sub

Public Sub SortTagNums(Arr() As Struct, NumOfFARecs As Integer, ByRef PCnt As Integer, ByRef PNumOFFARecs As Integer)
  
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempArr As Struct
  Dim OddRecNums As Integer
  
  On Error GoTo ERRORSTUFF
  
  ReDim OddNums(1 To NumOfFARecs) As Struct
  BigNum = 0
  OddRecNums = NumOfFARecs
  For x = 1 To NumOfFARecs
    PCnt = PCnt + 1
    ThisNum = CDbl(QPTrim$(Arr(x).who))
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
    frmFAShowPctComp.ShowPctComp PCnt, PNumOFFARecs
  Next x

  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  Do
    For x = Nextx To OddRecNums
      ThisNum = CDbl(QPTrim$(Arr(x).who))
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    TempArr = Arr(Nextx)
    Arr(Nextx) = Arr(ThisX)
    Arr(ThisX) = TempArr
BadNum1:
    Nextx = Nextx + 1
    PCnt = PCnt + 1
    If Nextx = NumOfFARecs Then Exit Do
    SmallNum = BigNum
    frmFAShowPctComp.ShowPctComp PCnt, PNumOFFARecs
  Loop

  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "SortTagNums", "modFACommon", Erl)
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
End Sub

Public Sub ExtractDepts()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim ThisDept As String
  Dim y As Integer
  
  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    If Len(QPTrim$(FAItemRec.IDEPT)) = 4 Then
      ThisDept = QPTrim$(FAItemRec.IDEPT)
      Exit For
    End If
  Next x
  ReDim DeptList(1 To NumOfFARecs) As String
  
  Nextx = 1
  DeptList(Nextx) = ThisDept
  NumOfDepts = NumOfDepts + 1
  y = 1
  Do
    For x = 1 To NumOfFARecs
      Get FAHandle, x, FAItemRec
        If QPTrim$(FAItemRec.IDEPT) = "" Then GoTo EmptyString
'        If Len(QPTrim$(FAItemRec.IDEPT)) <> 4 Then Stop
        For y = 1 To NumOfFARecs
          If QPTrim$(FAItemRec.IDEPT) = DeptList(y) Then
            Exit For
          End If
        Next y
        If y = NumOfFARecs + 1 Then
          NumOfDepts = NumOfDepts + 1
          DeptList(NumOfDepts) = QPTrim$(FAItemRec.IDEPT)
        End If
      If x = NumOfFARecs Then Exit Do
EmptyString:
    Next x
  Loop
  Close FAHandle
  ReDim Preserve DeptList(1 To NumOfDepts) As String
  Call SortDeptCodes(DeptList(), NumOfDepts)
'  For x = 1 To NumOfDepts
'    Debug.Print DeptList(x)
'  Next x
End Sub

Public Sub SortDeptCodes(DeptList() As String, NumOfFARecs As Integer)
  
  Dim BigNum As Double
  Dim ThisNum As Double
  Dim ThisX As Integer
  Dim SmallNum As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempDeptList As String
  Dim OddRecNums As Integer
  
  On Error GoTo ERRORSTUFF
  
  BigNum = 0
  OddRecNums = NumOfFARecs
  For x = 1 To NumOfFARecs
    If Len(QPTrim$(DeptList(x))) <> 4 Then GoTo BadNum
    ThisNum = CDbl(QPTrim$(DeptList(x)))
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
BadNum:
  Next x

  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  Do
    For x = Nextx To OddRecNums
      If Len(QPTrim$(DeptList(x))) <> 4 Then
        TempDeptList = DeptList(OddRecNums)
        DeptList(OddRecNums) = DeptList(x)
        DeptList(x) = TempDeptList
        OddRecNums = OddRecNums - 1
        GoTo BadNum1
      End If
      ThisNum = CDbl(QPTrim$(DeptList(x)))
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    TempDeptList = DeptList(Nextx)
    DeptList(Nextx) = DeptList(ThisX)
    DeptList(ThisX) = TempDeptList
    Nextx = Nextx + 1
BadNum1:
    If Nextx = OddRecNums Then Exit Do
    SmallNum = BigNum
  Loop
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "SortDeptCodes", "modFACommon", Erl)
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
End Sub

