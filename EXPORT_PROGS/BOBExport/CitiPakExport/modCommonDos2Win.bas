Attribute VB_Name = "modCommonDos2Win"
Option Explicit
  Public XFiles(1 To 12)
  Public OutFileNames(1 To 20) As String
  Public InFileNames(1 To 20) As String
  Public Const PRData = "prdata\"
  Public Const EarnNoMatchName = "PRDATA\PRERNOMC.DAT"
  Public Const TransHistFileName = "PRTRANSH.DAT"
  Public Const ErnCodeFileName = "PRERNCOD.DAT"
  Public Const EmpData2Name = "PREMP2.DAT"
  Public Const UnitFileName = "PRUNIT.DAT"
  Public Const LeaveFileName = "PRLEAVE.DAT"
Public Sub OpenTransHistFile(TransHistFileHandle As Integer)
  Dim TransHistFileRec As TransRecType
  Dim TransHistRecLen As Integer
  TransHistRecLen = Len(TransHistFileRec)
  TransHistFileHandle = FreeFile
  Open PRData + TransHistFileName For Random Shared As TransHistFileHandle Len = TransHistRecLen
End Sub
Public Sub OpenUnitFileFall04(FileHandleFall04 As Integer)
  Dim UnitFileRecFall04 As UnitFileRecTypeFall04
  Dim UnitRecLenFall04 As Integer
  UnitRecLenFall04 = Len(UnitFileRecFall04)
  FileHandleFall04 = FreeFile
  Open PRData + UnitFileName For Random Shared As FileHandleFall04 Len = UnitRecLenFall04
End Sub

Public Sub OpenEmpData2FileFall04(EmpData2FileHandleFall04 As Integer)
  Dim EmpData2FileRecFall04 As EmpData2TypeFall04
  Dim EmpData2RecLenFall04 As Integer
  EmpData2RecLenFall04 = Len(EmpData2FileRecFall04)
  EmpData2FileHandleFall04 = FreeFile
  Open PRData + EmpData2Name For Random Shared As EmpData2FileHandleFall04 Len = EmpData2RecLenFall04
End Sub

Public Sub OpenDosErnCodeFile(DosErnCodeFileHandle As Integer)
  Dim DosErnCodeFileRec As DosErnCodeRecType
  Dim DosErnCodeRecLen As Integer
  DosErnCodeRecLen = Len(DosErnCodeFileRec)
  DosErnCodeFileHandle = FreeFile
  Open PRData + ErnCodeFileName For Random Shared As DosErnCodeFileHandle Len = DosErnCodeRecLen
End Sub
Public Sub OpenErnCodeFile(ErnCodeFileHandle As Integer)
  Dim ErnCodeFileRec As ErnCodeRecType
  Dim ErnCodeRecLen As Integer
  ErnCodeRecLen = Len(ErnCodeFileRec)
  ErnCodeFileHandle = FreeFile
  Open PRData + ErnCodeFileName For Random Shared As ErnCodeFileHandle Len = ErnCodeRecLen
End Sub
Public Sub OpenEarnNoMatchFile(EarnNoMatchHandle As Integer) '12/12/02
  Dim EarnNoMatchLen As Integer
  Dim EarnNoMatch As EarnNoMatchType
  EarnNoMatchLen = Len(EarnNoMatch)
  EarnNoMatchHandle = FreeFile
  Open EarnNoMatchName For Random Shared As EarnNoMatchHandle Len = EarnNoMatchLen
End Sub
Public Sub OpenOldLeaveFileName(LeaveHandle As Integer)
  Dim LeaveRec As OldLeaveRecType
  Dim LeaveRecLen As Integer
  LeaveRecLen = Len(LeaveRec)
  LeaveHandle = FreeFile
  Open "PRData\" + LeaveFileName For Random Shared As LeaveHandle Len = LeaveRecLen
End Sub
Public Sub OpenLeaveFileName(LeaveHandle As Integer)
  Dim LeaveRec As LeaveRecType
  Dim LeaveRecLen As Integer
  LeaveRecLen = Len(LeaveRec)
  LeaveHandle = FreeFile
  Open "PRData\" + LeaveFileName For Random Shared As LeaveHandle Len = LeaveRecLen
End Sub

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

Public Sub OpenPRDraftFile(PRDraftFileHandle As Integer)
  Dim PRDraftFileRec As DraftInfoFileName
  Dim PRDraftRecLen As Integer
  Dim PRDraftFileName As String

  PRDraftFileName = "PRDRAFTI.DAT"
  PRDraftRecLen = Len(PRDraftFileRec)
  PRDraftFileHandle = FreeFile
  Open "PRData\" + PRDraftFileName For Random Shared As PRDraftFileHandle Len = PRDraftRecLen
End Sub

Public Sub KillFile(FileName As String)
  If Exist(FileName$) Then
    Kill FileName$
  End If
End Sub

Public Function OldRound#(n As Double)
  If n < -2000000000# Then n = 0
  OldRound# = Round(n, 2)
End Function

Public Function CheckForAllFiles() As Boolean
  Dim X As Integer

  X = 1
  CheckForAllFiles = True
  
  If Not Exist("PRDATA\PRCHECKS.DAT") Then
    XFiles(X) = "PRDATA\PRCHECKS.DAT     CREATED WHEN CHECKS ARE PRINTED."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRPRNSET.DAT") Then
    XFiles(X) = "PRDATA\PRPRNSET.DAT     CREATED WHEN PRINTER DEFAULTS ARE SET."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PREMP2.DAT") Then
    XFiles(X) = "PRDATA\PREMP2.DAT     CREATED WHEN AN EMPLOYEE DATA IS SAVED."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PREMP3.DAT") Then
    XFiles(X) = "PRDATA\PREMP3.DAT     CREATED WHEN EMPLOYEE DATA IS SAVED."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PREICTBL.DAT") Then
    XFiles(X) = "PRDATA\PREICTBL.DAT     CREATED WHEN EIC MAINTENANCE DATA IS SAVED."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRUNIT.DAT") Then
    XFiles(X) = "PRDATA\PRUNIT.DAT     CREATED WHEN EMPLOYER INFORMATION IS SAVED."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRDEDCOD.DAT") Then
    XFiles(X) = "PRDATA\PRDEDCOD.DAT     CREATED WHEN DEDUCTION MAINTENANCE DATA IS SAVED."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRTRANST.DAT") Then
    XFiles(X) = "PRDATA\PRTRANST.DAT     CREATED WHEN A PAYROLL IS RUN."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRTRANSH.DAT") Then
    XFiles(X) = "PRDATA\PRTRANSH.DAT     CREATED WHEN A PAYROLL IS RUN."
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRRETIRE.DAT") Then
    XFiles(X) = "PRDATA\PRRETIRE.DAT     CREATED WHEN RETIREMENT MAINTENANCE DATA IS SAVED"
    X = X + 1
    CheckForAllFiles = False
  End If
  If Not Exist("PRDATA\PRPPDEF.DAT") Then
    XFiles(X) = "PRDATA\PRPPDEF.DAT     CREATED WHEN PAYROLL DEFAULTS ARE SAVED."
    X = X + 1
    CheckForAllFiles = False
  End If

End Function

Public Sub TrapForBadDeds()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim DedRecLen As Integer
  Dim Emp2Rec As DosEmpData2Type
  Dim EHandle As Integer
  Dim Emp2RecLen As Integer
  Dim X As Integer
  Dim y As Integer
  Dim NumOfRecs As Integer
  Dim DedDesc$(1 To 12)
'  Dim TotWagePct As Double
'  Dim PayType As String * 1

  DedRecLen = Len(DedRec)
  DHandle = FreeFile
  Open "prdata\" + "PRDEDCOD.DAT" For Random Shared As DHandle Len = DedRecLen
  For X = 1 To 12
    Get DHandle, X, DedRec
    DedDesc(X) = DedRec.DCDESC1
  Next X
  Close DHandle

  Emp2RecLen = Len(Emp2Rec)
  EHandle = FreeFile
  Open "prdata\" + "PREMP2.DAT" For Random Shared As EHandle Len = Emp2RecLen

  NumOfRecs = LOF(EHandle) / Len(Emp2Rec)
'  For Y = 1 To NumOfRecs
'    TotWagePct = 0
'    Get EHandle, Y, Emp2Rec
'    PayType = QPTrim$(Emp2Rec.EMPPTYPE)
'    If PayType = "S" And Emp2Rec.EMPTDATE <= 0 Then '
'      For x = 1 To 8
'        If Emp2Rec.EDist(x).DAmt < 0 Then Emp2Rec.EDist(x).DAmt = 0
'        TotWagePct = Emp2Rec.EDist(x).DAmt + TotWagePct
'      Next x
'
'      If TotWagePct <> 100 Then
'        MsgBox "The wage distribution percentages for employee #" & QPTrim$(Emp2Rec.EmpNo) & " " & QPTrim$(Emp2Rec.EMPFNAME) & " " & QPTrim$(Emp2Rec.EMPLNAME) & " do not equal 100%. Please correct this error before proceeding."
'        Close EHandle
'        End
'      End If
'    End If
'  Next Y

  For y = 1 To NumOfRecs
    For X = 1 To 12
      If Emp2Rec.EmpDed(X).DAmt > 0 And QPTrim$(DedDesc(X)) = "" Then
        If MsgBox("There is an orphan value on Line" & X & " of " & QPTrim$(Emp2Rec.EMPFNAME) & " " & QPTrim$(Emp2Rec.EMPLNAME) & "'s Miscellaneous Deductions. If you wish to correct this error now then press NO to stop this conversion and return to the DOS program to make the change. If you wish to continue with this conversion then press YES and conversion will delete the orphan value.", vbYesNo) = vbYes Then
          Emp2Rec.EmpDed(X).DPct = ""
          Emp2Rec.EmpDed(X).DAmt = 0
          Emp2Rec.EmpDed(X).DOTI = ""
          Put EHandle, y, Emp2Rec
        Else
          Close EHandle
          End
        End If
      End If
    Next X
  Next y
  Close EHandle
End Sub

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

Public Sub UnloadAllFormsAndOpn()
  Unload frmConvertDos2Win
  Unload frmDos2WinInProg
  Unload frmDos2WinFilesMissing
  Unload frmWarnConvertFilesMissing
End Sub

Public Function FilesROK(frm As Form, InFileNames() As String, OutFileNames() As String, ThisMany As Integer) As Boolean
  Dim NextName As Integer
  Dim X As Integer
  'this function scans for files necessary to run a particular part
  'of the program and looks in the PRData folder for them...if they
  'are missing then a warning screen pops up telling the user what
  'the problem is and how to fix it (located in frmWarnFilesMissing)
  FilesROK = True
  NextName = 1
  For X = 1 To ThisMany 'for loop takes incoming files needing checking
  'and looks in PRData for them...if they are missing they are added
  'to OutFileNames and if they are OK then they are skipped
    If Not Exist(InFileNames(X)) Then
      OutFileNames(NextName) = InFileNames(X)
      NextName = NextName + 1
      FilesROK = False
    End If
  Next X
  If FilesROK = False Then
    frmWarnConvertFilesMissing.Show vbModal, frm
    For X = 1 To ThisMany
      InFileNames(X) = ""
      OutFileNames(X) = ""
    Next X
  End If
End Function


Public Sub InsertVoidFieldsIntoTrans()
  Dim TransRec As TransRecType
  Dim NumOfTransRecs As Long
  Dim X As Long
  Dim THandle As Integer
  Dim ThisDate As Integer
  Dim CheckNum As String
  Dim RegHrsPaid As Double
  Dim EmpPin As Integer
  Dim y As Integer
  
  OpenTransHistFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  If NumOfTransRecs = 0 Then
    Close THandle
    Exit Sub
  End If
  
  For X = 1 To NumOfTransRecs
    Get THandle, X, TransRec
'      If TransRec.CheckNum = -105376 Or TransRec.CheckNum = 105376 Then Stop
      If TransRec.GrossPay < 0 And TransRec.CheckNum <= 0 Then
         CheckNum = -TransRec.CheckNum
         ThisDate = TransRec.CheckDate
         RegHrsPaid = TransRec.RegHrsPaid
         EmpPin = TransRec.EmpPin
         For y = 1 To X
           Get THandle, y, TransRec
             If TransRec.CheckNum = CheckNum Then
               If ThisDate = TransRec.CheckDate And RegHrsPaid = -TransRec.RegHrsPaid And EmpPin = TransRec.EmpPin Then
                 TransRec.Voided = "Y"
                 TransRec.VoidRec = X
                 Put THandle, y, TransRec
                 Get THandle, X, TransRec
                 TransRec.Voided = "Y"
                 TransRec.VoidRec = y
                 Put THandle, X, TransRec
                 Exit For
               End If
             End If
         Next y
      End If
  Next X
  Close THandle
End Sub


