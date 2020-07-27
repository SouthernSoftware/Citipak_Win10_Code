Attribute VB_Name = "modCommon"
Option Explicit
Public XFiles(1 To 12)

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

'Public Sub OpenPRDraftFile(PRDraftFileHandle As Integer)
'  Dim PRDraftFileRec As DraftInfoFileName
'  Dim PRDraftRecLen As Integer
'  Dim PRDraftFileName As String
'
'  PRDraftFileName = "PRDRAFTI.DAT"
'  PRDraftRecLen = Len(PRDraftFileRec)
'  PRDraftFileHandle = FreeFile
'  Open "PRData\" + PRDraftFileName For Random Shared As PRDraftFileHandle Len = PRDraftRecLen
'End Sub

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
  Dim x As Integer

  x = 1
  CheckForAllFiles = True
  
'  If Not Exist("PRDATA\PRCHECKS.DAT") Then
'    XFiles(x) = "PRDATA\PRCHECKS.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
'  If Not Exist("PRDATA\PRPRNSET.DAT") Then
'    XFiles(x) = "PRDATA\PRPRNSET.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
  If Not Exist("PRDATA\PREMP2.DAT") Then
    XFiles(x) = "PRDATA\PREMP2.DAT"
    x = x + 1
    CheckForAllFiles = False
  End If
'  If Not Exist("PRDATA\PREMP3.DAT") Then
'    XFiles(x) = "PRDATA\PREMP3.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
'  If Not Exist("PRDATA\PREICTBL.DAT") Then
'    XFiles(x) = "PRDATA\PREICTBL.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
  If Not Exist("PRDATA\PRUNIT.DAT") Then
    XFiles(x) = "PRDATA\PRUNIT.DAT"
    x = x + 1
    CheckForAllFiles = False
  End If
'  If Not Exist("PRDATA\PRDEDCOD.DAT") Then
'    XFiles(x) = "PRDATA\PRDEDCOD.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
'  If Not Exist("PRDATA\PRTRANST.DAT") Then
'    XFiles(x) = "PRDATA\PRTRANST.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
'  If Not Exist("PRDATA\PRTRANSH.DAT") Then
'    XFiles(x) = "PRDATA\PRTRANSH.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
'  If Not Exist("PRDATA\PRRETIRE.DAT") Then
'    XFiles(x) = "PRDATA\PRRETIRE.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If
'  If Not Exist("PRDATA\PRPPDEF.DAT") Then
'    XFiles(x) = "PRDATA\PRPPDEF.DAT"
'    x = x + 1
'    CheckForAllFiles = False
'  End If

End Function

'Public Sub TrapForBadDeds()
'  Dim DedRec As DedCodeRecType
'  Dim DHandle As Integer
'  Dim DedRecLen As Integer
'  Dim Emp2Rec As DosEmpData2Type
'  Dim EHandle As Integer
'  Dim Emp2RecLen As Integer
'  Dim x As Integer
'  Dim Y As Integer
'  Dim NumOfRecs As Integer
'  Dim DedDesc$(1 To 12)
''  Dim TotWagePct As Double
''  Dim PayType As String * 1
'
'  DedRecLen = Len(DedRec)
'  DHandle = FreeFile
'  Open "prdata\" + "PRDEDCOD.DAT" For Random Shared As DHandle Len = DedRecLen
'  For x = 1 To 12
'    Get DHandle, x, DedRec
'    DedDesc(x) = DedRec.DCDESC1
'  Next x
'  Close DHandle
'
'  Emp2RecLen = Len(Emp2Rec)
'  EHandle = FreeFile
'  Open "prdata\" + "PREMP2.DAT" For Random Shared As EHandle Len = Emp2RecLen
'
'  NumOfRecs = LOF(EHandle) / Len(Emp2Rec)
''  For Y = 1 To NumOfRecs
''    TotWagePct = 0
''    Get EHandle, Y, Emp2Rec
''    PayType = QPTrim$(Emp2Rec.EMPPTYPE)
''    If PayType = "S" And Emp2Rec.EMPTDATE <= 0 Then '
''      For x = 1 To 8
''        If Emp2Rec.EDist(x).DAmt < 0 Then Emp2Rec.EDist(x).DAmt = 0
''        TotWagePct = Emp2Rec.EDist(x).DAmt + TotWagePct
''      Next x
''
''      If TotWagePct <> 100 Then
''        MsgBox "The wage distribution percentages for employee #" & QPTrim$(Emp2Rec.EmpNo) & " " & QPTrim$(Emp2Rec.EMPFNAME) & " " & QPTrim$(Emp2Rec.EMPLNAME) & " do not equal 100%. Please correct this error before proceeding."
''        Close EHandle
''        End
''      End If
''    End If
''  Next Y
'
'  For Y = 1 To NumOfRecs
'    For x = 1 To 12
'      If Emp2Rec.EmpDed(x).DAmt > 0 And QPTrim$(DedDesc(x)) = "" Then
'        If MsgBox("There is an orphan value on Line" & x & " of " & QPTrim$(Emp2Rec.EMPFNAME) & " " & QPTrim$(Emp2Rec.EMPLNAME) & "'s Miscellaneous Deductions. If you wish to correct this error now then press NO to stop this conversion and return to the DOS program to make the change. If you wish to continue with this conversion then press YES and conversion will delete the orphan value.", vbYesNo) = vbYes Then
'          Emp2Rec.EmpDed(x).DPct = ""
'          Emp2Rec.EmpDed(x).DAmt = 0
'          Emp2Rec.EmpDed(x).DOTI = ""
'          Put EHandle, Y, Emp2Rec
'        Else
'          Close EHandle
'          End
'        End If
'      End If
'    Next x
'  Next Y
'  Close EHandle
'End Sub

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
  Unload frmConvert
  Unload frmInProg
  Unload frmWarnMissingFiles
End Sub
