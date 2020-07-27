Attribute VB_Name = "modFAConvertCommon"
Option Explicit
  Public ScreenW As Long
  Public coladj As Double
  Public OutFileNames(1 To 20) As String
  Public InFileNames(1 To 20) As String
  Public DeptList() As String
  Public NumOfDepts As Integer
  Public ThisVersion As Integer
  Public NotNumber() As String
  Public NNDesc() As String
  Public NNRecNum() As Integer
  Public NumOfBad As Integer
  
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
  Public Const PRData = "prdata\"
  Public Const FAItemFileName = "FAITEMS.DAT"
  Public Const FADeptCodeName = "FADEPTCD.DAT"
  Public Const FASetUpFileName = "FASETUP.DAT"
  Public Const AssIdxName = "FAASSIDX.DAT"
  Public Const DeptIdxName = "FADEPIDX.DAT"
  Public Const FAAssetCodeName = "FACODES.DAT"
Public Sub OpenFAItemFileV1(FAItemHandleV1 As Integer)
  Dim FAItemRecV1 As DosFAItemRecTypeV1
  Dim FAItemRecLenV1 As Integer
  FAItemRecLenV1 = Len(FAItemRecV1)
  FAItemHandleV1 = FreeFile
  Open FAItemFileName For Random Shared As FAItemHandleV1 Len = FAItemRecLenV1
End Sub
Public Sub OpenAssIdxFile(AssIdxHandle As Integer)
  Dim AssIdxRec As ACNumbSortIdxType
  Dim AssIdxLen As Integer
  AssIdxLen = Len(AssIdxRec)
  AssIdxHandle = FreeFile
  Open AssIdxName For Random Shared As AssIdxHandle Len = AssIdxLen
End Sub
Public Sub OpenFACodeNameFile(FACodeNameHandle As Integer)
  Dim FACodeNameRec As FAAssetCodeRecType
  Dim FACodeNameRecLen As Integer
  FACodeNameRecLen = Len(FACodeNameRec)
  FACodeNameHandle = FreeFile
  Open FAAssetCodeName For Random Shared As FACodeNameHandle Len = FACodeNameRecLen
End Sub
Public Sub OpenFADeptCodeFile(FADeptCodeHandle As Integer)
  Dim FADeptCodeRec As FADeptCodeType
  Dim FADeptCodeRecLen As Integer
  FADeptCodeRecLen = Len(FADeptCodeRec)
  FADeptCodeHandle = FreeFile
  Open FADeptCodeName For Random Shared As FADeptCodeHandle Len = FADeptCodeRecLen
End Sub
Public Sub OpenFAItemFile(FAItemHandle As Integer)
  Dim FAItemRec As DosFAItemRecType
  Dim FAItemRecLen As Integer
  FAItemRecLen = Len(FAItemRec)
  FAItemHandle = FreeFile
  Open FAItemFileName For Random Shared As FAItemHandle Len = FAItemRecLen
End Sub
Public Sub OpenDeptIdxFile(DeptIdxHandle As Integer)
  Dim DeptIdxRec As DeptNumbSortIdxType
  Dim DeptIdxLen As Integer
  DeptIdxLen = Len(DeptIdxRec)
  DeptIdxHandle = FreeFile
  Open DeptIdxName For Random Shared As DeptIdxHandle Len = DeptIdxLen
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
Public Sub KillFile(FileName As String)
  On Local Error Resume Next
  If Exist(FileName$) Then 'added 7/24
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

Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

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

Public Function Date2Num%(TheDate$)
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function

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
  
  FAItemRecLen = Len(FAItemRec)
  FAHandle = FreeFile
  Open "FAITEMS.DAT" For Random Shared As FAHandle Len = FAItemRecLen
  
  NumOfFARecs = LOF(FAHandle) \ Len(FAItemRec)
  ReDim TempTagIdx(1 To NumOfFARecs) As TagNumbSortIdxType
  
  If NumOfFARecs = 1 Then
    TagIdxRecLen = Len(TagIdx)
    TagIdxHandle = FreeFile
    Open "FATAGIDX.DAT" For Random Shared As TagIdxHandle Len = TagIdxRecLen
    TempTagIdx(1).DataRecNum = 1
    TempTagIdx(1).TagNumb = QPTrim$(FAItemRec.ITEMTAG)
    Put TagIdxHandle, 1, TempTagIdx(1)
    Close
    Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To NumOfFARecs
    Get FAHandle, x, FAItemRec
    TempTagIdx(x).DataRecNum = x
    TempTagIdx(x).TagNumb = QPTrim$(FAItemRec.ITEMTAG)
    ThisNum = Val(ReplaceString(FAItemRec.ITEMTAG, "-", ""))
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
  Next x
  Close FAHandle
  
  BigNum = BigNum + 1
  SmallNum = BigNum
  Nextx = 1
  
  Do
    For x = Nextx To NumOfFARecs
      If QPTrim$(TempTagIdx(x).TagNumb) = "" Then GoTo EString
      ThisNum = Val(ReplaceString(TempTagIdx(x).TagNumb, "-", ""))
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
EString:
    Next x
    HoldThis = TempTagIdx(Nextx)
    TempTagIdx(Nextx) = TempTagIdx(ThisX)
    TempTagIdx(ThisX) = HoldThis
'    If Nextx = NumOfFARecs Then Exit Do
    If Nextx = NumOfFARecs - 1 Then Exit Do
    If Nextx = Int(NumOfFARecs / 4) Then
      frmFixedAssetsConversion.Label2.Caption = "Creating Tag Item Index...25%"
      DoEvents
    End If
    If Nextx = Int(NumOfFARecs / 2) Then
      frmFixedAssetsConversion.Label2.Caption = "Creating Tag Item Index...50%"
      DoEvents
    End If
    If Nextx = Int((NumOfFARecs / 4) * 3) Then
      frmFixedAssetsConversion.Label2.Caption = "Creating Tag Item Index...75%"
      DoEvents
    End If
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
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
  Dim CODEREC As FAAssetCodeRecType
  Dim ACItemRecLen As Integer
  Dim NumOfACRecs As Integer
  Dim AssIdx As ACNumbSortIdxType
  Dim AssIdxHandle As Integer
  Dim AssIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As ACNumbSortIdxType
  
  OpenFACodeNameFile CodeHandle
  
  NumOfACRecs = LOF(CodeHandle) \ Len(CODEREC)
  ReDim TempAssIdx(1 To NumOfACRecs) As ACNumbSortIdxType
  If NumOfACRecs = 1 Then
    OpenAssIdxFile AssIdxHandle
      Get CodeHandle, 1, CODEREC
      TempAssIdx(1).AssNumb = QPTrim$(CODEREC.ASSETCODE)
      TempAssIdx(1).AssRecNum = 1
      Put AssIdxHandle, 1, TempAssIdx(1)
      Close
      Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To NumOfACRecs
    Get CodeHandle, x, CODEREC
    TempAssIdx(x).AssRecNum = x
    TempAssIdx(x).AssNumb = QPTrim$(CODEREC.ASSETCODE)
    ThisNum = Val(QPTrim$(CODEREC.ASSETCODE))
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
    If Nextx = NumOfACRecs - 1 Or NumOfACRecs = 1 Then Exit Do
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

Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
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

Public Sub CheckVersion()
  Dim DOSFAItemRec As DosFAItemRecType
  Dim DOSFAItemRecLen As Integer
  Dim DOSFAHandle As Integer
  Dim DOSFAItemRecV1 As DosFAItemRecTypeV1
  Dim DOSFAItemRecLenV1 As Integer
  Dim DOSFAHandleV1 As Integer
  Dim x As Long
  Dim NumOfFARecs As Integer
  
  If ThisVersion = 0 Then
    frmFixedAssetsConversion.fpList1.Clear
    DOSFAItemRecLen = Len(DOSFAItemRec)
    DOSFAHandle = FreeFile
    Open "FAITEMS.DAT" For Random Shared As DOSFAHandle Len = DOSFAItemRecLen
    
    NumOfFARecs = LOF(DOSFAHandle) / Len(DOSFAItemRec)
    For x = 1 To NumOfFARecs
      Get DOSFAHandle, x, DOSFAItemRec
      If QPTrim$(DOSFAItemRec.VENDOR) <> "" Then
        frmFixedAssetsConversion.fpList1.AddItem (DOSFAItemRec.VENDOR)
      End If
    Next x
    
    Close DOSFAHandle
    
    frmFixedAssetsConversion.cmdConvertNow.Enabled = True
    
  ElseIf ThisVersion = 1 Then
    frmFixedAssetsConversion.fpList1.Clear
    DOSFAItemRecLenV1 = Len(DOSFAItemRecV1)
    DOSFAHandleV1 = FreeFile
    Open "FAITEMS.DAT" For Random Shared As DOSFAHandleV1 Len = DOSFAItemRecLenV1
    
    NumOfFARecs = LOF(DOSFAHandleV1) / Len(DOSFAItemRecV1)
    For x = 1 To NumOfFARecs
      Get DOSFAHandleV1, x, DOSFAItemRecV1
      If QPTrim$(DOSFAItemRecV1.VENDOR) <> "" Then
        frmFixedAssetsConversion.fpList1.AddItem (DOSFAItemRecV1.VENDOR)
      End If
    Next x
    
    Close DOSFAHandle
    
    frmFixedAssetsConversion.cmdConvertNow.Enabled = True
    
  End If
End Sub
