Attribute VB_Name = "ubStartUp"
Option Explicit
'Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Public Const SND_SYNC = &H0
'Public Const SND_ASYNC = &H1
'Public Const SND_NODEFAULT = &H2
'Public Const SND_LOOP = &H8
'Public Const SND_NOSTOP = &H10
'

Sub Main()
  Dim RetValue As Integer
  Dim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim Citipass As CitiPassType
  Dim cnt&, dl&

  DebugMode = False
  
''''  Dim DCnt As Integer, rcnt As Integer
''''  Dim zz As String * 181
''''  Dim UBSenGetRecLen As Integer
''''  Dim NumSenGetRecs As Long
''''  ReDim UBSenGetRdRec(1) As UBGilSensusGetReadRecType
''''  UBSenGetRecLen = Len(UBSenGetRdRec(1))
''''  Open "c:\hhexport\exssi001.dat" For Random Shared As #1 Len = UBSenGetRecLen
''''  NumSenGetRecs = LOF(1) / UBSenGetRecLen
''''  For rcnt = 1 To NumSenGetRecs
''''    Get #1, rcnt, UBSenGetRdRec(1)
''''    If Val(UBSenGetRdRec(1).CurRead) <= 0 Then
''''      If InStr(UBSenGetRdRec(1).CurRead, ".") > 0 Then
''''        DCnt = DCnt + 1
''''      End If
''''    End If
''''  Next
''''  Close
''''  End
  
'  Dim mess As BillOutRec3Type
'  Dim messlen As Integer
'  messlen = Len(mess)
'  MsgBox (messlen)
'  End
  
  Twiddle = "||//--\\"
  
  App.TaskVisible = False        'don't show in task list
  UBPath$ = QPTrim$(App.Path)    'start up path
  
  If Right$(UBPath$, 1) <> "\" Then
    UBPath$ = UBPath$ + "\"
  End If
  If App.PrevInstance Then
     ActivatePrevInstance
  End If
  screenW = (Screen.Width / Screen.TwipsPerPixelX)
  cnt& = 199
  App.HelpFile = "helpfiles\UB.hlp"

  '(*&(^%^&$%^&*&^%$#@#$%^&*^%%$#@##$%^&*^*&%^$%#$@#
  'BackColor = &HD0D0D0
  ComputerName$ = String$(200, 0)
  dl& = GetUserName(ComputerName$, cnt)
  ComputerName$ = QPTrim$(ComputerName$)

  TempIndexName = UBPath$ + "UBTEMP.IDX"
  BookIndexFile = UBPath$ + "UBCUSTBK.IDX"
  NameIndexFile = UBPath$ + "UBCUSTNM.IDX"
  UBCustFile = UBPath$ + "UBCUST.DAT"
  UBOwnerFile = UBPath$ + "UBOWNER.DAT"
  
  CrLf = Chr$(13) + Chr$(10)
  FF = Chr$(12)
  Chr9 = Chr$(9)
  SetTempPWPath

' Call ConvertData
' Stop
  setupnewcode
  LoadUBSetUpFile UBSetUpRec(), RecLen
  TOWNNAME$ = QPTrim$(UBSetUpRec(1).UTILNAME)
  If InStr(TOWNNAME$, "WADESBORO") Then
    WDflag = True
  Else
    WDflag = False
  End If
  If DebugMode = False Then
  If Exist(PassP$) Then
    GetTemp
    If PWcnt = 0 And PWUser$ = "Sosoft Support" Then
      LevelPass = 1
      PWUser = "Sosoft Support"
      PWcnt = 0
      OPERNUM = 0
      LevelAdj = True
      UBLog "Support Sign in"
      Load frmUBMainMenu
      DoEvents
      frmUBMainMenu.Show
    ElseIf PWcnt > 0 Then
    If Len(Dir$("Citipass.dat")) Then
      SetAttr ("CitiPass.dat"), vbNormal
      OpenCitiPassFile CitiPassFile, NumPassRecs
      If Not CitiPassFile = -1 Then
        Get CitiPassFile, PWcnt, Citipass
        If Not Citipass.DelFlag Then
          If Citipass.Module(9).FullAccess = True Then
            LevelPass = 1
          ElseIf Citipass.Module(9).PaymentAccess = True Then
            LevelPass = 2
          ElseIf Citipass.Module(9).ReportsOnly = True Then
            LevelPass = 3
          End If
          If Citipass.Module(9).Adjustments = True Then
            LevelAdj = True
          Else
            LevelAdj = False
          End If
          OPERNUM = Citipass.PassNum
          PWUser = QPTrim(Citipass.UserName)
        End If
      End If
      Close CitiPassFile
    End If
  End If
  End If
    If LevelPass > 0 Then
      Call UBLog("In UB, with Level " & LevelPass)
      DelayExit = True
      Load frmUBMainMenu
      DoEvents
      frmUBMainMenu.Show
    End If
  Else
    LevelPass = 1
    PWUser = "Sosoft Support"
    PWcnt = 0
    OPERNUM = 0
    DelayExit = True
    LevelAdj = True
    Load frmUBMainMenu
    frmUBMainMenu.Show
    DoEvents
  End If
'Only for testing
'    LevelPass = 1
'    PWUser = "Sosoft Support"
'    PWcnt = 0
'    OPERNUM = 0
'    DelayExit = True
'    Load frmUBMainMenu
'    frmUBMainMenu.Show
'    DoEvents
End Sub

Public Sub DoTheTime()
  Dim sec As Long
  sec = Timer
  Do
  Loop Until (sec + 1) < Timer
End Sub
Private Sub setupnewcode()
Dim IndexName As String, IdxRecLen As Integer, NumOfRecs As Long
Dim Handle As Integer, lcnt As Long, ThisBook As Integer, TBooks As Integer
Dim UBCust As Integer, UBCustRecLen As Integer, cnt As Long
Dim ThisCustRec As Long, TBCnt As Integer, WhatBook As Integer
Dim ghandle As Integer, GrpCodeRecLen As Integer, numofbRecs As Integer
Dim cntb As Integer
Dim UBCustRec(1 To 1) As NewUBCustRecType
Dim GroupCde As GroupCodeRecType
If Exist(UBPath$ + "UBSetup.dat") And Exist(UBPath$ + "UBCust.dat") And Exist(BookIndexFile) Then
If Not Exist(UBPath$ + "UBGrpCde.dat") Then
   UBCustRecLen = Len(UBCustRec(1))
   ReDim Bookconsump(0 To 1) As BookGroupType
    IndexName$ = BookIndexFile
    'UBLog "Loading index file: " + IndexName$
    IdxRecLen = 4
    NumOfRecs = FileSize(IndexName$) \ 4
    
    ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
    Handle = FreeFile
    Open IndexName$ For Random Shared As Handle Len = IdxRecLen
    For lcnt& = 1 To NumOfRecs
      Get #Handle, lcnt&, IndexArray(lcnt&)
    Next
    Close Handle
    ThisBook = 0
    TBooks = 0
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  For cnt = 1 To NumOfRecs
    ThisCustRec& = IndexArray(cnt).RecNum
    Get UBCust, ThisCustRec&, UBCustRec(1)
    If UBCustRec(1).DelFlag Then
      GoTo SkipEm
    End If
   If Not Len(QPTrim$(UBCustRec(1).Book)) = 0 Then
      If Val(UBCustRec(1).Book) <> ThisBook Then
        ThisBook = Val(UBCustRec(1).Book)
        If TBooks > 0 Then
          For TBCnt = 1 To TBooks
            If Bookconsump(TBooks).Book = UBCustRec(1).Book Then 'ThisBook Then
              WhatBook = TBCnt
              Exit For
            Else
              TBooks = TBooks + 1
              ReDim Preserve Bookconsump(0 To TBooks) As BookGroupType
              Bookconsump(TBooks).Book = UBCustRec(1).Book 'ThisBook
              WhatBook = TBooks
            End If
          Next
         Else
           TBooks = TBooks + 1
           Bookconsump(TBooks).Book = UBCustRec(1).Book 'ThisBook
           WhatBook = TBooks
         End If
       End If
    End If
SkipEm:
Next
Close
If TBooks > 0 Then
  GrpCodeRecLen = Len(GroupCde)
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  For cnt = 1 To TBooks
    GroupCde.Deleted = 0
    GroupCde.GroupCode = Bookconsump(cnt).Book
    GroupCde.GroupCodeName = "Book " & Bookconsump(cnt).Book
    GroupCde.xtrastuff = ""
    Put #ghandle, cnt, GroupCde
  Next
  'Close ghandle
  numofbRecs = FileSize("UBGrpCde.dat") \ GrpCodeRecLen
  
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  For cnt = 1 To NumOfRecs
    ThisCustRec& = IndexArray(cnt).RecNum
    Get UBCust, ThisCustRec&, UBCustRec(1)
    If UBCustRec(1).DelFlag = 0 Then
      For cntb = 1 To numofbRecs
        Get #ghandle, cntb, GroupCde
        If Val(UBCustRec(1).Book) = Val(GroupCde.GroupCode) Then
          UBCustRec(1).GroupCodeRec = cntb  'recnum to the groupcode
          Put UBCust, ThisCustRec&, UBCustRec(1)
          Exit For
        End If
      Next
    End If
  Next
  Close
  End If
  


End If 'file already exist do nothing
End If ''not setup yet
End Sub

