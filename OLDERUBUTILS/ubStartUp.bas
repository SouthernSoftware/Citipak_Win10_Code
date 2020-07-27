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
  
  DebugMode = False
  
  Twiddle = "||//--\\"
    
  'Load frmCustAddEdit
  
  'RetValue = sndPlaySound("UBToil.dat", SND_ASYNC Or SND_NODEFAULT)
  
  App.TaskVisible = False        'don't show in task list
  UBPath$ = QPTrim$(App.Path)    'start up path
  
  If Right$(UBPath$, 1) <> "\" Then
    UBPath$ = UBPath$ + "\"
  End If
  
  TempIndexName = UBPath$ + "UBTEMP.IDX"
  BookIndexFile = UBPath$ + "UBCUSTBK.IDX"
  NameIndexFile = UBPath$ + "UBCUSTNM.IDX"
  UBCustFile = UBPath$ + "UBCUST.DAT"
  UBOwnerFile = UBPath$ + "UBOWNER.DAT"
  
  CrLf = Chr$(13) + Chr$(10)
  FF = Chr$(12)
  Chr9 = Chr$(9)
  

' Call ConvertData
' Stop
  
  LoadUBSetUpFile UBSetUpRec(), RecLen
  TOWNNAME$ = QPTrim$(UBSetUpRec(1).UTILNAME)
  If DebugMode = False Then
    Load frmSecurityUBUtil
    DoEvents
    frmSecurityUBUtil.Show
    DoEvents
  Else
   ' LevelPass = 1
    Load frmUBEditMenu
    frmUBEditMenu.Show
    DoEvents
  End If
  'Load frmDisplayList
  'frmDisplayList.Show
End Sub

Public Sub DoTheTime()
  Dim sec As Long
  sec = Timer
  Do
  Loop Until (sec + 1) < Timer
End Sub
