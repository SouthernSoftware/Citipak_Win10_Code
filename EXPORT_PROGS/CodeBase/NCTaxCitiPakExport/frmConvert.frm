VERSION 5.00
Begin VB.Form frmConvert 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert DOS Data to WINDOWS Data"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      Caption         =   "       CONVERSION                 COMPLETE         PRESS TO EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1068
      Left            =   4560
      TabIndex        =   4
      Top             =   5184
      Width           =   2556
   End
   Begin VB.TextBox txtUpdate 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3984
      MousePointer    =   11  'Hourglass
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1536
      Width           =   3948
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "BEGIN UPDATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   6348
      TabIndex        =   2
      Top             =   5724
      Width           =   1932
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   3372
      TabIndex        =   1
      Top             =   5724
      Width           =   1932
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   " This procedure updates any DOS data that is not compliant with the WINDOWS operating system."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   2916
      TabIndex        =   0
      Top             =   2700
      Width           =   6060
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2316
      Left            =   2628
      Top             =   2268
      Width           =   6636
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload frmConvert
  DoEvents
  KillFile "prrun.opn"
  End
End Sub

Private Sub cmdConvert_Click()
  Dim OldChkRec As DosPRCheckRecType
  Dim NewChkRec As PRCheckRecType
  Dim OldChecksHandle As Integer
  Dim NewChecksHandle As Integer
  Dim OldChecksRecLen As Integer
  Dim NewChecksRecLen As Integer
  Dim NumOfChkRecs As Long
  Dim OldPrnSet As DosPRNSetupRecType
  Dim NewPrnSet As PRNSetupRecType
  Dim OldWinEmpData2FileHandle As Integer
  Dim cnt As Integer
  Dim OldWinEmpData2FileRec As OldWinEmpData2Type
  Dim NewEmpData2FileRec As EmpData2Type
  Dim OldWinEmpData2RecLen As Integer
  Dim NewEmpData2RecLen As Integer
  Dim NumOfRecs As Integer
  Dim Cnt2 As Integer
  Dim OldPrnLen As Integer, NewPrnLen As Integer
  Dim OldHandle As Integer, NewHandle As Integer
  Dim TempPrnLen As Integer
  Dim EICHandle As Integer
  Dim EICRec As EICRecType
  Dim EICFileRec As EICRecType
  Dim EICRecLen As Integer
  Dim OldWinFileHandle As Integer
  Dim OldWinUnitRec As OldWinUnitFileRecType
  Dim OldWinUnitRecLen As Integer
  Dim FileHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim UnitRecLen As Integer
  Dim SysHandle As Integer
  Dim SysRec As RegDSysFileRecType
  Dim SysRecLen As Integer
  Dim Idx As Integer
  Dim OldDedCnt As Integer
  Dim OldDedFileRec As OldDedCodeRecType
  Dim DedRecLen As Integer
  Dim OldDedHandle As Integer
  Dim NewDedCnt As Integer
  Dim NewDedFileRec As DedCodeRecType
  Dim NewDedHandle As Integer
  Dim NewTransHistFileRec As TransRecType
  Dim NewTransHistRecLen As Integer
  Dim OldTransHistFileRec As DosTransRecType
  Dim OldTransHistRecLen As Integer
  Dim NewTransWorkFileRec As TransRecType
  Dim NewTransWorkRecLen As Integer
  Dim OldTransWorkFileRec As DosTransRecType
  Dim OldTransWorkRecLen As Integer
  Dim NewTransHistFileHandle As Integer
  Dim Y As Integer
  Dim OldTransHistFileHandle As Integer
  Dim OldTransWorkFileHandle As Integer
  Dim NewTransWorkFileHandle As Integer
  Dim RetHandle As Integer
  Dim RetRec As RetireRecType
  Dim RetRecLen As Integer, Nextx As Integer
  Dim OLDTransHistRecCnt As Integer
  Dim OLDTransWorkRecCnt As Integer
  Dim x As Integer, TempDedRec(1 To 12) As Integer
  Dim DraftLen As Integer
  Dim OldW2SetUp As DosW2SetUpType
  Dim OldW2Handle As Integer
  Dim OldW2RecLen As Integer
  Dim NewW2SetUp As W2SetUpType
  Dim NewW2Handle As Integer
  Dim NewW2RecLen As Integer
  Dim OldPRDefaultRec As DosPeriodDefaultRecType
  Dim OldPRDRecLen As Integer
  Dim OldPRDHandle As Integer
  Dim NewPRDefaultRec As PeriodDefaultRecType
  Dim NewPRDHandle As Integer
  Dim NewPRDRecLen As Integer
  
  If CheckForAllFiles = False Then
    Unload frmConvert
    DoEvents
    frmWarnMissingFiles.Show
    Exit Sub
  End If
  
  cmdConvert.MousePointer = vbHourglass
  
  OldPrnLen = Len(OldPrnSet)
  NewPrnLen = Len(NewPrnSet)

  OldTransHistRecLen = Len(OldTransHistFileRec)
  OldWinEmpData2RecLen = Len(OldWinEmpData2FileRec)
  NewEmpData2RecLen = Len(NewEmpData2FileRec)
  
  OldEmpData3RecLen = Len(OldEmpData3FileRec)
  NewEmpData3RecLen = Len(NewEmpData3FileRec)

  OldChecksRecLen = Len(OldChkRec)
  NewChecksRecLen = Len(NewChkRec)
  
  Call TrapForBadDeds
  
  cmdCancel.Enabled = False
  
  OldW2RecLen = Len(OldW2SetUp)
  NewW2RecLen = Len(NewW2SetUp)
  
  OldHandle = FreeFile
  NewHandle = FreeFile
  
  Open "prdata\" + "PREMP2.DAT" For Random Shared As OldHandle Len = OldWinEmpData2RecLen
  
  NumOfRecs = LOF(OldHandle) / Len(OldWinEmpData2FileRec)
  
  ReDim TEmpNo$(1 To NumOfRecs), TEmpSSN$(1 To NumOfRecs), TEMPLNAME$(1 To NumOfRecs)
  ReDim TEMPFNAME$(1 To NumOfRecs), TEmpAddr1$(1 To NumOfRecs), TEMPADDR2$(1 To NumOfRecs)
  ReDim TEmpCity$(1 To NumOfRecs), TEmpState$(1 To NumOfRecs), TEmpZip$(1 To NumOfRecs)
  ReDim TEMPBDAY(1 To NumOfRecs) As Integer, TEMPGENDR$(1 To NumOfRecs), TEMPRACE$(1 To NumOfRecs)
  ReDim TEMPRETNO$(1 To NumOfRecs), TEMPRETTP$(1 To NumOfRecs), TDRAFTCOD$(1 To NumOfRecs)
  ReDim TEMPDDACC$(1 To NumOfRecs), TPRENOTED$(1 To NumOfRecs), TBANKNAME$(1 To NumOfRecs)
  ReDim TBANKLOC$(1 To NumOfRecs), TTRANSIT$(1 To NumOfRecs), TEMPJOB$(1 To NumOfRecs)
  ReDim TEMPWCCLS$(1 To NumOfRecs), TEMPSTATS$(1 To NumOfRecs), TEMPBCODE(1 To NumOfRecs) As Double
  ReDim TEMPPTYPE$(1 To NumOfRecs), TEMPPFREQ$(1 To NumOfRecs), TEMPPRATE(1 To NumOfRecs) As Double
  ReDim TEMPORATE(1 To NumOfRecs) As Double, TEMPHDATE(1 To NumOfRecs) As Integer
  ReDim TEMPRDATE(1 To NumOfRecs) As Integer, TEMPTDATE(1 To NumOfRecs) As Integer
  ReDim TEMPFEDX$(1 To NumOfRecs), TEMPFEDO2$(1 To NumOfRecs), TEMPFEDO1(1 To NumOfRecs) As Double
  ReDim TEMPFEDS$(1 To NumOfRecs), TEMPFEDA(1 To NumOfRecs) As Integer, TEMPFEDAA(1 To NumOfRecs) As Double
  ReDim TEMPSTAX$(1 To NumOfRecs), TEMPSTAO2$(1 To NumOfRecs), TEMPSTAO1(1 To NumOfRecs) As Double
  ReDim TEMPSTAA(1 To NumOfRecs) As Integer, TEMPSTAAA(1 To NumOfRecs) As Double
  ReDim TEMPSTAS$(1 To NumOfRecs), TLastTransRec(1 To NumOfRecs)
  ReDim TEMPSOCX$(1 To NumOfRecs), TEMPMEDX$(1 To NumOfRecs), TempEIC$(1 To NumOfRecs)
  ReDim TempDed(1 To NumOfRecs, 1 To 50) As EmpDedType
  ReDim TEMPEACT1$(1 To NumOfRecs), TEMPEAMT1(1 To NumOfRecs) As Double
  ReDim TEMPEACT2$(1 To NumOfRecs), TEMPEAMT2(1 To NumOfRecs) As Double
  ReDim TEMPEACT3$(1 To NumOfRecs), TEMPEAMT3(1 To NumOfRecs) As Double
  ReDim TEMPHP$(1 To NumOfRecs), TEDist(1 To NumOfRecs, 1 To 8) As EmpWageDistType
  ReDim TEMPVACE(1 To NumOfRecs) As Double, TEMPVUSED(1 To NumOfRecs) As Double
  ReDim TEMPVBAL(1 To NumOfRecs) As Double, TEMPSLE(1 To NumOfRecs) As Double
  ReDim TEMPSLUSE(1 To NumOfRecs) As Double, TEMPSLBAL(1 To NumOfRecs) As Double
  ReDim TEMPCTE(1 To NumOfRecs) As Double, TEMPCTUSE(1 To NumOfRecs) As Double
  ReDim TEMPCTBAL(1 To NumOfRecs) As Double, TPERERN(1 To NumOfRecs) As Double
  ReDim TPerUsed(1 To NumOfRecs) As Double, TPERBAL(1 To NumOfRecs) As Double
  ReDim THOLERN(1 To NumOfRecs) As Double, THolUsed(1 To NumOfRecs) As Double
  ReDim THOLBAL(1 To NumOfRecs) As Double, TLeaveTbl(1 To NumOfRecs) As Double
  ReDim TExcludeESC$(1 To NumOfRecs), TUseLife$(1 To NumOfRecs)
  ReDim TLastTranRec(1 To NumOfRecs) As Integer, TEmpPin(1 To NumOfRecs) As Integer
  ReDim TDeleted(1 To NumOfRecs) As Integer, TLDTDate(1 To NumOfRecs) As Integer
  ReDim TCDTDate(1 To NumOfRecs) As Integer, TInprocFlag(1 To NumOfRecs) As Integer
  ReDim TUnused$(1 To NumOfRecs)
  
  txtUpdate.Visible = True
  DoEvents
  txtUpdate.Text = "Updating Employee Records"
  DoEvents
  For cnt = 1 To NumOfRecs
    Get OldHandle, cnt, OldWinEmpData2FileRec
    If QPTrim$(OldWinEmpData2FileRec.EmpNo) = "" Then
      TEmpNo(cnt) = "0"
    Else
      TEmpNo(cnt) = OldWinEmpData2FileRec.EmpNo
    End If
    If QPTrim$(OldWinEmpData2FileRec.EmpSSN) = "" Then
      TEmpSSN(cnt) = "000-00-000"
    Else
      TEmpSSN(cnt) = OldWinEmpData2FileRec.EmpSSN
    End If
    'Found some situations where employees were zeroed out so they had
    'a value but the program didn't recognize zero as a name
    If QPTrim$(OldWinEmpData2FileRec.EMPLNAME) = "" Or InStr(OldWinEmpData2FileRec.EMPLNAME, "0") Then
      TEMPLNAME(cnt) = "NONE"
    Else
      TEMPLNAME(cnt) = UCase(OldWinEmpData2FileRec.EMPLNAME)
    End If
    If QPTrim$(OldWinEmpData2FileRec.EMPFNAME) = "" Or InStr(OldWinEmpData2FileRec.EMPFNAME, "0") Then
      TEMPFNAME(cnt) = "NONE"
    Else
      TEMPFNAME(cnt) = UCase(OldWinEmpData2FileRec.EMPFNAME)
    End If
    
    If QPTrim$(OldWinEmpData2FileRec.EmpAddr1) = "" Then
      TEmpAddr1(cnt) = "None"
    Else
      TEmpAddr1(cnt) = OldWinEmpData2FileRec.EmpAddr1
    End If
    TEMPADDR2(cnt) = OldWinEmpData2FileRec.EMPADDR2
    
    If QPTrim$(OldWinEmpData2FileRec.EmpCity) = "" Then
      TEmpCity(cnt) = "None"
    Else
      TEmpCity(cnt) = OldWinEmpData2FileRec.EmpCity
    End If
    If QPTrim$(OldWinEmpData2FileRec.EmpState) = "" Then
      TEmpState(cnt) = "None"
    Else
      TEmpState(cnt) = OldWinEmpData2FileRec.EmpState
    End If
    
    If QPTrim$(OldWinEmpData2FileRec.EmpZip) = "" Then
      TEmpZip(cnt) = "00000"
    Else
      TEmpZip(cnt) = OldWinEmpData2FileRec.EmpZip
    End If
    If Val(OldWinEmpData2FileRec.EMPBDAY) <= -22000 Then
      TEMPBDAY(cnt) = 0
    Else
      TEMPBDAY(cnt) = OldWinEmpData2FileRec.EMPBDAY
    End If
    TEMPGENDR(cnt) = OldWinEmpData2FileRec.EMPGENDR
    TEMPRACE(cnt) = OldWinEmpData2FileRec.EMPRACE
    TEMPRETNO(cnt) = OldWinEmpData2FileRec.EMPRETNO
    TEMPRETTP(cnt) = OldWinEmpData2FileRec.EMPRETTP
    If QPTrim$(OldWinEmpData2FileRec.DRAFTCOD) = "0" Then
      OldWinEmpData2FileRec.DRAFTCOD = ""
    End If
    If QPTrim$(OldWinEmpData2FileRec.DRAFTCOD) = "" Then
      TDRAFTCOD(cnt) = ""
      TEMPDDACC(cnt) = ""
      TPRENOTED(cnt) = ""
      TBANKNAME(cnt) = ""
      TBANKLOC(cnt) = ""
      TTRANSIT(cnt) = ""
    Else
      TDRAFTCOD(cnt) = OldWinEmpData2FileRec.DRAFTCOD
      TEMPDDACC(cnt) = OldWinEmpData2FileRec.EMPDDACC
      TPRENOTED(cnt) = OldWinEmpData2FileRec.PRENOTED
      TBANKNAME(cnt) = OldWinEmpData2FileRec.BANKNAME
      TBANKLOC(cnt) = OldWinEmpData2FileRec.BANKLOC
      TTRANSIT(cnt) = OldWinEmpData2FileRec.TRANSIT
    End If
    TEMPJOB(cnt) = OldWinEmpData2FileRec.EMPJOB
    If QPTrim$(OldWinEmpData2FileRec.EMPWCCLS) = "" Then
      TEMPWCCLS(cnt) = "0"
    Else
      TEMPWCCLS(cnt) = OldWinEmpData2FileRec.EMPWCCLS
    End If
    
    TEMPSTATS(cnt) = OldWinEmpData2FileRec.EMPSTATS
    If OldWinEmpData2FileRec.EMPBCODE < 0 Then
      TEMPBCODE(cnt) = 0
    Else
      TEMPBCODE(cnt) = OldWinEmpData2FileRec.EMPBCODE
    End If
    TEMPPTYPE(cnt) = OldWinEmpData2FileRec.EMPPTYPE
    TEMPPFREQ(cnt) = OldWinEmpData2FileRec.EMPPFREQ
    If OldWinEmpData2FileRec.EMPPRATE < 0 Then
      TEMPPRATE(cnt) = 0
    Else
      TEMPPRATE(cnt) = OldWinEmpData2FileRec.EMPPRATE
    End If
    If OldWinEmpData2FileRec.EMPORATE < 0 Then
      TEMPORATE(cnt) = 0
    Else
      TEMPORATE(cnt) = OldWinEmpData2FileRec.EMPORATE
    End If
    If Val(OldWinEmpData2FileRec.EMPHDATE) <= -10957 Then
      TEMPHDATE(cnt) = 0
    Else
      TEMPHDATE(cnt) = OldWinEmpData2FileRec.EMPHDATE
    End If
    
    If OldWinEmpData2FileRec.EMPRDATE < 0 Then
      TEMPRDATE(cnt) = 0
    Else
      TEMPRDATE(cnt) = OldWinEmpData2FileRec.EMPRDATE
    End If
    
    If OldWinEmpData2FileRec.EMPTDATE <= -3652 Then
      TEMPTDATE(cnt) = 0
    Else
      TEMPTDATE(cnt) = OldWinEmpData2FileRec.EMPTDATE
    End If
    If QPTrim$(OldWinEmpData2FileRec.EMPFEDX) = "" Then
      TEMPFEDX(cnt) = "N"
    Else
      TEMPFEDX(cnt) = OldWinEmpData2FileRec.EMPFEDX
    End If
    TEMPFEDO2(cnt) = OldWinEmpData2FileRec.EMPFEDO2
    
    'neutralize situation where there is no value in the Fed Amt/Pct
    'field but there is a value in the Fed figure field
    If Len(QPTrim$(OldWinEmpData2FileRec.EMPFEDX)) = 0 And OldWinEmpData2FileRec.EMPFEDO1 > 0 Then
      TEMPFEDO1(cnt) = 0
      GoTo BadFedAmtPct
    End If
    If OldWinEmpData2FileRec.EMPFEDO1 < -0 Then
      TEMPFEDO1(cnt) = 0
    Else
      TEMPFEDO1(cnt) = OldWinEmpData2FileRec.EMPFEDO1
    End If
BadFedAmtPct:
    If QPTrim$(OldWinEmpData2FileRec.EMPFEDS) = "" Then
      TEMPFEDS(cnt) = "0"
    Else
      TEMPFEDS(cnt) = OldWinEmpData2FileRec.EMPFEDS
    End If
    If OldWinEmpData2FileRec.EMPFEDA < 0 Then
      TEMPFEDA(cnt) = 0
    Else
      TEMPFEDA(cnt) = OldWinEmpData2FileRec.EMPFEDA
    End If
    If OldWinEmpData2FileRec.EMPFEDAA < 0 Then
      TEMPFEDAA(cnt) = 0
    Else
      TEMPFEDAA(cnt) = OldWinEmpData2FileRec.EMPFEDAA
    End If
    
    If QPTrim$(OldWinEmpData2FileRec.EMPSTAX) = "" Then
      TEMPSTAX(cnt) = "N"
    Else
      TEMPSTAX(cnt) = OldWinEmpData2FileRec.EMPSTAX
    End If
    TEMPSTAO2(cnt) = OldWinEmpData2FileRec.EMPSTAO2
    'neutralize situation where there is no value in the State Amt/Pct
    'field but there is a value in the State figure field
    If Len(QPTrim$(OldWinEmpData2FileRec.EMPSTAX)) = 0 And OldWinEmpData2FileRec.EMPSTAO1 > 0 Then
      TEMPSTAO1(cnt) = 0
      GoTo BadStaAmtPct
    End If
    
    If OldWinEmpData2FileRec.EMPSTAO1 < 0 Then
      TEMPSTAO1(cnt) = 0
    Else
      TEMPSTAO1(cnt) = OldWinEmpData2FileRec.EMPSTAO1
    End If
BadStaAmtPct:
    If QPTrim$(OldWinEmpData2FileRec.EMPSTAS) = "" Then
      TEMPSTAS(cnt) = "0"
    Else
      TEMPSTAS(cnt) = OldWinEmpData2FileRec.EMPSTAS
    End If
    If OldWinEmpData2FileRec.EMPSTAA < 0 Then
      TEMPSTAA(cnt) = 0
    Else
      TEMPSTAA(cnt) = OldWinEmpData2FileRec.EMPSTAA
    End If
    If OldWinEmpData2FileRec.EMPSTAAA < 0 Then
      TEMPSTAAA(cnt) = 0
    Else
      TEMPSTAAA(cnt) = OldWinEmpData2FileRec.EMPSTAAA
    End If
    If QPTrim$(OldWinEmpData2FileRec.EMPSOCX) = "" Then
      TEMPSOCX(cnt) = "N"
    Else
      TEMPSOCX(cnt) = OldWinEmpData2FileRec.EMPSOCX
    End If
    If QPTrim$(OldWinEmpData2FileRec.EMPMEDX) = "" Then
      TEMPMEDX(cnt) = "N"
    Else
      TEMPMEDX(cnt) = OldWinEmpData2FileRec.EMPMEDX
    End If
    TempEIC(cnt) = OldWinEmpData2FileRec.EMPEIC
    For Cnt2 = 1 To 50
      If OldWinEmpData2FileRec.EmpDed(Cnt2).DAmt < 0 Then
        TempDed(cnt, Cnt2).DAmt = 0
      Else
        TempDed(cnt, Cnt2).DAmt = OldWinEmpData2FileRec.EmpDed(Cnt2).DAmt
      End If
      If Len(OldWinEmpData2FileRec.EmpDed(Cnt2).DPct) > 0 And Val(OldWinEmpData2FileRec.EmpDed(Cnt2).DAmt) <= 0 Then
        TempDed(cnt, Cnt2).DPct = ""
      Else
        TempDed(cnt, Cnt2).DPct = OldWinEmpData2FileRec.EmpDed(Cnt2).DPct
      End If
    TempDed(cnt, Cnt2).DOTI = OldWinEmpData2FileRec.EmpDed(Cnt2).DOTI
    Next Cnt2
    TEMPEACT1(cnt) = OldWinEmpData2FileRec.EMPEACT1
    If OldWinEmpData2FileRec.EMPEAMT1 < 0 Then
      TEMPEAMT1(cnt) = 0
    Else
      TEMPEAMT1(cnt) = OldWinEmpData2FileRec.EMPEAMT1
    End If
    
    TEMPEACT2(cnt) = OldWinEmpData2FileRec.EMPEACT2
    If OldWinEmpData2FileRec.EMPEAMT2 < 0 Then
      TEMPEAMT2(cnt) = 0
    Else
      TEMPEAMT2(cnt) = OldWinEmpData2FileRec.EMPEAMT2
    End If
    TEMPEACT3(cnt) = OldWinEmpData2FileRec.EMPEACT3
    If OldWinEmpData2FileRec.EMPEAMT3 < -0 Then
      TEMPEAMT3(cnt) = 0
    Else
      TEMPEAMT3(cnt) = OldWinEmpData2FileRec.EMPEAMT3
    End If
    TEMPHP(cnt) = OldWinEmpData2FileRec.EMPHP
    For Cnt2 = 1 To 8
      TEDist(cnt, Cnt2).DAcct = OldWinEmpData2FileRec.EDist(Cnt2).DAcct
      If OldWinEmpData2FileRec.EDist(Cnt2).DAmt < 0 Then
        TEDist(cnt, Cnt2).DAmt = 0
      Else
        TEDist(cnt, Cnt2).DAmt = OldWinEmpData2FileRec.EDist(Cnt2).DAmt
      End If
    Next Cnt2
    
    If OldWinEmpData2FileRec.EMPVACE < -1000 Then
      TEMPVACE(cnt) = 0
    Else
      TEMPVACE(cnt) = OldWinEmpData2FileRec.EMPVACE
    End If
    If OldWinEmpData2FileRec.EMPVUSED < -1000 Then
      TEMPVUSED(cnt) = 0
    Else
      TEMPVUSED(cnt) = OldWinEmpData2FileRec.EMPVUSED
    End If
    If OldWinEmpData2FileRec.EMPVBAL < -1000 Then
      TEMPVBAL(cnt) = 0
    Else
      TEMPVBAL(cnt) = OldWinEmpData2FileRec.EMPVBAL
    End If
    
    If OldWinEmpData2FileRec.EMPSLE < -1000 Then
      TEMPSLE(cnt) = 0
    Else
      TEMPSLE(cnt) = OldWinEmpData2FileRec.EMPSLE
    End If
    If OldWinEmpData2FileRec.EMPSLUSE < -1000 Then
      TEMPSLUSE(cnt) = 0
    Else
      TEMPSLUSE(cnt) = OldWinEmpData2FileRec.EMPSLUSE
    End If
    If OldWinEmpData2FileRec.EMPSLBAL < -1000 Then
      TEMPSLBAL(cnt) = 0
    Else
      TEMPSLBAL(cnt) = OldWinEmpData2FileRec.EMPSLBAL
    End If
    
    If OldWinEmpData2FileRec.EMPCTE < -1000 Then
      TEMPCTE(cnt) = 0
    Else
      TEMPCTE(cnt) = OldWinEmpData2FileRec.EMPCTE
    End If
    If OldWinEmpData2FileRec.EMPCTUSE < -1000 Then
      TEMPCTUSE(cnt) = 0
    Else
      TEMPCTUSE(cnt) = OldWinEmpData2FileRec.EMPCTUSE
    End If
    If OldWinEmpData2FileRec.EMPCTBAL < -1000 Then
      TEMPCTBAL(cnt) = 0
    Else
      TEMPCTBAL(cnt) = OldWinEmpData2FileRec.EMPCTBAL
    End If
    
    If OldWinEmpData2FileRec.PERERN < -1000 Then
      TPERERN(cnt) = 0
    Else
      TPERERN(cnt) = OldWinEmpData2FileRec.PERERN
    End If
    If OldWinEmpData2FileRec.PerUsed < -1000 Then
      TPerUsed(cnt) = 0
    Else
      TPerUsed(cnt) = OldWinEmpData2FileRec.PerUsed
    End If
    If OldWinEmpData2FileRec.PERBAL < -1000 Then
      TPERBAL(cnt) = 0
    Else
      TPERBAL(cnt) = OldWinEmpData2FileRec.PERBAL
    End If
    
    If OldWinEmpData2FileRec.HOLERN < -1000 Then
      THOLERN(cnt) = 0
    Else
      THOLERN(cnt) = OldWinEmpData2FileRec.HOLERN
    End If
    If OldWinEmpData2FileRec.HolUsed < -1000 Then
      THolUsed(cnt) = 0
    Else
      THolUsed(cnt) = OldWinEmpData2FileRec.HolUsed
    End If
    If OldWinEmpData2FileRec.HOLBAL < -1000 Then
      THOLBAL(cnt) = 0
    Else
      THOLBAL(cnt) = OldWinEmpData2FileRec.HOLBAL
    End If
    TLeaveTbl(cnt) = OldWinEmpData2FileRec.LeaveTbl
    If OldWinEmpData2FileRec.ExcludeESC = "0" Then OldWinEmpData2FileRec.ExcludeESC = "N"
    TExcludeESC(cnt) = OldWinEmpData2FileRec.ExcludeESC
    TUseLife(cnt) = OldWinEmpData2FileRec.UseLife
    If OldWinEmpData2FileRec.EmpPin < 0 Then
      TEmpPin(cnt) = 0
    Else
      TEmpPin(cnt) = OldWinEmpData2FileRec.EmpPin
    End If
    TLastTransRec(cnt) = OldWinEmpData2FileRec.LastTransRec
    TDeleted(cnt) = OldWinEmpData2FileRec.Deleted
    TLDTDate(cnt) = OldWinEmpData2FileRec.LDTDate
    TCDTDate(cnt) = OldWinEmpData2FileRec.CDTDate
    TInprocFlag(cnt) = OldWinEmpData2FileRec.InprocFlag
    TUnused(cnt) = OldWinEmpData2FileRec.Unused
  Next cnt

  Close OldHandle
 
'**************************************************************
  Open "prdata\" + "PREMP2.DAT" For Random Shared As NewHandle Len = NewEmpData2RecLen
  
  For cnt = 1 To NumOfRecs
    NewEmpData2FileRec.EmpNo = TEmpNo(cnt)
    NewEmpData2FileRec.EmpSSN = TEmpSSN(cnt)
    NewEmpData2FileRec.EMPLNAME = TEMPLNAME(cnt)
    NewEmpData2FileRec.EMPFNAME = TEMPFNAME(cnt)
    NewEmpData2FileRec.EmpAddr1 = TEmpAddr1(cnt)
    NewEmpData2FileRec.EMPADDR2 = TEMPADDR2(cnt)
    NewEmpData2FileRec.EmpCity = TEmpCity(cnt)
    NewEmpData2FileRec.EmpState = TEmpState(cnt)
    NewEmpData2FileRec.EmpZip = TEmpZip(cnt)
    NewEmpData2FileRec.EMPBDAY = TEMPBDAY(cnt)
    NewEmpData2FileRec.EMPGENDR = TEMPGENDR(cnt)
    NewEmpData2FileRec.EMPRACE = TEMPRACE(cnt)
    NewEmpData2FileRec.EMPRETNO = TEMPRETNO(cnt)
    NewEmpData2FileRec.EMPRETTP = TEMPRETTP(cnt)
    NewEmpData2FileRec.DRAFTCOD = TDRAFTCOD(cnt)
    NewEmpData2FileRec.EMPDDACC = TEMPDDACC(cnt)
    NewEmpData2FileRec.PRENOTED = TPRENOTED(cnt)
    NewEmpData2FileRec.BANKNAME = TBANKNAME(cnt)
    NewEmpData2FileRec.BANKLOC = TBANKLOC(cnt)
    NewEmpData2FileRec.TRANSIT = TTRANSIT(cnt)
    NewEmpData2FileRec.EMPJOB = TEMPJOB(cnt)
    NewEmpData2FileRec.EMPWCCLS = TEMPWCCLS(cnt)
    NewEmpData2FileRec.EMPSTATS = TEMPSTATS(cnt)
    NewEmpData2FileRec.EMPBCODE = TEMPBCODE(cnt)
    NewEmpData2FileRec.EMPPTYPE = TEMPPTYPE(cnt)
    NewEmpData2FileRec.EMPPFREQ = TEMPPFREQ(cnt)
    NewEmpData2FileRec.EMPPRATE = TEMPPRATE(cnt)
    NewEmpData2FileRec.EMPORATE = TEMPORATE(cnt)
    NewEmpData2FileRec.EMPHDATE = TEMPHDATE(cnt)
    NewEmpData2FileRec.EMPRDATE = TEMPRDATE(cnt)
    NewEmpData2FileRec.EMPTDATE = TEMPTDATE(cnt)
    NewEmpData2FileRec.EMPFEDX = TEMPFEDX(cnt)
    NewEmpData2FileRec.EMPFEDO2 = TEMPFEDO2(cnt)
    NewEmpData2FileRec.EMPFEDO1 = TEMPFEDO1(cnt)
    NewEmpData2FileRec.EMPFEDS = TEMPFEDS(cnt)
    NewEmpData2FileRec.EMPFEDA = TEMPFEDA(cnt)
    NewEmpData2FileRec.EMPFEDAA = TEMPFEDAA(cnt)
    NewEmpData2FileRec.EMPSTAX = TEMPSTAX(cnt)
    NewEmpData2FileRec.EMPSTAO2 = TEMPSTAO2(cnt)
    NewEmpData2FileRec.EMPSTAO1 = TEMPSTAO1(cnt)
    NewEmpData2FileRec.EMPSTAS = TEMPSTAS(cnt)
    NewEmpData2FileRec.EMPSTAA = TEMPSTAA(cnt)
    NewEmpData2FileRec.EMPSTAAA = TEMPSTAAA(cnt)
    NewEmpData2FileRec.EMPSOCX = TEMPSOCX(cnt)
    NewEmpData2FileRec.EMPMEDX = TEMPMEDX(cnt)
    NewEmpData2FileRec.EMPEIC = TempEIC(cnt)
    
    For Cnt2 = 1 To 50
      NewEmpData2FileRec.EmpDed(Cnt2).DAmt = TempDed(cnt, Cnt2).DAmt
      NewEmpData2FileRec.EmpDed(Cnt2).DPct = TempDed(cnt, Cnt2).DPct
      NewEmpData2FileRec.EmpDed(Cnt2).DOTI = TempDed(cnt, Cnt2).DOTI
    Next Cnt2
    For Cnt2 = 13 To 50
      NewEmpData2FileRec.EmpDed(Cnt2).DAmt = 0#
      NewEmpData2FileRec.EmpDed(Cnt2).DPct = ""
      NewEmpData2FileRec.EmpDed(Cnt2).DOTI = ""
    Next Cnt2
    
    NewEmpData2FileRec.EMPEACT1 = TEMPEACT1(cnt)
    NewEmpData2FileRec.EMPEAMT1 = TEMPEAMT1(cnt)
    NewEmpData2FileRec.EMPEACT2 = TEMPEACT2(cnt)
    NewEmpData2FileRec.EMPEAMT2 = TEMPEAMT2(cnt)
    NewEmpData2FileRec.EMPEACT3 = TEMPEACT3(cnt)
    NewEmpData2FileRec.EMPEAMT3 = TEMPEAMT3(cnt)
    NewEmpData2FileRec.EMPHP = TEMPHP(cnt)
    For Cnt2 = 1 To 8
      NewEmpData2FileRec.EDist(Cnt2).DAcct = TEDist(cnt, Cnt2).DAcct
      NewEmpData2FileRec.EDist(Cnt2).DAmt = TEDist(cnt, Cnt2).DAmt
    Next Cnt2
    
    NewEmpData2FileRec.EMPVACE = TEMPVACE(cnt)
    NewEmpData2FileRec.EMPVUSED = TEMPVUSED(cnt)
    NewEmpData2FileRec.EMPVBAL = TEMPVBAL(cnt)
    NewEmpData2FileRec.EMPSLE = TEMPSLE(cnt)
    NewEmpData2FileRec.EMPSLUSE = TEMPSLUSE(cnt)
    NewEmpData2FileRec.EMPSLBAL = TEMPSLBAL(cnt)
    NewEmpData2FileRec.EMPCTE = TEMPCTE(cnt)
    NewEmpData2FileRec.EMPCTUSE = TEMPCTUSE(cnt)
    NewEmpData2FileRec.EMPCTBAL = TEMPCTBAL(cnt)
    NewEmpData2FileRec.PERERN = TPERERN(cnt)
    NewEmpData2FileRec.PerUsed = TPerUsed(cnt)
    NewEmpData2FileRec.PERBAL = TPERBAL(cnt)
    NewEmpData2FileRec.HOLERN = THOLERN(cnt)
    NewEmpData2FileRec.HolUsed = THolUsed(cnt)
    NewEmpData2FileRec.HOLBAL = THOLBAL(cnt)
    NewEmpData2FileRec.LeaveTbl = TLeaveTbl(cnt)
    If TExcludeESC(cnt) = "0" Then TExcludeESC(cnt) = "N"
    NewEmpData2FileRec.ExcludeESC = TExcludeESC(cnt)
    NewEmpData2FileRec.UseLife = TUseLife(cnt)
    NewEmpData2FileRec.EmpPin = TEmpPin(cnt)
    NewEmpData2FileRec.LastTransRec = TLastTransRec(cnt)
    NewEmpData2FileRec.Deleted = TDeleted(cnt)
    NewEmpData2FileRec.LDTDate = TLDTDate(cnt)
    NewEmpData2FileRec.CDTDate = TCDTDate(cnt)
    NewEmpData2FileRec.InprocFlag = TInprocFlag(cnt)
    NewEmpData2FileRec.Unused = TUnused(cnt)
    NewEmpData2FileRec.CheckType = 10
    NewEmpData2FileRec.YN401K = "N"   'As String * 1
    NewEmpData2FileRec.PrimeDept = "1" 'As String * 6
    NewEmpData2FileRec.HomePhone = "555-5555" 'As String * 14
    NewEmpData2FileRec.EmrgncyCntctName = "AnyName" 'As String * 48
    NewEmpData2FileRec.EmrgncyCntctPhnNum = "555-5555" 'As String * 14
    NewEmpData2FileRec.EmrgncyCntctRelation = "Spouse" 'As String * 14
    
    Put NewHandle, cnt, NewEmpData2FileRec
  Next cnt
  Close NewHandle
'*************************************************************
  Old3Handle = FreeFile
  New3Handle = FreeFile

  Open "prdata\" + "PREMP3.DAT" For Random Shared As Old3Handle Len = OldEmpData3RecLen

  NumOfRecs = LOF(Old3Handle) / Len(OldEmpData3FileRec)
  ReDim TempData1RecNum(1 To NumOfRecs) As Integer
  ReDim TempYTDGrossPay(1 To NumOfRecs) As Double
  ReDim TempYTDSocGrossPay(1 To NumOfRecs) As Double
  ReDim TempYTDMedGrossPay(1 To NumOfRecs) As Double
  ReDim TempYTDFedGrossPay(1 To NumOfRecs) As Double
  ReDim TempYTDStaGrossPay(1 To NumOfRecs) As Double
  ReDim TempYTDOTPay(1 To NumOfRecs) As Double
  ReDim TempYTDRegPay(1 To NumOfRecs) As Double
  ReDim TempYTDNet(1 To NumOfRecs) As Double
  ReDim TempYTDSocial(1 To NumOfRecs) As Double
  ReDim TempYTDMedicare(1 To NumOfRecs) As Double
  ReDim TempYTDFederal(1 To NumOfRecs) As Double
  ReDim TempYTDState(1 To NumOfRecs) As Double
  ReDim TempYTDRetire(1 To NumOfRecs) As Double
  ReDim TempYTDDAmt(1 To NumOfRecs, 1 To 12) As Double
  ReDim TempYTDDAmtT(1 To NumOfRecs) As Double
  ReDim TempYTDEarn1(1 To NumOfRecs) As Double
  ReDim TempYTDEarn2(1 To NumOfRecs) As Double
  ReDim TempYTDEarn3(1 To NumOfRecs) As Double
  ReDim TempYTDEarnT(1 To NumOfRecs) As Double
  ReDim TempYTDEIC(1 To NumOfRecs) As Double
  ReDim TempYTDOther2(1 To NumOfRecs) As Double

  For cnt = 1 To NumOfRecs
    Get Old3Handle, cnt, OldEmpData3FileRec
    If OldEmpData3FileRec.Data1RecNum > 0 Then
      TempData1RecNum(cnt) = OldEmpData3FileRec.Data1RecNum
    Else
      TempData1RecNum(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDGrossPay > 0 Then
      TempYTDGrossPay(cnt) = OldEmpData3FileRec.YTDGrossPay
    Else
      TempYTDGrossPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDSocGrossPay > 0 Then
      TempYTDSocGrossPay(cnt) = OldEmpData3FileRec.YTDSocGrossPay
    Else
      TempYTDSocGrossPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDMedGrossPay > 0 Then
      TempYTDMedGrossPay(cnt) = OldEmpData3FileRec.YTDMedGrossPay
    Else
      TempYTDMedGrossPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDFedGrossPay > 0 Then
      TempYTDFedGrossPay(cnt) = OldEmpData3FileRec.YTDFedGrossPay
    Else
      TempYTDFedGrossPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDStaGrossPay > 0 Then
      TempYTDStaGrossPay(cnt) = OldEmpData3FileRec.YTDStaGrossPay
    Else
      TempYTDStaGrossPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDOTPay > 0 Then
      TempYTDOTPay(cnt) = OldEmpData3FileRec.YTDOTPay
    Else
      TempYTDOTPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDRegPay > 0 Then
      TempYTDRegPay(cnt) = OldEmpData3FileRec.YTDRegPay
    Else
      TempYTDRegPay(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDNet > 0 Then
      TempYTDNet(cnt) = OldEmpData3FileRec.YTDNet
    Else
      TempYTDNet(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDSocial > 0 Then
      TempYTDSocial(cnt) = OldEmpData3FileRec.YTDSocial
    Else
      TempYTDSocial(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDMedicare > 0 Then
      TempYTDMedicare(cnt) = OldEmpData3FileRec.YTDMedicare
    Else
      TempYTDMedicare(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDFederal > 0 Then
      TempYTDFederal(cnt) = OldEmpData3FileRec.YTDFederal
    Else
      TempYTDFederal(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDState > 0 Then
      TempYTDState(cnt) = OldEmpData3FileRec.YTDState
    Else
      TempYTDState(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDRetire > 0 Then
      TempYTDRetire(cnt) = OldEmpData3FileRec.YTDRetire
    Else
      TempYTDRetire(cnt) = 0
    End If

    For Cnt2 = 1 To 12
      If OldEmpData3FileRec.YTDDAmt(Cnt2) > 0 Then
        TempYTDDAmt(cnt, Cnt2) = OldEmpData3FileRec.YTDDAmt(Cnt2)
      Else
        TempYTDDAmt(cnt, Cnt2) = 0
      End If
    Next Cnt2

    If OldEmpData3FileRec.YTDDAmtT > 0 Then
      TempYTDDAmtT(cnt) = OldEmpData3FileRec.YTDDAmtT
    Else
      TempYTDDAmtT(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDEarn1 > 0 Then
      TempYTDEarn1(cnt) = OldEmpData3FileRec.YTDEarn1
    Else
      TempYTDEarn1(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDEarn2 > 0 Then
      TempYTDEarn2(cnt) = OldEmpData3FileRec.YTDEarn2
    Else
      TempYTDEarn2(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDEarn3 > 0 Then
      TempYTDEarn3(cnt) = OldEmpData3FileRec.YTDEarn3
    Else
      TempYTDEarn3(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDEarnT > 0 Then
      TempYTDEarnT(cnt) = OldEmpData3FileRec.YTDEarnT
    Else
      TempYTDEarnT(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDEIC > 0 Then
      TempYTDEIC(cnt) = OldEmpData3FileRec.YTDEIC
    Else
      TempYTDEIC(cnt) = 0
    End If

    If OldEmpData3FileRec.YTDOther2 > 0 Then
      TempYTDOther2(cnt) = OldEmpData3FileRec.YTDOther2
    Else
      TempYTDOther2(cnt) = 0
    End If
  Next cnt
  Close Old3Handle

  Open "prdata\" + "PREMP3.DAT" For Random Shared As New3Handle Len = NewEmpData3RecLen

  For cnt = 1 To NumOfRecs
    NewEmpData3FileRec.Data1RecNum = TempData1RecNum(cnt)
    NewEmpData3FileRec.YTDGrossPay = TempYTDGrossPay(cnt)
    NewEmpData3FileRec.YTDSocGrossPay = TempYTDSocGrossPay(cnt)
    NewEmpData3FileRec.YTDMedGrossPay = TempYTDMedGrossPay(cnt)
    NewEmpData3FileRec.YTDFedGrossPay = TempYTDFedGrossPay(cnt)
    NewEmpData3FileRec.YTDStaGrossPay = TempYTDStaGrossPay(cnt)
    NewEmpData3FileRec.YTDOTPay = TempYTDOTPay(cnt)
    NewEmpData3FileRec.YTDRegPay = TempYTDRegPay(cnt)
    NewEmpData3FileRec.YTDNet = TempYTDNet(cnt)
    NewEmpData3FileRec.YTDSocial = TempYTDSocial(cnt)
    NewEmpData3FileRec.YTDMedicare = TempYTDMedicare(cnt)
    NewEmpData3FileRec.YTDFederal = TempYTDFederal(cnt)
    NewEmpData3FileRec.YTDState = TempYTDState(cnt)
    NewEmpData3FileRec.YTDRetire = TempYTDRetire(cnt)
    For Cnt2 = 1 To 12
      NewEmpData3FileRec.YTDDAmt(Cnt2) = TempYTDDAmt(cnt, Cnt2)
    Next Cnt2

    For Cnt2 = 13 To 50
      NewEmpData3FileRec.YTDDAmt(Cnt2) = 0
    Next Cnt2

    NewEmpData3FileRec.YTDDAmtT = TempYTDDAmtT(cnt)
    NewEmpData3FileRec.YTDState = TempYTDState(cnt)
    NewEmpData3FileRec.YTDEarn1 = TempYTDEarn1(cnt)
    NewEmpData3FileRec.YTDEarn2 = TempYTDEarn2(cnt)
    NewEmpData3FileRec.YTDEarn3 = TempYTDEarn3(cnt)
    NewEmpData3FileRec.YTDEIC = TempYTDEIC(cnt)
    NewEmpData3FileRec.YTDOther2 = TempYTDOther2(cnt)
  Put New3Handle, cnt, NewEmpData3FileRec
  Next cnt
  Close New3Handle
   
  Dim TempPACTIVE  As Integer
  Dim TempPERBEG   As Integer
  Dim TempPEREND   As Integer
  Dim TempUSEDEF   As String * 1

  Dim TempPAYWK    As String * 1
  Dim TempPAYBIWK  As String * 1
  Dim TempPAYSEMIM As String * 1
  Dim TempPAYMO    As String * 1
  Dim TempPAYQTR   As String * 1
  Dim TempPAYSEMIA As String * 1
  Dim TempPAYANNL  As String * 1

  Dim TempUseDed(1 To 12)   As String * 1
  Dim TempUSEAE1   As String * 1
  Dim TempUSEAE2   As String * 1
  Dim TempUSEAE3   As String * 1
  Dim TempMACTIVE  As Integer

  OldPRDRecLen = Len(OldPRDefaultRec)
  NewPRDRecLen = Len(NewPRDefaultRec)
  OldPRDHandle = FreeFile

  If Exist("prdata\" + "PRPPDef.DAT") Then
    Open "prdata\" + "PRPPDef.DAT" For Random Shared As OldPRDHandle Len = OldPRDRecLen
    Get OldPRDHandle, 1, OldPRDefaultRec
    Close OldPRDHandle
  Else
    MsgBox "PRData\PRPPDef.DAT could not be found"
    Exit Sub
  End If

  TempPACTIVE = OldPRDefaultRec.PACTIVE
  TempPERBEG = OldPRDefaultRec.PERBEG
  TempPEREND = OldPRDefaultRec.PEREND
  TempUSEDEF = OldPRDefaultRec.USEDEF

  TempPAYWK = OldPRDefaultRec.PAYWK
  TempPAYBIWK = OldPRDefaultRec.PAYBIWK
  TempPAYSEMIM = OldPRDefaultRec.PAYSEMIM
  TempPAYMO = OldPRDefaultRec.PAYMO
  TempPAYQTR = OldPRDefaultRec.PAYQTR
  TempPAYSEMIA = OldPRDefaultRec.PAYSEMIA
  TempPAYANNL = OldPRDefaultRec.PAYANNL
  For x = 1 To 12
    TempUseDed(x) = OldPRDefaultRec.UseDed(x)
  Next x
  TempUSEAE1 = OldPRDefaultRec.USEAE1
  TempUSEAE2 = OldPRDefaultRec.USEAE2
  TempUSEAE3 = OldPRDefaultRec.USEAE3
  TempMACTIVE = OldPRDefaultRec.MACTIVE

  NewPRDHandle = FreeFile
  Open "prdata\" + "PRPPDef.DAT" For Random Shared As NewPRDHandle Len = NewPRDRecLen
  NewPRDefaultRec.PACTIVE = TempPACTIVE
  NewPRDefaultRec.PERBEG = TempPERBEG
  NewPRDefaultRec.PEREND = TempPEREND
  NewPRDefaultRec.USEDEF = TempUSEDEF
  If QPTrim$(TempPAYWK) = "" Then TempPAYWK = "N"
  NewPRDefaultRec.PAYWK = TempPAYWK
  If QPTrim$(TempPAYBIWK) = "" Then TempPAYBIWK = "N"
  NewPRDefaultRec.PAYBIWK = TempPAYBIWK
  If QPTrim$(TempPAYSEMIM) = "" Then TempPAYSEMIM = "N"
  NewPRDefaultRec.PAYSEMIM = TempPAYSEMIM
  If QPTrim$(TempPAYMO) = "" Then TempPAYMO = "N"
  NewPRDefaultRec.PAYMO = TempPAYMO
  If QPTrim$(TempPAYQTR) = "" Then TempPAYQTR = "N"
  NewPRDefaultRec.PAYQTR = TempPAYQTR
  If QPTrim$(TempPAYSEMIA) = "" Then TempPAYSEMIA = "N"
  NewPRDefaultRec.PAYSEMIA = TempPAYSEMIA
  If QPTrim$(TempPAYANNL) = "" Then TempPAYANNL = "N"
  NewPRDefaultRec.PAYANNL = TempPAYANNL
  For x = 1 To 12
    NewPRDefaultRec.UseDed(x) = TempUseDed(x)
  Next x
  For x = 13 To 50
    NewPRDefaultRec.UseDed(x) = "N"
  Next x
  NewPRDefaultRec.USEAE1 = TempUSEAE1
  NewPRDefaultRec.USEAE2 = TempUSEAE2
  NewPRDefaultRec.USEAE3 = TempUSEAE3
  NewPRDefaultRec.MACTIVE = TempMACTIVE

  Put NewPRDHandle, 1, NewPRDefaultRec
  Close NewPRDHandle
  
Here:
  'Had to convert Deduction records because some were coming
  'over as empty fields which screwed up the way deductions
  'appeared in the payroll processing edit transactions screen
  OldDedHandle = FreeFile
  DedRecLen = Len(OldDedFileRec)
  Open "prdata\" + "PRDEDCOD.DAT" For Random Shared As OldDedHandle Len = DedRecLen
  OldDedCnt = LOF(OldDedHandle) / Len(OldDedFileRec)

  ReDim TempDCDESC1(1 To OldDedCnt) As String * 10
  ReDim TempDCACCT1(1 To OldDedCnt) As String * 14
  ReDim TempDCFWT1(1 To OldDedCnt) As String * 1
  ReDim TempDCSWT1(1 To OldDedCnt) As String * 1
  ReDim TempDCSOC1(1 To OldDedCnt) As String * 1
  ReDim TempDCMED1(1 To OldDedCnt) As String * 1

  For x = 1 To OldDedCnt
    Get OldDedHandle, x, OldDedFileRec
    TempDCDESC1(x) = OldDedFileRec.DCDESC1
    TempDCACCT1(x) = OldDedFileRec.DCACCT1
    TempDCFWT1(x) = OldDedFileRec.DCFWT1
    TempDCSWT1(x) = OldDedFileRec.DCSWT1
    TempDCSOC1(x) = OldDedFileRec.DCSOC1
    TempDCMED1(x) = OldDedFileRec.DCMED1
  Next x
  Close OldDedHandle

  NewDedHandle = FreeFile
  DedRecLen = Len(NewDedFileRec)
  KillFile "prdata\" + "PRDEDCOD.DAT"
  Open "prdata\" + "PRDEDCOD.DAT" For Random Shared As NewDedHandle Len = DedRecLen

  Nextx = 1
  For x = 1 To OldDedCnt
    Get NewDedHandle, x, NewDedFileRec
    If QPTrim(TempDCDESC1(x)) = "" Then GoTo NoDesc
    NewDedFileRec.DCDESC1 = QPTrim$(TempDCDESC1(x))
    NewDedFileRec.DCACCT1 = QPTrim$(TempDCACCT1(x))
    NewDedFileRec.DCFWT1 = QPTrim$(TempDCFWT1(x))
    NewDedFileRec.DCSWT1 = QPTrim$(TempDCSWT1(x))
    NewDedFileRec.DCSOC1 = QPTrim$(TempDCSOC1(x))
    NewDedFileRec.DCMED1 = QPTrim$(TempDCMED1(x))
    Put NewDedHandle, Nextx, NewDedFileRec
    Nextx = Nextx + 1
NoDesc:
  Next x
  NewDedCnt = LOF(NewDedHandle) / Len(NewDedFileRec)

  Close NewDedHandle
'*************************************************************
  txtUpdate.Text = "Updating EIC Records"
  DoEvents

  EICRecLen = Len(EICFileRec)
  EICHandle = FreeFile
  Open "PRData\PREICTBL.DAT" For Random Shared As EICHandle Len = EICRecLen

  Get EICHandle, 1, EICRec
  If EICRec.EIC(1).EIC1AMT0 < -2000000 Then EICRec.EIC(1).EIC1AMT0 = 0
  If EICRec.EIC(1).EIC1AMT1 < -2000000 Then EICRec.EIC(1).EIC1AMT1 = 0
  If EICRec.EIC(1).EIC1AMT2 < -2000000 Then EICRec.EIC(1).EIC1AMT2 = 0
  If EICRec.EIC(1).EIC1EXES < -2000000 Then EICRec.EIC(1).EIC1EXES = 0
  If EICRec.EIC(1).EIC1LESS < -2000000 Then EICRec.EIC(1).EIC1LESS = 0
  If EICRec.EIC(1).EIC1NVR0 < -2000000 Then EICRec.EIC(1).EIC1NVR0 = 0
  If EICRec.EIC(1).EIC1NVR1 < -2000000 Then EICRec.EIC(1).EIC1NVR1 = 0
  If EICRec.EIC(1).EIC1NVR2 < -2000000 Then EICRec.EIC(1).EIC1NVR2 = 0
  If EICRec.EIC(1).EIC1OVR0 < -2000000 Then EICRec.EIC(1).EIC1OVR0 = 0
  If EICRec.EIC(1).EIC1OVR1 < -2000000 Then EICRec.EIC(1).EIC1OVR1 = 0
  If EICRec.EIC(1).EIC1OVR2 < -2000000 Then EICRec.EIC(1).EIC1OVR2 = 0

  If EICRec.EIC(2).EIC1AMT0 < -2000000 Then EICRec.EIC(2).EIC1AMT0 = 0
  If EICRec.EIC(2).EIC1AMT1 < -2000000 Then EICRec.EIC(2).EIC1AMT1 = 0
  If EICRec.EIC(2).EIC1AMT2 < -2000000 Then EICRec.EIC(2).EIC1AMT2 = 0
  If EICRec.EIC(2).EIC1EXES < -2000000 Then EICRec.EIC(2).EIC1EXES = 0
  If EICRec.EIC(2).EIC1LESS < -2000000 Then EICRec.EIC(2).EIC1LESS = 0
  If EICRec.EIC(2).EIC1NVR0 < -2000000 Then EICRec.EIC(2).EIC1NVR0 = 0
  If EICRec.EIC(2).EIC1NVR1 < -2000000 Then EICRec.EIC(2).EIC1NVR1 = 0
  If EICRec.EIC(2).EIC1NVR2 < -2000000 Then EICRec.EIC(2).EIC1NVR2 = 0
  If EICRec.EIC(2).EIC1OVR0 < -2000000 Then EICRec.EIC(2).EIC1OVR0 = 0
  If EICRec.EIC(2).EIC1OVR1 < -2000000 Then EICRec.EIC(2).EIC1OVR1 = 0
  If EICRec.EIC(2).EIC1OVR2 < -2000000 Then EICRec.EIC(2).EIC1OVR2 = 0
  Put EICHandle, 1, EICRec
  Close EICHandle

  txtUpdate.Text = "Updating Printer Settings"
  DoEvents

  OldHandle = FreeFile
  Open "prdata\prprnset.dat" For Binary As OldHandle
  TempPrnLen = LOF(OldHandle)
  Close OldHandle
  If TempPrnLen > OldPrnLen Then
    GoTo SkipToNext
  End If

  OldHandle = FreeFile
  Open "prdata\prprnset.dat" For Random As OldHandle Len = OldPrnLen
  NewHandle = FreeFile
  Open "prdata\prprnset.dat" For Random As NewHandle Len = NewPrnLen
  Get #OldHandle, 1, OldPrnSet
  NewPrnSet.Printer = "None"
  NewPrnSet.RPT(1) = OldPrnSet.RPT1
  NewPrnSet.RPT(2) = OldPrnSet.RPT2
  NewPrnSet.RPT(3) = OldPrnSet.RPT3
  NewPrnSet.RPT(4) = OldPrnSet.RPT4
  NewPrnSet.RPT(5) = OldPrnSet.RPT5
  NewPrnSet.RPT(6) = OldPrnSet.RPT6
  NewPrnSet.RPT(7) = OldPrnSet.RPT7
  NewPrnSet.RPT(8) = OldPrnSet.RPT8
  NewPrnSet.RPT(9) = OldPrnSet.RPT9
  NewPrnSet.RPT(10) = OldPrnSet.RPT10
  NewPrnSet.RPT(11) = OldPrnSet.RPT11
  NewPrnSet.RPT(12) = OldPrnSet.RPT12
  NewPrnSet.RPT(13) = OldPrnSet.RPT13
  NewPrnSet.RPT(14) = OldPrnSet.RPT14
  NewPrnSet.RPT(15) = OldPrnSet.RPT15
  NewPrnSet.RPT(16) = OldPrnSet.RPT16
  NewPrnSet.RPT(17) = 10
  NewPrnSet.RPT(18) = 10
'  NewPrnSet.RPT(19) = 10 '8/14 added
  NewPrnSet.CheckType = 10
  Put #NewHandle, 1, NewPrnSet
  Close OldHandle, NewHandle

SkipToNext:
  txtUpdate.Text = "Updating System Settings"
  DoEvents

  SysHandle = FreeFile
  SysRecLen = Len(SysRec)
  Open "prdata\prsys.dat" For Random As SysHandle Len = SysRecLen
  Get SysHandle, 1, SysRec

  If QPTrim$(SysRec.USEIMP) = "" Then SysRec.USEIMP = "P"
  If QPTrim$(SysRec.CASHACCT) = "" Then SysRec.CASHACCT = "0"
  If QPTrim$(SysRec.IDRACCT) = "" Then SysRec.IDRACCT = "0"
  If QPTrim$(SysRec.ICRACCT) = "" Then SysRec.ICRACCT = "0"
  For cnt = 1 To 5
    If QPTrim$(SysRec.Liab(cnt).Acct) = "" Then SysRec.Liab(cnt).Acct = 0
  Next cnt
  If QPTrim$(SysRec.CITIDIR) = "" Then SysRec.CITIDIR = "None"
  If QPTrim$(SysRec.SplitFlag) = "" Then SysRec.SplitFlag = "N"
  If QPTrim$(SysRec.EXPMETHD) = "" Then SysRec.EXPMETHD = "0"
  If SysRec.FRNGRATE < 0 Then SysRec.FRNGRATE = 0
  If QPTrim$(SysRec.FRNGEXP) = "" Then SysRec.FRNGEXP = "0"
  If QPTrim$(SysRec.FRNGDR) = "" Then SysRec.FRNGDR = "0"
  If QPTrim$(SysRec.FRNGCR) = "" Then SysRec.FRNGCR = "0"
  If SysRec.INDRATE < 0 Then SysRec.INDRATE = 0
  If QPTrim$(SysRec.INDEXP) = "" Then SysRec.INDEXP = "0"
  If QPTrim$(SysRec.INDDR) = "" Then SysRec.INDDR = "0"
  If QPTrim$(SysRec.INDCR) = "" Then SysRec.INDCR = "0"
  If QPTrim$(SysRec.SOCEXP) = "" Then SysRec.SOCEXP = "0"
  If QPTrim$(SysRec.SOCLIAB) = "" Then SysRec.SOCLIAB = "0"
  If QPTrim$(SysRec.MEDEXP) = "" Then SysRec.MEDEXP = "0"
  If QPTrim$(SysRec.MEDLIAB) = "" Then SysRec.MEDLIAB = "0"
  If QPTrim$(SysRec.RETEXP) = "" Then SysRec.RETEXP = "0"
  If QPTrim$(SysRec.RETLIAB) = "" Then SysRec.RETLIAB = "0"
  If SysRec.AcctCnt < 0 Then SysRec.AcctCnt = 0
  If SysRec.GLActLen < 0 Then SysRec.GLActLen = 0
  SysRec.CheckStyle = 0
  SysRec.GLCheckYN = "Y"
'  SysRec.VAC2SICK = "N"
  Put SysHandle, 1, SysRec
  Close SysHandle
  
  Dim TempUFEMPR   As String
  Dim TempUFATTN   As String
  Dim TempUFADDR1  As String
  Dim TempUFADDR2  As String
  Dim TempUFCITY   As String
  Dim TempUFSTATE  As String
  Dim TempUFZIP    As String
  Dim TempUFFEDID  As String
  Dim TempUFSTAID  As String
  Dim TempUFRETID  As String
  Dim TempUFRETIDL As String
  Dim TempESCRTYPE As Integer
  Dim TempTAXWBASE As Double
  Dim TempBBTCNTNO As String
  Dim TempBBTBATCH As String
  Dim TempUSEACH   As String
  Dim TempIMMDNUM  As String
  Dim TempIMMONUM  As String
  Dim TempDBANK    As String
  Dim TempOBANK    As String
  Dim TempFileVer  As Double
    
  OldWinUnitRecLen = Len(UnitRec)
  OldWinFileHandle = FreeFile
  Open "PRData\PRUNIT.DAT" For Random Shared As OldWinFileHandle Len = OldWinUnitRecLen
  Get OldWinFileHandle, 1, OldWinUnitRec
  txtUpdate.Text = "Updating Unit Data"
  DoEvents

  If QPTrim$(OldWinUnitRec.BBTBATCH) = "" Then
    TempBBTBATCH = "0"
  Else
    TempBBTBATCH = QPTrim$(OldWinUnitRec.BBTBATCH)
  End If
  Close OldWinFileHandle
  
  If QPTrim$(OldWinUnitRec.BBTCNTNO) = "" Then
    TempBBTCNTNO = "0"
  Else
    TempBBTCNTNO = QPTrim$(OldWinUnitRec.BBTCNTNO)
  End If
  
  If Len(OldWinUnitRec.FileVer) = 0 Then
    TempFileVer = 0
  Else
     TempFileVer = OldWinUnitRec.FileVer
  End If
  
  If QPTrim$(OldWinUnitRec.UFCITY) = "" Then
    TempUFCITY = "YourTown"
  Else
    TempUFCITY = QPTrim$(OldWinUnitRec.UFCITY)
  End If
  
  If QPTrim$(OldWinUnitRec.UFEMPR) = "" Then
    TempUFEMPR = "YourTown"
  Else
    TempUFEMPR = QPTrim$(OldWinUnitRec.UFEMPR)
  End If
  
  If QPTrim$(OldWinUnitRec.UFFEDID) = "" Then
    TempUFFEDID = "0"
  Else
    TempUFFEDID = QPTrim$(OldWinUnitRec.UFFEDID)
  End If
  
  If QPTrim$(OldWinUnitRec.UFSTAID) = "" Then
    TempUFSTAID = "0"
  Else
    TempUFSTAID = QPTrim$(OldWinUnitRec.UFSTAID)
  End If
  
  If QPTrim$(OldWinUnitRec.UFRETID) = "" Then
    TempUFRETID = "0"
  Else
    TempUFRETID = QPTrim$(OldWinUnitRec.UFRETID)
  End If
  
  If QPTrim$(OldWinUnitRec.UFRETIDL) = "" Then
    TempUFRETIDL = "0"
  Else
    TempUFRETIDL = QPTrim$(OldWinUnitRec.UFRETIDL)
  End If
  
  If QPTrim$(OldWinUnitRec.IMMDNUM) = "" Then
    TempIMMDNUM = "0"
  Else
    TempIMMDNUM = QPTrim$(OldWinUnitRec.IMMDNUM)
  End If
  
  If QPTrim$(OldWinUnitRec.IMMONUM) = "" Then
    TempIMMONUM = "0"
  Else
    TempIMMONUM = QPTrim$(OldWinUnitRec.IMMONUM)
  End If
  
  If OldWinUnitRec.ESCRTYPE < 0 Then
    TempESCRTYPE = 0
  Else
    TempESCRTYPE = OldWinUnitRec.ESCRTYPE
  End If
  
  If OldWinUnitRec.TAXWBASE < 0 Then
    TempTAXWBASE = 0
  Else
    TempTAXWBASE = OldWinUnitRec.TAXWBASE
  End If
  
  If QPTrim$(OldWinUnitRec.UFSTATE) = "" Then
    TempUFSTATE = "NC"
  Else
    TempUFSTATE = QPTrim$(OldWinUnitRec.UFSTATE)
  End If
  
  UnitRecLen = Len(UnitRec)
  FileHandle = FreeFile
  Open "PRData\PRUNIT.DAT" For Random Shared As FileHandle Len = UnitRecLen
  Get FileHandle, 1, UnitRec
  UnitRec.BBTBATCH = TempBBTBATCH
  UnitRec.BBTCNTNO = TempBBTCNTNO
  UnitRec.FileVer = 0
  UnitRec.UFCITY = TempUFCITY
  UnitRec.UFEMPR = TempUFEMPR
  UnitRec.UFFEDID = TempUFFEDID
  UnitRec.UFSTAID = TempUFSTAID
  UnitRec.UFRETID = TempUFRETID
  UnitRec.UFRETIDL = TempUFRETIDL
  UnitRec.IMMDNUM = TempIMMDNUM
  UnitRec.IMMONUM = TempIMMONUM
  UnitRec.ESCRTYPE = TempESCRTYPE
  UnitRec.TAXWBASE = TempTAXWBASE
  UnitRec.UFSTATE = UnitRec.UFSTATE
  DraftLen = FileSize("PRDATA\PRDRAFTI.DAT")
  If DraftLen > 0 Then
    UnitRec.BankDraft = "Y"
  Else
    UnitRec.BankDraft = "N"
  End If
  UnitRec.ESCRemitNum = "1"
  UnitRec.ESCEmplrNum = "2"

  Put FileHandle, 1, UnitRec
  Close FileHandle
  
  If Not Exist("PRDATA\PRRETIRE.DAT") Then GoTo NoRetireSaved

  txtUpdate.Text = "Updating Retirement Records"
  DoEvents

  RetHandle = FreeFile
  RetRecLen = Len(RetRec)
  Open "PRdata\PRRETIRE.DAT" For Random Shared As RetHandle Len = RetRecLen
  For cnt = 1 To 6
    Get RetHandle, cnt, RetRec
    If RetRec.TYPEM1 < 0 Then RetRec.TYPEM1 = 0
    If RetRec.TYPEWH1 < 0 Then RetRec.TYPEWH1 = 0
    If RetRec.TYPEM1 = 0 And RetRec.TYPEWH1 = 0 Then 'added 8/23
      RetRec.TYPEOT1 = "" 'added 8/23
      RetRec.TYPETD1 = "" 'added 8/23
    End If
    'commented out on 8/23
'    If QPTrim$(RetRec.TYPEOT1) = "" Then RetRec.TYPEOT1 = "N"
'    If QPTrim$(RetRec.TYPETD1) = "" Then RetRec.TYPETD1 = "N"
    Put RetHandle, cnt, RetRec
  Next cnt

NoRetireSaved:
  txtUpdate.Text = "Updating Transaction History"
  DoEvents

  OldTransHistFileHandle = FreeFile
  Open "prdata\" + "PRTRANSH.DAT" For Random Shared As OldTransHistFileHandle Len = OldTransHistRecLen

  OLDTransHistRecCnt = LOF(OldTransHistFileHandle) / Len(OldTransHistFileRec)

  ReDim TempTActive(1 To OLDTransHistRecCnt) As Integer, TempPrevTransRec(1 To OLDTransHistRecCnt) As Integer
  ReDim TempEmpPin(1 To OLDTransHistRecCnt) As Integer, TempPaySFlag$(1 To OLDTransHistRecCnt), TempCheckNum(1 To OLDTransHistRecCnt) As Long
  ReDim TempPayPdStart(1 To OLDTransHistRecCnt) As Integer, TempPayPdEnd(1 To OLDTransHistRecCnt) As Integer
  ReDim TempCheckDate(1 To OLDTransHistRecCnt) As Integer, TempPostDate(1 To OLDTransHistRecCnt) As Integer
  ReDim TempPayType$(1 To OLDTransHistRecCnt), TempBaseRate(1 To OLDTransHistRecCnt) As Double, TempOTRate(1 To OLDTransHistRecCnt) As Double
  ReDim TempRegHrsWork(1 To OLDTransHistRecCnt) As Double, TempVacUsed(1 To OLDTransHistRecCnt) As Double
  ReDim TempSickUsed(1 To OLDTransHistRecCnt) As Double, TempCompUsed(1 To OLDTransHistRecCnt) As Double
  ReDim TempHOLHOURS(1 To OLDTransHistRecCnt) As Double, TempPerHours(1 To OLDTransHistRecCnt) As Double
  ReDim TempRegHrsPaid(1 To OLDTransHistRecCnt) As Double, TempOTHours(1 To OLDTransHistRecCnt) As Double
  ReDim TempOTHrsPaid(1 To OLDTransHistRecCnt) As Double, TempOT2COMP(1 To OLDTransHistRecCnt) As Double
  ReDim TempTDist(1 To OLDTransHistRecCnt, 1 To 8) As TransWageDistType
  ReDim TempTotRegWage(1 To OLDTransHistRecCnt) As Double, TempTotOTWage(1 To OLDTransHistRecCnt) As Double
  ReDim TempGrossWage(1 To OLDTransHistRecCnt) As Double, TempEAmt(1 To OLDTransHistRecCnt, 1 To 3) As Double
  ReDim TempEDist(1 To OLDTransHistRecCnt, 1 To 6) As TransEarnDistType
  ReDim TempTotAdditEarn(1 To OLDTransHistRecCnt) As Double, TempGrossPay(1 To OLDTransHistRecCnt) As Double
  ReDim TempSocGrossPay(1 To OLDTransHistRecCnt) As Double, TempMedGrossPay(1 To OLDTransHistRecCnt) As Double
  ReDim TempFedGrossPay(1 To OLDTransHistRecCnt) As Double, TempStaGrossPay(1 To OLDTransHistRecCnt) As Double
  ReDim TempSocTaxAmt(1 To OLDTransHistRecCnt) As Double, TempMedTaxAmt(1 To OLDTransHistRecCnt) As Double
  ReDim TempFedTaxAmt(1 To OLDTransHistRecCnt) As Double, TempStaTaxAmt(1 To OLDTransHistRecCnt) As Double
  ReDim TempTotTaxAmt(1 To OLDTransHistRecCnt) As Double, TempRetireAmt(1 To OLDTransHistRecCnt) As Double
  ReDim TempDAmt(1 To OLDTransHistRecCnt, 1 To 12) As Double, TempTotDedAmt(1 To OLDTransHistRecCnt) As Double
  ReDim TempEICAmt(1 To OLDTransHistRecCnt) As Double, TempNetPay(1 To OLDTransHistRecCnt) As Double
  ReDim TempPeriodHistRec(1 To OLDTransHistRecCnt) As Integer, TempMatchRetAmt(1 To OLDTransHistRecCnt) As Double
  ReDim TempMatchSocAmt(1 To OLDTransHistRecCnt) As Double, TempMatchMedAmt(1 To OLDTransHistRecCnt) As Double
  ReDim TempRetGrossPay(1 To OLDTransHistRecCnt) As Double, TempTaxFring(1 To OLDTransHistRecCnt) As Double
  ReDim TempPad1$(1 To OLDTransHistRecCnt)

  For x = 1 To OLDTransHistRecCnt
    Get OldTransHistFileHandle, x, OldTransHistFileRec
    TempTActive(x) = OldTransHistFileRec.TActive
    TempPrevTransRec(x) = OldTransHistFileRec.PrevTransRec
    TempEmpPin(x) = OldTransHistFileRec.EmpPin
    TempPaySFlag(x) = OldTransHistFileRec.PaySFlag
    If OldTransHistFileRec.NetPay < 0 Then
      If OldTransHistFileRec.CheckNum > 0 Then
        OldTransHistFileRec.CheckNum = -OldTransHistFileRec.CheckNum
      End If
    End If
    TempCheckNum(x) = OldTransHistFileRec.CheckNum
    TempPayPdStart(x) = OldTransHistFileRec.PayPdStart
    TempPayPdEnd(x) = OldTransHistFileRec.PayPdEnd
    TempCheckDate(x) = OldTransHistFileRec.CheckDate
    TempPostDate(x) = OldTransHistFileRec.PostDate
    TempPayType(x) = OldTransHistFileRec.PayType
    TempBaseRate(x) = OldRound#(OldTransHistFileRec.BaseRate)
    TempOTRate(x) = OldRound#(OldTransHistFileRec.OTRate)
    If OldTransHistFileRec.RegHrsWork > 0 And OldTransHistFileRec.RegHrsWork < 1 Then
      OldTransHistFileRec.RegHrsWork = 0
    End If
    TempRegHrsWork(x) = OldRound#(OldTransHistFileRec.RegHrsWork)
    TempVacUsed(x) = OldRound#(OldTransHistFileRec.VacUsed)
    TempSickUsed(x) = OldRound#(OldTransHistFileRec.SickUsed)
    TempCompUsed(x) = OldRound#(OldTransHistFileRec.CompUsed)
    TempHOLHOURS(x) = OldRound#(OldTransHistFileRec.HOLHOURS)
    TempPerHours(x) = OldRound#(OldTransHistFileRec.PerHours)
    TempRegHrsPaid(x) = OldRound#(OldTransHistFileRec.RegHrsPaid)
    TempOTHours(x) = OldRound#(OldTransHistFileRec.OTHours)
    TempOTHrsPaid(x) = OldRound#(OldTransHistFileRec.OTHrsPaid)
    TempOT2COMP(x) = OldRound#(OldTransHistFileRec.OT2COMP)
    For Y = 1 To 8
      OldTransHistFileRec.TDist(Y).DOHrs = OldRound#(OldTransHistFileRec.TDist(Y).DOHrs)
      OldTransHistFileRec.TDist(Y).DOWage = OldRound#(OldTransHistFileRec.TDist(Y).DOWage)
      OldTransHistFileRec.TDist(Y).DPct = OldRound#(OldTransHistFileRec.TDist(Y).DPct)
      OldTransHistFileRec.TDist(Y).DRHrs = OldRound#(OldTransHistFileRec.TDist(Y).DRHrs)
      OldTransHistFileRec.TDist(Y).DRWage = OldRound#(OldTransHistFileRec.TDist(Y).DRWage)
      TempTDist(x, Y) = OldTransHistFileRec.TDist(Y)
    Next Y
    TempTotRegWage(x) = OldRound#(OldTransHistFileRec.TotRegWage)
    TempTotOTWage(x) = OldRound#(OldTransHistFileRec.TotOTWage)
    TempGrossWage(x) = OldRound#(OldTransHistFileRec.GrossWage)
    For Y = 1 To 3
      TempEAmt(x, Y) = OldRound#(OldTransHistFileRec.EAmt(Y))
    Next Y
    For Y = 1 To 6
      OldTransHistFileRec.EDist(Y).EAmt = OldRound#(OldTransHistFileRec.EDist(Y).EAmt)
      TempEDist(x, Y) = OldTransHistFileRec.EDist(Y)
    Next Y
    TempTotAdditEarn(x) = OldRound#(OldTransHistFileRec.TotAdditEarn)
    TempGrossPay(x) = OldRound#(OldTransHistFileRec.GrossPay)
    TempSocGrossPay(x) = OldRound#(OldTransHistFileRec.SocGrossPay)
    TempMedGrossPay(x) = OldRound#(OldTransHistFileRec.MedGrossPay)
    TempFedGrossPay(x) = OldRound#(OldTransHistFileRec.FedGrossPay)
    TempStaGrossPay(x) = OldRound#(OldTransHistFileRec.StaGrossPay)
    TempSocTaxAmt(x) = OldRound#(OldTransHistFileRec.SocTaxAmt)
    TempMedTaxAmt(x) = OldRound#(OldTransHistFileRec.MedTaxAmt)
    TempFedTaxAmt(x) = OldRound#(OldTransHistFileRec.FedTaxAmt)
    TempStaTaxAmt(x) = OldRound#(OldTransHistFileRec.StaTaxAmt)
    TempTotTaxAmt(x) = OldRound#(OldTransHistFileRec.TotTaxAmt)
    TempRetireAmt(x) = OldTransHistFileRec.RetireAmt
    For Y = 1 To 12
      TempDAmt(x, Y) = OldRound#(OldTransHistFileRec.DAmt(Y))
    Next Y
    TempTotDedAmt(x) = OldRound#(OldTransHistFileRec.TotDedAmt)
    TempEICAmt(x) = OldRound#(OldTransHistFileRec.EICAmt)
    TempNetPay(x) = OldRound#(OldTransHistFileRec.NetPay)
    TempPeriodHistRec(x) = OldTransHistFileRec.PeriodHistRec
    TempMatchRetAmt(x) = OldRound#(OldTransHistFileRec.MatchRetAmt)
    TempMatchSocAmt(x) = OldRound#(OldTransHistFileRec.MatchSocAmt)
    TempMatchMedAmt(x) = OldRound#(OldTransHistFileRec.MatchMedAmt)
    TempRetGrossPay(x) = OldRound#(OldTransHistFileRec.RetGrossPay)
    TempTaxFring(x) = OldRound#(OldTransHistFileRec.TaxFring)
  Next x

  Close OldTransHistFileHandle

  NewTransHistRecLen = Len(NewTransHistFileRec)
  NewTransHistFileHandle = FreeFile
  Open "prdata\" + "PRTRANSH.DAT" For Random Shared As NewTransHistFileHandle Len = NewTransHistRecLen

  For x = 1 To OLDTransHistRecCnt
    NewTransHistFileRec.TActive = TempTActive(x)
    NewTransHistFileRec.PrevTransRec = TempPrevTransRec(x)
    NewTransHistFileRec.EmpPin = TempEmpPin(x)
    NewTransHistFileRec.PaySFlag = TempPaySFlag(x)
    NewTransHistFileRec.CheckNum = TempCheckNum(x)
    NewTransHistFileRec.PayPdStart = TempPayPdStart(x)
    NewTransHistFileRec.PayPdEnd = TempPayPdEnd(x)
    NewTransHistFileRec.CheckDate = TempCheckDate(x)
    NewTransHistFileRec.PostDate = TempPostDate(x)
    NewTransHistFileRec.PayType = TempPayType(x)
    NewTransHistFileRec.BaseRate = TempBaseRate(x)
    NewTransHistFileRec.OTRate = TempOTRate(x)
    NewTransHistFileRec.RegHrsWork = TempRegHrsWork(x)
    NewTransHistFileRec.VacUsed = TempVacUsed(x)
    NewTransHistFileRec.SickUsed = TempSickUsed(x)
    NewTransHistFileRec.CompUsed = TempCompUsed(x)
    NewTransHistFileRec.HOLHOURS = TempHOLHOURS(x)
    NewTransHistFileRec.PerHours = TempPerHours(x)
    NewTransHistFileRec.RegHrsPaid = TempRegHrsPaid(x)
    NewTransHistFileRec.OTHours = TempOTHours(x)
    NewTransHistFileRec.OTHrsPaid = TempOTHrsPaid(x)
    NewTransHistFileRec.OT2COMP = TempOT2COMP(x)
    For Y = 1 To 8
      NewTransHistFileRec.TDist(Y) = TempTDist(x, Y)
    Next Y
    NewTransHistFileRec.TotRegWage = TempTotRegWage(x)
    NewTransHistFileRec.TotOTWage = TempTotOTWage(x)
    NewTransHistFileRec.GrossWage = TempGrossWage(x)
    For Y = 1 To 3
      NewTransHistFileRec.EAmt(Y) = TempEAmt(x, Y)
    Next Y
    For Y = 1 To 6
      NewTransHistFileRec.EDist(Y) = TempEDist(x, Y)
    Next Y
    NewTransHistFileRec.TotAdditEarn = TempTotAdditEarn(x)
    NewTransHistFileRec.GrossPay = TempGrossPay(x)
    NewTransHistFileRec.SocGrossPay = TempSocGrossPay(x)
    NewTransHistFileRec.MedGrossPay = TempMedGrossPay(x)
    NewTransHistFileRec.FedGrossPay = TempFedGrossPay(x)
    NewTransHistFileRec.StaGrossPay = TempStaGrossPay(x)
    NewTransHistFileRec.SocTaxAmt = TempSocTaxAmt(x)
    NewTransHistFileRec.MedTaxAmt = TempMedTaxAmt(x)
    NewTransHistFileRec.FedTaxAmt = TempFedTaxAmt(x)
    NewTransHistFileRec.StaTaxAmt = TempStaTaxAmt(x)
    NewTransHistFileRec.TotTaxAmt = TempTotTaxAmt(x)
    NewTransHistFileRec.RetireAmt = TempRetireAmt(x)
    For Y = 1 To 12
      NewTransHistFileRec.DAmt(Y) = TempDAmt(x, Y)
    Next Y
    For Y = 13 To 50
      NewTransHistFileRec.DAmt(Y) = 0#
    Next Y

    NewTransHistFileRec.TotDedAmt = TempTotDedAmt(x)
    NewTransHistFileRec.EICAmt = TempEICAmt(x)
    NewTransHistFileRec.NetPay = TempNetPay(x)
    NewTransHistFileRec.PeriodHistRec = TempPeriodHistRec(x)
    NewTransHistFileRec.MatchRetAmt = TempMatchRetAmt(x)
    NewTransHistFileRec.MatchSocAmt = TempMatchSocAmt(x)
    NewTransHistFileRec.MatchMedAmt = TempMatchMedAmt(x)
    NewTransHistFileRec.RetGrossPay = TempRetGrossPay(x)
    NewTransHistFileRec.TaxFring = TempTaxFring(x)
    NewTransHistFileRec.Pad1 = ""
    Put NewTransHistFileHandle, x, NewTransHistFileRec
  Next x

  Close NewTransHistFileHandle

  OldTransWorkRecLen = Len(OldTransWorkFileRec)
  OldTransWorkFileHandle = FreeFile
  Open "prdata\" + "PRTRANST.DAT" For Random Shared As OldTransWorkFileHandle Len = OldTransWorkRecLen
  OLDTransWorkRecCnt = LOF(OldTransWorkFileHandle) / Len(OldTransWorkFileRec)

  For x = 1 To OLDTransHistRecCnt
    Get OldTransWorkFileHandle, x, OldTransWorkFileRec
    TempTActive(x) = 0 'OldTransWorkFileRec.TActive
    TempPrevTransRec(x) = OldTransWorkFileRec.PrevTransRec
    TempEmpPin(x) = OldTransWorkFileRec.EmpPin
    TempPaySFlag(x) = OldTransWorkFileRec.PaySFlag
    TempCheckNum(x) = OldTransWorkFileRec.CheckNum
    TempPayPdStart(x) = OldTransWorkFileRec.PayPdStart
    TempPayPdEnd(x) = OldTransWorkFileRec.PayPdEnd
    TempCheckDate(x) = OldTransWorkFileRec.CheckDate
    TempPostDate(x) = OldTransWorkFileRec.PostDate
    TempPayType(x) = OldTransWorkFileRec.PayType
    TempBaseRate(x) = OldTransWorkFileRec.BaseRate
    TempOTRate(x) = OldTransWorkFileRec.OTRate
    TempRegHrsWork(x) = OldTransWorkFileRec.RegHrsWork
    TempVacUsed(x) = OldTransWorkFileRec.VacUsed
    TempSickUsed(x) = OldTransWorkFileRec.SickUsed
    TempCompUsed(x) = OldTransWorkFileRec.CompUsed
    TempHOLHOURS(x) = OldTransWorkFileRec.HOLHOURS
    TempPerHours(x) = OldTransWorkFileRec.PerHours
    TempRegHrsPaid(x) = OldTransWorkFileRec.RegHrsPaid
    TempOTHours(x) = OldTransWorkFileRec.OTHours
    TempOTHrsPaid(x) = OldTransWorkFileRec.OTHrsPaid
    TempOT2COMP(x) = OldTransWorkFileRec.OT2COMP
    For Y = 1 To 8
      TempTDist(x, Y) = OldTransWorkFileRec.TDist(Y)
    Next Y
    TempTotRegWage(x) = OldTransWorkFileRec.TotRegWage
    TempTotOTWage(x) = OldTransWorkFileRec.TotOTWage
    TempGrossWage(x) = OldTransWorkFileRec.GrossWage
    For Y = 1 To 3
      TempEAmt(x, Y) = OldTransWorkFileRec.EAmt(Y)
    Next Y
    For Y = 1 To 6
      TempEDist(x, Y) = OldTransWorkFileRec.EDist(Y)
    Next Y
    TempTotAdditEarn(x) = OldTransWorkFileRec.TotAdditEarn
    TempGrossPay(x) = OldTransWorkFileRec.GrossPay
    TempSocGrossPay(x) = OldTransWorkFileRec.SocGrossPay
    TempMedGrossPay(x) = OldTransWorkFileRec.MedGrossPay
    TempFedGrossPay(x) = OldTransWorkFileRec.FedGrossPay
    TempStaGrossPay(x) = OldTransWorkFileRec.StaGrossPay
    TempSocTaxAmt(x) = OldTransWorkFileRec.SocTaxAmt
    TempMedTaxAmt(x) = OldTransWorkFileRec.MedTaxAmt
    TempFedTaxAmt(x) = OldTransWorkFileRec.FedTaxAmt
    TempStaTaxAmt(x) = OldTransWorkFileRec.StaTaxAmt
    TempTotTaxAmt(x) = OldTransWorkFileRec.TotTaxAmt
    TempRetireAmt(x) = OldTransWorkFileRec.RetireAmt
    For Y = 1 To 12
      TempDAmt(x, Y) = OldTransWorkFileRec.DAmt(Y)
    Next Y
    TempTotDedAmt(x) = OldTransWorkFileRec.TotDedAmt
    TempEICAmt(x) = OldTransWorkFileRec.EICAmt
    TempNetPay(x) = OldTransWorkFileRec.NetPay
    TempPeriodHistRec(x) = OldTransWorkFileRec.PeriodHistRec
    TempMatchRetAmt(x) = OldTransWorkFileRec.MatchRetAmt
    TempMatchSocAmt(x) = OldTransWorkFileRec.MatchSocAmt
    TempMatchMedAmt(x) = OldTransWorkFileRec.MatchMedAmt
    TempRetGrossPay(x) = OldTransWorkFileRec.RetGrossPay
    TempTaxFring(x) = OldTransWorkFileRec.TaxFring
  Next x

  Close OldTransWorkFileHandle

  NewTransWorkRecLen = Len(NewTransWorkFileRec)
  NewTransWorkFileHandle = FreeFile
  Open "prdata\" + "PRTRANST.DAT" For Random Shared As NewTransWorkFileHandle Len = NewTransWorkRecLen

  For x = 1 To OLDTransHistRecCnt
    NewTransWorkFileRec.TActive = TempTActive(x)
    NewTransWorkFileRec.PrevTransRec = TempPrevTransRec(x)
    NewTransWorkFileRec.EmpPin = TempEmpPin(x)
    NewTransWorkFileRec.PaySFlag = TempPaySFlag(x)
    NewTransWorkFileRec.CheckNum = TempCheckNum(x)
    NewTransWorkFileRec.PayPdStart = TempPayPdStart(x)
    NewTransWorkFileRec.PayPdEnd = TempPayPdEnd(x)
    NewTransWorkFileRec.CheckDate = TempCheckDate(x)
    NewTransWorkFileRec.PostDate = TempPostDate(x)
    NewTransWorkFileRec.PayType = TempPayType(x)
    NewTransWorkFileRec.BaseRate = TempBaseRate(x)
    NewTransWorkFileRec.OTRate = TempOTRate(x)
    NewTransWorkFileRec.RegHrsWork = TempRegHrsWork(x)
    NewTransWorkFileRec.VacUsed = TempVacUsed(x)
    NewTransWorkFileRec.SickUsed = TempSickUsed(x)
    NewTransWorkFileRec.CompUsed = TempCompUsed(x)
    NewTransWorkFileRec.HOLHOURS = TempHOLHOURS(x)
    NewTransWorkFileRec.PerHours = TempPerHours(x)
    NewTransWorkFileRec.RegHrsPaid = TempRegHrsPaid(x)
    NewTransWorkFileRec.OTHours = TempOTHours(x)
    NewTransWorkFileRec.OTHrsPaid = TempOTHrsPaid(x)
    NewTransWorkFileRec.OT2COMP = TempOT2COMP(x)
    For Y = 1 To 8
      NewTransWorkFileRec.TDist(Y) = TempTDist(x, Y)
    Next Y
    NewTransWorkFileRec.TotRegWage = TempTotRegWage(x)
    NewTransWorkFileRec.TotOTWage = TempTotOTWage(x)
    NewTransWorkFileRec.GrossWage = TempGrossWage(x)
    For Y = 1 To 3
      NewTransWorkFileRec.EAmt(Y) = TempEAmt(x, Y)
    Next Y
    For Y = 1 To 6
      NewTransWorkFileRec.EDist(Y) = TempEDist(x, Y)
    Next Y
    NewTransWorkFileRec.TotAdditEarn = TempTotAdditEarn(x)
    NewTransWorkFileRec.GrossPay = TempGrossPay(x)
    NewTransWorkFileRec.SocGrossPay = TempSocGrossPay(x)
    NewTransWorkFileRec.MedGrossPay = TempMedGrossPay(x)
    NewTransWorkFileRec.FedGrossPay = TempFedGrossPay(x)
    NewTransWorkFileRec.StaGrossPay = TempStaGrossPay(x)
    NewTransWorkFileRec.SocTaxAmt = TempSocTaxAmt(x)
    NewTransWorkFileRec.MedTaxAmt = TempMedTaxAmt(x)
    NewTransWorkFileRec.FedTaxAmt = TempFedTaxAmt(x)
    NewTransWorkFileRec.StaTaxAmt = TempStaTaxAmt(x)
    NewTransWorkFileRec.TotTaxAmt = TempTotTaxAmt(x)
    NewTransWorkFileRec.RetireAmt = TempRetireAmt(x)
    For Y = 1 To 12
      NewTransWorkFileRec.DAmt(Y) = TempDAmt(x, Y)
    Next Y

    For Y = 13 To 50
      NewTransWorkFileRec.DAmt(Y) = 0#
    Next Y

    NewTransWorkFileRec.TotDedAmt = TempTotDedAmt(x)
    NewTransWorkFileRec.EICAmt = TempEICAmt(x)
    NewTransWorkFileRec.NetPay = TempNetPay(x)
    NewTransWorkFileRec.PeriodHistRec = TempPeriodHistRec(x)
    NewTransWorkFileRec.MatchRetAmt = TempMatchRetAmt(x)
    NewTransWorkFileRec.MatchSocAmt = TempMatchSocAmt(x)
    NewTransWorkFileRec.MatchMedAmt = TempMatchMedAmt(x)
    NewTransWorkFileRec.RetGrossPay = TempRetGrossPay(x)
    NewTransWorkFileRec.TaxFring = TempTaxFring(x)
    NewTransWorkFileRec.Pad1 = ""
    Put NewTransWorkFileHandle, x, NewTransWorkFileRec
  Next x

  Close NewTransWorkFileHandle
  txtUpdate.Text = "Updating Check Records"
  DoEvents

  OldChecksHandle = FreeFile
  Open "prdata\" + "PRCHECKS.DAT" For Random Shared As OldChecksHandle Len = OldChecksRecLen
  NumOfChkRecs = LOF(OldChecksHandle) / Len(OldChkRec)

  ReDim TempCActive(1 To NumOfChkRecs) As Integer
  ReDim TempCheckNum(1 To NumOfChkRecs) As Long
  ReDim TempCheckDate(1 To NumOfChkRecs) As Integer
  ReDim TempEmpName(1 To NumOfChkRecs) As String * 33
  ReDim TempEmpNo(1 To NumOfChkRecs) As String * 10
  ReDim TempEmpSSN(1 To NumOfChkRecs) As String * 11
  ReDim TempEmpAddr1(1 To NumOfChkRecs) As String * 36
  ReDim TempEmpCity(1 To NumOfChkRecs) As String * 24
  ReDim TempEmpState(1 To NumOfChkRecs) As String * 2
  ReDim TempEmpZip(1 To NumOfChkRecs) As String * 10
  ReDim TempPayEndDate(1 To NumOfChkRecs) As Integer
  ReDim TempBaseRate(1 To NumOfChkRecs) As Double
  ReDim TempGrossPay(1 To NumOfChkRecs) As Double
  ReDim TempFedTaxAmt(1 To NumOfChkRecs) As Double
  ReDim TempStaTaxAmt(1 To NumOfChkRecs) As Double
  ReDim TempMedTaxAmt(1 To NumOfChkRecs) As Double
  ReDim TempSocTaxAmt(1 To NumOfChkRecs) As Double
  ReDim TempTotDedAmt(1 To NumOfChkRecs) As Double
  ReDim TempRetireAmt(1 To NumOfChkRecs) As Double
  ReDim TempNetPay(1 To NumOfChkRecs) As Double
  ReDim TempYTDGrossPay(1 To NumOfChkRecs) As Double
  ReDim TempYTDFederal(1 To NumOfChkRecs) As Double
  ReDim TempYTDState(1 To NumOfChkRecs) As Double
  ReDim TempYTDSocial(1 To NumOfChkRecs) As Double
  ReDim TempYTDMedicare(1 To NumOfChkRecs) As Double
  ReDim TempYTDTotDed(1 To NumOfChkRecs) As Double
  ReDim TempYTDNetPay(1 To NumOfChkRecs) As Double
  ReDim TempYTDRetire(1 To NumOfChkRecs) As Double
  ReDim TempVactBal(1 To NumOfChkRecs) As Double
  ReDim TempSickBal(1 To NumOfChkRecs) As Double
  ReDim TempCompBal(1 To NumOfChkRecs) As Double
  ReDim TempCompEarn(1 To NumOfChkRecs) As Double
  ReDim TempRegHrsWork(1 To NumOfChkRecs) As Double
  ReDim TempOTHrsPaid(1 To NumOfChkRecs) As Double
  ReDim TempTotRegWage(1 To NumOfChkRecs) As Double
  ReDim TempVacUsed(1 To NumOfChkRecs) As Double
  ReDim TempSickUsed(1 To NumOfChkRecs) As Double
  ReDim TempCompUsed(1 To NumOfChkRecs) As Double
  ReDim TempHolUsed(1 To NumOfChkRecs) As Double
  ReDim TempPerUsed(1 To NumOfChkRecs) As Double
  ReDim TempRegHrsPaid(1 To NumOfChkRecs) As Double
  ReDim TempTotOTWage(1 To NumOfChkRecs) As Double
  ReDim TempAEarn(1 To NumOfChkRecs, 1 To 3) As PRDEDType
  ReDim TempTotAdditEarn(1 To NumOfChkRecs) As Double
  ReDim TempEICAmt(1 To NumOfChkRecs) As Double
  ReDim TempTaxFring(1 To NumOfChkRecs) As Double
  ReDim TempCDED(1 To NumOfChkRecs, 1 To 12) As PRDEDType
  ReDim TempDDFlag(1 To NumOfChkRecs) As Integer


  For x = 1 To NumOfChkRecs
    Get OldChecksHandle, x, OldChkRec
    TempCActive(x) = OldChkRec.CActive
    TempCheckNum(x) = OldChkRec.CheckNum
    TempCheckDate(x) = OldChkRec.CheckDate
    TempEmpName(x) = OldChkRec.EmpName
    TempEmpNo(x) = OldChkRec.EmpNo
    TempEmpSSN(x) = OldChkRec.EmpSSN
    TempEmpAddr1(x) = OldChkRec.EmpAddr1
    TempEmpCity(x) = OldChkRec.EmpCity
    TempEmpState(x) = OldChkRec.EmpState
    TempEmpZip(x) = OldChkRec.EmpZip
    TempPayEndDate(x) = OldChkRec.PayEndDate
    TempBaseRate(x) = OldChkRec.BaseRate
    TempGrossPay(x) = OldChkRec.GrossPay
    TempFedTaxAmt(x) = OldChkRec.FedTaxAmt
    TempStaTaxAmt(x) = OldChkRec.StaTaxAmt
    TempMedTaxAmt(x) = OldChkRec.MedTaxAmt
    TempSocTaxAmt(x) = OldChkRec.SocTaxAmt
    TempTotDedAmt(x) = OldChkRec.TotDedAmt
    TempRetireAmt(x) = OldChkRec.RetireAmt
    TempNetPay(x) = OldChkRec.NetPay
    TempYTDGrossPay(x) = OldChkRec.YTDGrossPay
    TempYTDFederal(x) = OldChkRec.YTDFederal
    TempYTDState(x) = OldChkRec.YTDState
    TempYTDSocial(x) = OldChkRec.YTDSocial
    TempYTDMedicare(x) = OldChkRec.YTDMedicare
    TempYTDTotDed(x) = OldChkRec.YTDTotDed
    TempYTDNetPay(x) = OldChkRec.YTDNetPay
    TempYTDRetire(x) = OldChkRec.YTDRetire
    TempVactBal(x) = OldChkRec.VactBal
    TempSickBal(x) = OldChkRec.SickBal
    TempCompBal(x) = OldChkRec.CompBal
    TempCompEarn(x) = OldChkRec.CompEarn
    TempRegHrsWork(x) = OldChkRec.RegHrsWork
    TempOTHrsPaid(x) = OldChkRec.OTHrsPaid
    TempTotRegWage(x) = OldChkRec.TotRegWage
    TempVacUsed(x) = OldChkRec.VacUsed
    TempSickUsed(x) = OldChkRec.SickUsed
    TempCompUsed(x) = OldChkRec.CompUsed
    TempHolUsed(x) = OldChkRec.HolUsed
    TempPerUsed(x) = OldChkRec.PerUsed
    TempRegHrsPaid(x) = OldChkRec.RegHrsPaid
    TempTotOTWage(x) = OldChkRec.TotOTWage
    For Y = 1 To 3
      TempAEarn(x, Y).DAmt = OldChkRec.AEarn(Y).DAmt
      TempAEarn(x, Y).DCode = OldChkRec.AEarn(Y).DCode
      TempAEarn(x, Y).YTDDAmt = OldChkRec.AEarn(Y).YTDDAmt
    Next Y
    TempTotAdditEarn(x) = OldChkRec.TotAdditEarn
    TempEICAmt(x) = OldChkRec.EICAmt
    TempTaxFring(x) = OldChkRec.TaxFring
    For Y = 1 To 12
      TempCDED(x, Y).DAmt = OldChkRec.CDED(Y).DAmt
      TempCDED(x, Y).DCode = OldChkRec.CDED(Y).DCode
      TempCDED(x, Y).YTDDAmt = OldChkRec.CDED(Y).YTDDAmt
    Next Y
    TempDDFlag(x) = OldChkRec.DDFlag
  Next x

  Close OldChecksHandle

  NewChecksHandle = FreeFile
  Open "prdata\" + "PRCHECKS.DAT" For Random Shared As NewChecksHandle Len = NewChecksRecLen

  For x = 1 To NumOfChkRecs
    Get NewChecksHandle, x, NewChkRec
    NewChkRec.CActive = TempCActive(x)
    NewChkRec.CheckNum = TempCheckNum(x)
    NewChkRec.CheckDate = TempCheckDate(x)
    NewChkRec.EmpName = TempEmpName(x)
    NewChkRec.EmpNo = TempEmpNo(x)
    NewChkRec.EmpSSN = TempEmpSSN(x)
    NewChkRec.EmpAddr1 = TempEmpAddr1(x)
    NewChkRec.EmpCity = TempEmpCity(x)
    NewChkRec.EmpState = TempEmpState(x)
    NewChkRec.EmpZip = TempEmpZip(x)
    NewChkRec.PayEndDate = TempPayEndDate(x)
    NewChkRec.BaseRate = TempBaseRate(x)
    NewChkRec.GrossPay = TempGrossPay(x)
    NewChkRec.FedTaxAmt = TempFedTaxAmt(x)
    NewChkRec.StaTaxAmt = TempStaTaxAmt(x)
    NewChkRec.MedTaxAmt = TempMedTaxAmt(x)
    NewChkRec.SocTaxAmt = TempSocTaxAmt(x)
    NewChkRec.TotDedAmt = TempTotDedAmt(x)
    NewChkRec.RetireAmt = TempRetireAmt(x)
    NewChkRec.NetPay = TempNetPay(x)
    NewChkRec.YTDGrossPay = TempYTDGrossPay(x)
    NewChkRec.YTDFederal = TempYTDFederal(x)
    NewChkRec.YTDState = TempYTDState(x)
    NewChkRec.YTDSocial = TempYTDSocial(x)
    NewChkRec.YTDMedicare = TempYTDMedicare(x)
    NewChkRec.YTDTotDed = TempYTDTotDed(x)
    NewChkRec.YTDNetPay = TempYTDNetPay(x)
    NewChkRec.YTDRetire = TempYTDRetire(x)
    NewChkRec.VactBal = TempVactBal(x)
    NewChkRec.SickBal = TempSickBal(x)
    NewChkRec.CompBal = TempCompBal(x)
    NewChkRec.CompEarn = TempCompEarn(x)
    NewChkRec.RegHrsWork = TempRegHrsWork(x)
    NewChkRec.OTHrsPaid = TempOTHrsPaid(x)
    NewChkRec.TotRegWage = TempTotRegWage(x)
    NewChkRec.VacUsed = TempVacUsed(x)
    NewChkRec.SickUsed = TempSickUsed(x)
    NewChkRec.CompUsed = TempCompUsed(x)
    NewChkRec.HolUsed = TempHolUsed(x)
    NewChkRec.PerUsed = TempPerUsed(x)
    NewChkRec.RegHrsPaid = TempRegHrsPaid(x)
    NewChkRec.TotOTWage = TempTotOTWage(x)
    For Y = 1 To 3
      NewChkRec.AEarn(Y).DAmt = TempAEarn(x, Y).DAmt
      NewChkRec.AEarn(Y).DCode = TempAEarn(x, Y).DCode
      NewChkRec.AEarn(Y).YTDDAmt = TempAEarn(x, Y).YTDDAmt
    Next Y
    NewChkRec.TotAdditEarn = TempTotAdditEarn(x)
    NewChkRec.EICAmt = TempEICAmt(x)
    NewChkRec.TaxFring = TempTaxFring(x)
    For Y = 1 To 12
      NewChkRec.CDED(Y).DAmt = TempCDED(x, Y).DAmt
      NewChkRec.CDED(Y).DCode = TempCDED(x, Y).DCode
      NewChkRec.CDED(Y).YTDDAmt = TempCDED(x, Y).YTDDAmt
    Next Y
    For Y = 13 To 50
      NewChkRec.CDED(Y).DAmt = 0
      NewChkRec.CDED(Y).DCode = ""
      NewChkRec.CDED(Y).YTDDAmt = 0
    Next Y
    NewChkRec.DDFlag = TempDDFlag(x)

    Put NewChecksHandle, x, NewChkRec
  Next x
  Close NewChecksHandle


  Dim TempExtrYear As Integer
  Dim TempDeds(0 To 50) As W2DedType

  OldW2Handle = FreeFile
  Open "prdata\" + "PRW2SETU.DAT" For Random Shared As OldW2Handle Len = OldW2RecLen

  Get OldW2Handle, 1, OldW2SetUp
  TempExtrYear = OldW2SetUp.ExtrYear
  For x = 0 To 12
    TempDeds(x).AMTBOX = OldW2SetUp.Deds(x).AMTBOX
    TempDeds(x).CHKDED = OldW2SetUp.Deds(x).CHKDED
    TempDeds(x).DedCode = OldW2SetUp.Deds(x).DedCode
  Next x
  For x = 13 To 50
    TempDeds(x).AMTBOX = ""
    TempDeds(x).CHKDED = ""
    TempDeds(x).DedCode = ""
  Next x

  Close OldW2Handle

  NewW2Handle = FreeFile
  Open "prdata\" + "PRW2SETU.DAT" For Random Shared As NewW2Handle Len = NewW2RecLen
  NewW2SetUp.ExtrYear = TempExtrYear
  For x = 0 To 50
    NewW2SetUp.Deds(x).AMTBOX = TempDeds(x).AMTBOX
    NewW2SetUp.Deds(x).CHKDED = TempDeds(x).CHKDED
    NewW2SetUp.Deds(x).DedCode = TempDeds(x).DedCode
  Next x

  Put NewW2Handle, 1, NewW2SetUp
  Close NewW2Handle
  
  txtUpdate.Visible = False
  cmdCancel.Enabled = True
  
  txtUpdate.Text = "UPDATE COMPLETE. PRESS OK TO GO TO PAYROLL MAIN MENU"
  cmdOK.Visible = True
  cmdConvert.Visible = False
  cmdCancel.Visible = False
End Sub
Private Sub cmdOK_Click()
  Close
  Unload frmConvert
  KillFile "prrun.opn"
  End
End Sub

Private Sub Form_Load()
  If App.PrevInstance Then
    ActivatePrevInstance
  End If
  txtUpdate.Visible = False
  cmdOK.Visible = False
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdCancel.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn
      KillFile "prrun.opn"
      End
    End If
  End If
End Sub

