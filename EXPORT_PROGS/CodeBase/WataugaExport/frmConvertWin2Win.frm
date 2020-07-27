VERSION 5.00
Begin VB.Form frmConvertWin2Win 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Windows to WIndows"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   3252
      TabIndex        =   3
      Top             =   6090
      Width           =   1932
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
      Left            =   6228
      TabIndex        =   2
      Top             =   6090
      Width           =   1932
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
      Left            =   3864
      MousePointer    =   11  'Hourglass
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1902
      Width           =   3948
   End
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
      Left            =   4440
      TabIndex        =   0
      Top             =   5550
      Width           =   2556
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2316
      Left            =   2508
      Top             =   2634
      Width           =   6636
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   " This procedure updates any WINDOWS payroll data that is not compatible with the latest WINDOWS payroll program."
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
      Left            =   2796
      TabIndex        =   4
      Top             =   3066
      Width           =   6060
   End
End
Attribute VB_Name = "frmConvertWin2Win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload frmConvertWin2Win
  DoEvents
  
  End
End Sub

Private Sub cmdConvert_Click()
  Dim OldWinEmpData2FileHandle As Integer
  Dim cnt As Integer
  Dim OldWinEmpData2FileRec As OldWinEmpData2Type
  Dim NewEmpData2FileRec As EmpData2Type
  Dim OldWinEmpData2RecLen As Integer
  Dim NewEmpData2RecLen As Integer
  Dim NumOfRecs As Integer
  Dim Cnt2 As Integer
  Dim OldHandle As Integer, NewHandle As Integer
  Dim OldWinFileHandle As Integer
  Dim OldWinUnitRec As OldWinUnitFileRecType
  Dim OldWinUnitRecLen As Integer
  Dim FileHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim UnitRecLen As Integer
  Dim Idx As Integer
  Dim Y As Integer
  Dim x As Integer, TempDedRec(1 To 12) As Integer
  Dim DraftLen As Integer
  Dim ESCFileName$
  Dim ESCHandle As Integer
  Dim RemitNumb$, EmplrAcct$
  
  If Exist("PRDATA\PRESCCFG.DAT") Then
    ESCHandle = FreeFile
    Open "PRDATA\PRESCCFG.DAT" For Input As ESCHandle
    Line Input #ESCHandle, RemitNumb$
    Line Input #ESCHandle, EmplrAcct$
    Close ESCHandle
  Else
    RemitNumb$ = "0"
    EmplrAcct$ = "0"
  End If
  
  If CheckForAllFiles = False Then
    Unload frmConvertWin2Win
    DoEvents
    frmWin2WinMissingFiles.Show
    Exit Sub
  End If
  
  cmdConvert.MousePointer = vbHourglass
  
  OldWinEmpData2RecLen = Len(OldWinEmpData2FileRec)
  NewEmpData2RecLen = Len(NewEmpData2FileRec)
  
  cmdCancel.Enabled = False
  
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
    If TLeaveTbl(cnt) < 0 Then TLeaveTbl(cnt) = 0
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
    NewEmpData2FileRec.YN401K = "Y"
    NewEmpData2FileRec.PrimeDept = "1"
    NewEmpData2FileRec.HomePhone = "(000)-000-0000"
    NewEmpData2FileRec.EmrgncyCntctName = "Nothing Saved"
    NewEmpData2FileRec.EmrgncyCntctPhnNum = "(000)-000-0000"
    NewEmpData2FileRec.EmrgncyCntctRelation = "Nothing Saved"
    
    Put NewHandle, cnt, NewEmpData2FileRec
  Next cnt
  Close NewHandle
'*************************************************************
  
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
  UnitRec.ESCRemitNum = RemitNumb$
  UnitRec.ESCEmplrNum = EmplrAcct$

  Put FileHandle, 1, UnitRec
  Close FileHandle
  
  
  txtUpdate.Visible = False
  cmdCancel.Enabled = True
  
  txtUpdate.Text = "UPDATE COMPLETE. PRESS OK TO GO TO PAYROLL MAIN MENU"
  cmdOK.Visible = True
  cmdConvert.Visible = False
  cmdCancel.Visible = False
End Sub
Private Sub cmdOK_Click()
  Close
  Unload frmConvertWin2Win
  
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
      End
    End If
  End If
End Sub


