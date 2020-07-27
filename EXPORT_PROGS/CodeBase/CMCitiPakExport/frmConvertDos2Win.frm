VERSION 5.00
Begin VB.Form frmConvertDos2Win 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Dos to Windows"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmConvertDos2Win.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkNotFall04 
      BackColor       =   &H0080FFFF&
      Caption         =   "BEFORE VERSION 2.05 UPDATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7200
      TabIndex        =   9
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CheckBox chkFall04Update 
      BackColor       =   &H0080FFFF&
      Caption         =   "VERSION 2.05 UPDATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      TabIndex        =   8
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   3252
      TabIndex        =   3
      Top             =   7230
      Width           =   1932
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "BEGIN UPDATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   6228
      TabIndex        =   2
      Top             =   7230
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
      Top             =   1668
      Width           =   3948
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "   CONVERSION        COMPLETE        PRESS TO EXIT"
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
      Top             =   6084
      Width           =   2556
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmConvertDos2Win.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   924
      Left            =   1524
      TabIndex        =   7
      Top             =   528
      Width           =   8604
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "This version automatically adds the .LMT401YN field to the Unit record."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   684
      Left            =   2736
      TabIndex        =   6
      Top             =   4800
      Width           =   6204
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "This version automatically adds the .Less401k field = False to the transaction history records."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   684
      Left            =   2736
      TabIndex        =   5
      Top             =   3936
      Width           =   6204
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3276
      Left            =   2508
      Top             =   2400
      Width           =   6636
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   " This procedure updates any DOS payroll data that is not compatible with the WINDOWS payroll program."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1260
      Left            =   2736
      TabIndex        =   4
      Top             =   2544
      Width           =   6204
   End
End
Attribute VB_Name = "frmConvertDos2Win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload frmConvertDos2Win
  DoEvents
  End
End Sub

Private Sub Convert()
  Dim OldChkRec As DosPRCheckRecType
  Dim NewChkRec As PRCheckRecType
  Dim OldChecksHandle As Integer
  Dim NewChecksHandle As Integer
  Dim OldChecksRecLen As Integer
  Dim NewChecksRecLen As Integer
  Dim NumOfChkRecs As Long
  Dim OldPrnSet As DosPRNSetupRecType
  Dim NewPrnSet As PRNSetupRecType
  Dim OldEmpData2FileHandle As Integer
  Dim cnt As Integer
  Dim OldEmpData2FileRec As DosEmpData2Type
  Dim NewEmpData2FileRec As EmpData2Type
  Dim OldEmpData2RecLen As Integer
  Dim NewEmpData2RecLen As Integer
  Dim OldEmpData3FileRec As DosEmpData3Type
  Dim NewEmpData3FileRec As EmpData3Type
  Dim OldEmpData3RecLen As Integer
  Dim NewEmpData3RecLen As Integer
  Dim NumOfRecs As Integer
  Dim Old3Handle As Integer, New3Handle As Integer
  Dim Cnt2 As Integer
  Dim OldPrnLen As Integer, NewPrnLen As Integer
  Dim OldHandle As Integer, NewHandle As Integer
  Dim TempPrnLen As Integer
  Dim EICHandle As Integer
  Dim EICRec As EICRecType
  Dim EICFileRec As EICRecType
  Dim EICRecLen As Integer
  Dim DosFileHandle As Integer
  Dim DosUnitRec As DosUnitFileRecType
  Dim DosUnitRecLen As Integer
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
  Dim y As Integer
  Dim OldTransHistFileHandle As Integer
  Dim OldTransWorkFileHandle As Integer
  Dim NewTransWorkFileHandle As Integer
  Dim RetHandle As Integer
  Dim RetRec As RetireRecType
  Dim RetRecLen As Integer, Nextx As Integer
  Dim OLDTransHistRecCnt As Integer
  Dim OLDTransWorkRecCnt As Integer
  Dim X As Integer, TempDedRec(1 To 12) As Integer
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
  Dim ESCFileName$
  Dim ESCHandle As Integer
  Dim RemitNumb$, EmplrAcct$
  Dim OldLeaveRec As OldLeaveRecType
  Dim OldLvHandle As Integer
  Dim LeaveRec As LeaveRecType
  Dim NewLvHandle As Integer
  Dim NumOfLvRecs As Integer
  
'  On Error GoTo ErrorHandler
  InFileNames(1) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames, 1) = False Then
    Exit Sub
  End If
  
  OpenOldLeaveFileName OldLvHandle
  NumOfLvRecs = LOF(OldLvHandle) / Len(OldLeaveRec)
  
  ReDim TempVacMax(1 To NumOfLvRecs) As Double
  ReDim TempVEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  ReDim TempSICKMAX(1 To NumOfLvRecs) As Double
  ReDim TempSEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  
  If Exist("PRDATA\PRSYS.DAT") Then
    If FileLen("PRDATA\PRSYS.DAT") <> 337 Then
      MsgBox "Conversion has already taken place for this customer. Please convert leave table files with the leave table specific conversion program."
      Close
      Exit Sub
    End If
  End If

  For X = 1 To NumOfLvRecs
    Get OldLvHandle, X, OldLeaveRec
    TempVacMax(X) = OldLeaveRec.VacMax
    TempSICKMAX(X) = OldLeaveRec.SICKMAX
    For y = 1 To 20
      TempVEntry(X, y).EARN = OldLeaveRec.VEntry(y).EARN
      TempSEntry(X, y).EARN = OldLeaveRec.SEntry(y).EARN
      TempVEntry(X, y).YEARS = OldLeaveRec.VEntry(y).YEARS
      TempSEntry(X, y).YEARS = OldLeaveRec.SEntry(y).YEARS
    Next y
  Next X
  Close OldLvHandle
  
  OpenLeaveFileName NewLvHandle
  For X = 1 To NumOfLvRecs
    LeaveRec.VacMax = TempVacMax(X)
    LeaveRec.SICKMAX = TempSICKMAX(X)
    LeaveRec.HolMax = 0
    LeaveRec.PerMax = 0
    For y = 1 To 20
      LeaveRec.VEntry(y).EARN = TempVEntry(X, y).EARN
      LeaveRec.SEntry(y).EARN = TempSEntry(X, y).EARN
      LeaveRec.HEntry(y).EARN = 0
      LeaveRec.PEntry(y).EARN = 0
      LeaveRec.VEntry(y).YEARS = TempVEntry(X, y).YEARS
      LeaveRec.SEntry(y).YEARS = TempSEntry(X, y).YEARS
      LeaveRec.HEntry(y).YEARS = 0
      LeaveRec.PEntry(y).YEARS = 0
    Next y
    Put NewLvHandle, X, LeaveRec
  Next X
  Close NewLvHandle
  
SkipLeaveConversion:
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
    Unload frmConvertDos2Win
    DoEvents
    frmDos2WinFilesMissing.Show
    Exit Sub
  End If
  
  cmdConvert.MousePointer = vbHourglass
  
  OldPrnLen = Len(OldPrnSet)
  NewPrnLen = Len(NewPrnSet)
    
  OldTransHistRecLen = Len(OldTransHistFileRec)
  OldEmpData2RecLen = Len(OldEmpData2FileRec)
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
  
  Open "prdata\" + "PREMP2.DAT" For Random Shared As OldHandle Len = OldEmpData2RecLen
  
  NumOfRecs = LOF(OldHandle) / Len(OldEmpData2FileRec)
  
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
  ReDim TempDed(1 To NumOfRecs, 1 To 12) As EmpDedType
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
    Get OldHandle, cnt, OldEmpData2FileRec
    If QPTrim$(OldEmpData2FileRec.EmpNo) = "" Then
      TEmpNo(cnt) = "0"
    Else
      TEmpNo(cnt) = OldEmpData2FileRec.EmpNo
    End If
    If QPTrim$(OldEmpData2FileRec.EmpSSN) = "" Then
      TEmpSSN(cnt) = "000-00-000"
    Else
      TEmpSSN(cnt) = OldEmpData2FileRec.EmpSSN
    End If
    'Found some situations where employees were zeroed out so they had
    'a value but the program didn't recognize zero as a name
    If QPTrim$(OldEmpData2FileRec.EMPLNAME) = "" Or InStr(OldEmpData2FileRec.EMPLNAME, "0") Then
      TEMPLNAME(cnt) = "NONE"
    Else
      TEMPLNAME(cnt) = UCase(OldEmpData2FileRec.EMPLNAME)
    End If
    If QPTrim$(OldEmpData2FileRec.EMPFNAME) = "" Or InStr(OldEmpData2FileRec.EMPFNAME, "0") Then
      TEMPFNAME(cnt) = "NONE"
    Else
      TEMPFNAME(cnt) = UCase(OldEmpData2FileRec.EMPFNAME)
    End If
    
    If QPTrim$(OldEmpData2FileRec.EmpAddr1) = "" Then
      TEmpAddr1(cnt) = "None"
    Else
      TEmpAddr1(cnt) = OldEmpData2FileRec.EmpAddr1
    End If
    TEMPADDR2(cnt) = OldEmpData2FileRec.EMPADDR2
    
    If QPTrim$(OldEmpData2FileRec.EmpCity) = "" Then
      TEmpCity(cnt) = "None"
    Else
      TEmpCity(cnt) = OldEmpData2FileRec.EmpCity
    End If
    If QPTrim$(OldEmpData2FileRec.EmpState) = "" Then
      TEmpState(cnt) = "None"
    Else
      TEmpState(cnt) = OldEmpData2FileRec.EmpState
    End If
    
    If QPTrim$(OldEmpData2FileRec.EmpZip) = "" Then
      TEmpZip(cnt) = "00000"
    Else
      TEmpZip(cnt) = OldEmpData2FileRec.EmpZip
    End If
    If Val(OldEmpData2FileRec.EMPBDAY) <= -22000 Then
      TEMPBDAY(cnt) = 0
    Else
      TEMPBDAY(cnt) = OldEmpData2FileRec.EMPBDAY
    End If
    TEMPGENDR(cnt) = OldEmpData2FileRec.EMPGENDR
    TEMPRACE(cnt) = OldEmpData2FileRec.EMPRACE
    TEMPRETNO(cnt) = OldEmpData2FileRec.EMPRETNO
    TEMPRETTP(cnt) = OldEmpData2FileRec.EMPRETTP
    If QPTrim$(OldEmpData2FileRec.DRAFTCOD) = "0" Then
      OldEmpData2FileRec.DRAFTCOD = ""
    End If
    If QPTrim$(OldEmpData2FileRec.DRAFTCOD) = "" Then
      TDRAFTCOD(cnt) = ""
      TEMPDDACC(cnt) = ""
      TPRENOTED(cnt) = ""
      TBANKNAME(cnt) = ""
      TBANKLOC(cnt) = ""
      TTRANSIT(cnt) = ""
    Else
      TDRAFTCOD(cnt) = OldEmpData2FileRec.DRAFTCOD
      TEMPDDACC(cnt) = OldEmpData2FileRec.EMPDDACC
      TPRENOTED(cnt) = OldEmpData2FileRec.PRENOTED
      TBANKNAME(cnt) = OldEmpData2FileRec.BANKNAME
      TBANKLOC(cnt) = OldEmpData2FileRec.BANKLOC
      TTRANSIT(cnt) = OldEmpData2FileRec.TRANSIT
    End If
    TEMPJOB(cnt) = OldEmpData2FileRec.EMPJOB
    If QPTrim$(OldEmpData2FileRec.EMPWCCLS) = "" Then
      TEMPWCCLS(cnt) = "0"
    Else
      TEMPWCCLS(cnt) = OldEmpData2FileRec.EMPWCCLS
    End If
    
    TEMPSTATS(cnt) = OldEmpData2FileRec.EMPSTATS
    If OldEmpData2FileRec.EMPBCODE < 0 Then
      TEMPBCODE(cnt) = 0
    Else
      TEMPBCODE(cnt) = OldEmpData2FileRec.EMPBCODE
    End If
    TEMPPTYPE(cnt) = OldEmpData2FileRec.EMPPTYPE
    TEMPPFREQ(cnt) = OldEmpData2FileRec.EMPPFREQ
    If OldEmpData2FileRec.EMPPRATE < 0 Then
      TEMPPRATE(cnt) = 0
    Else
      TEMPPRATE(cnt) = OldEmpData2FileRec.EMPPRATE
    End If
    If OldEmpData2FileRec.EMPORATE < 0 Then
      TEMPORATE(cnt) = 0
    Else
      TEMPORATE(cnt) = OldEmpData2FileRec.EMPORATE
    End If
    If Val(OldEmpData2FileRec.EMPHDATE) <= -10957 Then
      TEMPHDATE(cnt) = 0
    Else
      TEMPHDATE(cnt) = OldEmpData2FileRec.EMPHDATE
    End If
    
    If OldEmpData2FileRec.EMPRDATE < 0 Then
      TEMPRDATE(cnt) = 0
    Else
      TEMPRDATE(cnt) = OldEmpData2FileRec.EMPRDATE
    End If
    
    If OldEmpData2FileRec.EMPTDATE <= -3652 Then
      TEMPTDATE(cnt) = 0
    Else
      TEMPTDATE(cnt) = OldEmpData2FileRec.EMPTDATE
    End If
    If QPTrim$(OldEmpData2FileRec.EMPFEDX) = "" Then
      TEMPFEDX(cnt) = "N"
    Else
      TEMPFEDX(cnt) = OldEmpData2FileRec.EMPFEDX
    End If
    TEMPFEDO2(cnt) = OldEmpData2FileRec.EMPFEDO2
    
    'neutralize situation where there is no value in the Fed Amt/Pct
    'field but there is a value in the Fed figure field
    If Len(QPTrim$(OldEmpData2FileRec.EMPFEDX)) = 0 And OldEmpData2FileRec.EMPFEDO1 > 0 Then
      TEMPFEDO1(cnt) = 0
      GoTo BadFedAmtPct
    End If
    If OldEmpData2FileRec.EMPFEDO1 < -0 Then
      TEMPFEDO1(cnt) = 0
    Else
      TEMPFEDO1(cnt) = OldEmpData2FileRec.EMPFEDO1
    End If
BadFedAmtPct:
    If QPTrim$(OldEmpData2FileRec.EMPFEDS) = "" Then
      TEMPFEDS(cnt) = "S"
    Else
      TEMPFEDS(cnt) = OldEmpData2FileRec.EMPFEDS
    End If
    If OldEmpData2FileRec.EMPFEDA < 0 Then
      TEMPFEDA(cnt) = 0
    Else
      TEMPFEDA(cnt) = OldEmpData2FileRec.EMPFEDA
    End If
    If OldEmpData2FileRec.EMPFEDAA < 0 Then
      TEMPFEDAA(cnt) = 0
    Else
      TEMPFEDAA(cnt) = OldEmpData2FileRec.EMPFEDAA
    End If
    
    If QPTrim$(OldEmpData2FileRec.EMPSTAX) = "" Then
      TEMPSTAX(cnt) = "N"
    Else
      TEMPSTAX(cnt) = OldEmpData2FileRec.EMPSTAX
    End If
    TEMPSTAO2(cnt) = OldEmpData2FileRec.EMPSTAO2
    'neutralize situation where there is no value in the State Amt/Pct
    'field but there is a value in the State figure field
    If Len(QPTrim$(OldEmpData2FileRec.EMPSTAX)) = 0 And OldEmpData2FileRec.EMPSTAO1 > 0 Then
      TEMPSTAO1(cnt) = 0
      GoTo BadStaAmtPct
    End If
    
    If OldEmpData2FileRec.EMPSTAO1 < 0 Then
      TEMPSTAO1(cnt) = 0
    Else
      TEMPSTAO1(cnt) = OldEmpData2FileRec.EMPSTAO1
    End If
BadStaAmtPct:
    If QPTrim$(OldEmpData2FileRec.EMPSTAS) = "" Then
      TEMPSTAS(cnt) = "S"
    Else
      TEMPSTAS(cnt) = OldEmpData2FileRec.EMPSTAS
    End If
    If OldEmpData2FileRec.EMPSTAA < 0 Then
      TEMPSTAA(cnt) = 0
    Else
      TEMPSTAA(cnt) = OldEmpData2FileRec.EMPSTAA
    End If
    If OldEmpData2FileRec.EMPSTAAA < 0 Then
      TEMPSTAAA(cnt) = 0
    Else
      TEMPSTAAA(cnt) = OldEmpData2FileRec.EMPSTAAA
    End If
    If QPTrim$(OldEmpData2FileRec.EMPSOCX) = "" Then
      TEMPSOCX(cnt) = "N"
    Else
      TEMPSOCX(cnt) = OldEmpData2FileRec.EMPSOCX
    End If
    If QPTrim$(OldEmpData2FileRec.EMPMEDX) = "" Then
      TEMPMEDX(cnt) = "N"
    Else
      TEMPMEDX(cnt) = OldEmpData2FileRec.EMPMEDX
    End If
    TempEIC(cnt) = OldEmpData2FileRec.EMPEIC
    For Cnt2 = 1 To 12
      If OldEmpData2FileRec.EmpDed(Cnt2).DAmt < 0 Then
        TempDed(cnt, Cnt2).DAmt = 0
      Else
        TempDed(cnt, Cnt2).DAmt = OldEmpData2FileRec.EmpDed(Cnt2).DAmt
      End If
      If Len(OldEmpData2FileRec.EmpDed(Cnt2).DPct) > 0 And Val(OldEmpData2FileRec.EmpDed(Cnt2).DAmt) <= 0 Then
        TempDed(cnt, Cnt2).DPct = ""
      Else
        TempDed(cnt, Cnt2).DPct = OldEmpData2FileRec.EmpDed(Cnt2).DPct
      End If
    TempDed(cnt, Cnt2).DOTI = OldEmpData2FileRec.EmpDed(Cnt2).DOTI
    Next Cnt2
    TEMPEACT1(cnt) = OldEmpData2FileRec.EMPEACT1
    If OldEmpData2FileRec.EMPEAMT1 < 0 Then
      TEMPEAMT1(cnt) = 0
    Else
      TEMPEAMT1(cnt) = OldEmpData2FileRec.EMPEAMT1
    End If
    
    TEMPEACT2(cnt) = OldEmpData2FileRec.EMPEACT2
    If OldEmpData2FileRec.EMPEAMT2 < 0 Then
      TEMPEAMT2(cnt) = 0
    Else
      TEMPEAMT2(cnt) = OldEmpData2FileRec.EMPEAMT2
    End If
    TEMPEACT3(cnt) = OldEmpData2FileRec.EMPEACT3
    If OldEmpData2FileRec.EMPEAMT3 < -0 Then
      TEMPEAMT3(cnt) = 0
    Else
      TEMPEAMT3(cnt) = OldEmpData2FileRec.EMPEAMT3
    End If
    TEMPHP(cnt) = OldEmpData2FileRec.EMPHP
    For Cnt2 = 1 To 8
      TEDist(cnt, Cnt2).DAcct = OldEmpData2FileRec.EDist(Cnt2).DAcct
      If OldEmpData2FileRec.EDist(Cnt2).DAmt < 0 Then
        TEDist(cnt, Cnt2).DAmt = 0
      Else
        TEDist(cnt, Cnt2).DAmt = OldEmpData2FileRec.EDist(Cnt2).DAmt
      End If
    Next Cnt2
    
    If OldEmpData2FileRec.EMPVACE < -1000 Then
      TEMPVACE(cnt) = 0
    Else
      TEMPVACE(cnt) = OldEmpData2FileRec.EMPVACE
    End If
    If OldEmpData2FileRec.EMPVUSED < -1000 Then
      TEMPVUSED(cnt) = 0
    Else
      TEMPVUSED(cnt) = OldEmpData2FileRec.EMPVUSED
    End If
    If OldEmpData2FileRec.EMPVBAL < -1000 Then
      TEMPVBAL(cnt) = 0
    Else
      TEMPVBAL(cnt) = OldEmpData2FileRec.EMPVBAL
    End If
    
    If OldEmpData2FileRec.EMPSLE < -1000 Then
      TEMPSLE(cnt) = 0
    Else
      TEMPSLE(cnt) = OldEmpData2FileRec.EMPSLE
    End If
    If OldEmpData2FileRec.EMPSLUSE < -1000 Then
      TEMPSLUSE(cnt) = 0
    Else
      TEMPSLUSE(cnt) = OldEmpData2FileRec.EMPSLUSE
    End If
    If OldEmpData2FileRec.EMPSLBAL < -1000 Then
      TEMPSLBAL(cnt) = 0
    Else
      TEMPSLBAL(cnt) = OldEmpData2FileRec.EMPSLBAL
    End If
    
    If OldEmpData2FileRec.EMPCTE < -1000 Then
      TEMPCTE(cnt) = 0
    Else
      TEMPCTE(cnt) = OldEmpData2FileRec.EMPCTE
    End If
    If OldEmpData2FileRec.EMPCTUSE < -1000 Then
      TEMPCTUSE(cnt) = 0
    Else
      TEMPCTUSE(cnt) = OldEmpData2FileRec.EMPCTUSE
    End If
    If OldEmpData2FileRec.EMPCTBAL < -1000 Then
      TEMPCTBAL(cnt) = 0
    Else
      TEMPCTBAL(cnt) = OldEmpData2FileRec.EMPCTBAL
    End If
    
    If OldEmpData2FileRec.PERERN < -1000 Then
      TPERERN(cnt) = 0
    Else
      TPERERN(cnt) = OldEmpData2FileRec.PERERN
    End If
    If OldEmpData2FileRec.PerUsed < -1000 Then
      TPerUsed(cnt) = 0
    Else
      TPerUsed(cnt) = OldEmpData2FileRec.PerUsed
    End If
    If OldEmpData2FileRec.PERBAL < -1000 Then
      TPERBAL(cnt) = 0
    Else
      TPERBAL(cnt) = OldEmpData2FileRec.PERBAL
    End If
    
    If OldEmpData2FileRec.HOLERN < -1000 Then
      THOLERN(cnt) = 0
    Else
      THOLERN(cnt) = OldEmpData2FileRec.HOLERN
    End If
    If OldEmpData2FileRec.HolUsed < -1000 Then
      THolUsed(cnt) = 0
    Else
      THolUsed(cnt) = OldEmpData2FileRec.HolUsed
    End If
    If OldEmpData2FileRec.HOLBAL < -1000 Then
      THOLBAL(cnt) = 0
    Else
      THOLBAL(cnt) = OldEmpData2FileRec.HOLBAL
    End If
    TLeaveTbl(cnt) = OldEmpData2FileRec.LeaveTbl
    If OldEmpData2FileRec.ExcludeESC = "0" Then OldEmpData2FileRec.ExcludeESC = "N"
    TExcludeESC(cnt) = OldEmpData2FileRec.ExcludeESC
    TUseLife(cnt) = OldEmpData2FileRec.UseLife
    If OldEmpData2FileRec.EmpPin < 0 Then
      TEmpPin(cnt) = 0
    Else
      TEmpPin(cnt) = OldEmpData2FileRec.EmpPin
    End If
    TLastTransRec(cnt) = OldEmpData2FileRec.LastTransRec
    TDeleted(cnt) = OldEmpData2FileRec.Deleted
    TLDTDate(cnt) = OldEmpData2FileRec.LDTDate
    TCDTDate(cnt) = OldEmpData2FileRec.CDTDate
    TInprocFlag(cnt) = OldEmpData2FileRec.InprocFlag
    TUnused(cnt) = OldEmpData2FileRec.Unused
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
    
    For Cnt2 = 1 To 12
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
    If TPERBAL(cnt) > 500000 Then TPERBAL(cnt) = 0
    NewEmpData2FileRec.PERBAL = TPERBAL(cnt)
    NewEmpData2FileRec.HOLERN = THOLERN(cnt)
    NewEmpData2FileRec.HolUsed = THolUsed(cnt)
    If THOLBAL(cnt) > 500000 Then THOLBAL(cnt) = 0
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
    If QPTrim$(TEMPRETTP(cnt)) = "" Then
      NewEmpData2FileRec.YN401K = "N"
    Else
      NewEmpData2FileRec.YN401K = "Y"
    End If
    NewEmpData2FileRec.PrimeDept = "1"
    NewEmpData2FileRec.HomePhone = "(000)-000-0000"
    NewEmpData2FileRec.EmrgncyCntctName = "Nothing Saved"
    NewEmpData2FileRec.EmrgncyCntctPhnNum = "(000)-000-0000"
    NewEmpData2FileRec.EmrgncyCntctRelation = "Nothing Saved"
    
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
  For X = 1 To 12
    TempUseDed(X) = OldPRDefaultRec.UseDed(X)
  Next X
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
  For X = 1 To 12
    NewPRDefaultRec.UseDed(X) = TempUseDed(X)
  Next X
  For X = 13 To 50
    NewPRDefaultRec.UseDed(X) = "N"
  Next X
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
  
  For X = 1 To OldDedCnt
    Get OldDedHandle, X, OldDedFileRec
    TempDCDESC1(X) = OldDedFileRec.DCDESC1
    TempDCACCT1(X) = OldDedFileRec.DCACCT1
    TempDCFWT1(X) = OldDedFileRec.DCFWT1
    TempDCSWT1(X) = OldDedFileRec.DCSWT1
    TempDCSOC1(X) = OldDedFileRec.DCSOC1
    TempDCMED1(X) = OldDedFileRec.DCMED1
  Next X
  Close OldDedHandle
  
  NewDedHandle = FreeFile
  DedRecLen = Len(NewDedFileRec)
  KillFile "prdata\" + "PRDEDCOD.DAT"
  Open "prdata\" + "PRDEDCOD.DAT" For Random Shared As NewDedHandle Len = DedRecLen
  
  Nextx = 1
  For X = 1 To OldDedCnt
    Get NewDedHandle, X, NewDedFileRec
    If QPTrim(TempDCDESC1(X)) = "" Then GoTo NoDesc
    NewDedFileRec.DCDESC1 = QPTrim$(TempDCDESC1(X))
    NewDedFileRec.DCACCT1 = QPTrim$(TempDCACCT1(X))
    NewDedFileRec.DCFWT1 = QPTrim$(TempDCFWT1(X))
    NewDedFileRec.DCSWT1 = QPTrim$(TempDCSWT1(X))
    NewDedFileRec.DCSOC1 = QPTrim$(TempDCSOC1(X))
    NewDedFileRec.DCMED1 = QPTrim$(TempDCMED1(X))
    Put NewDedHandle, Nextx, NewDedFileRec
    Nextx = Nextx + 1
NoDesc:
  Next X
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
    
  DosUnitRecLen = Len(UnitRec)
  DosFileHandle = FreeFile
  Open "PRData\PRUNIT.DAT" For Random Shared As DosFileHandle Len = DosUnitRecLen
  Get DosFileHandle, 1, DosUnitRec
  txtUpdate.Text = "Updating Unit Data"
  DoEvents

  If QPTrim$(DosUnitRec.BBTBATCH) = "" Then
    TempBBTBATCH = "0"
  Else
    TempBBTBATCH = QPTrim$(DosUnitRec.BBTBATCH)
  End If
  Close DosFileHandle
  
  If QPTrim$(DosUnitRec.BBTCNTNO) = "" Then
    TempBBTCNTNO = "0"
  Else
    TempBBTCNTNO = QPTrim$(DosUnitRec.BBTCNTNO)
  End If
  
  If Len(DosUnitRec.FileVer) = 0 Then
    TempFileVer = 0
  Else
     TempFileVer = DosUnitRec.FileVer
  End If
  
  If QPTrim$(DosUnitRec.UFCITY) = "" Then
    TempUFCITY = "YourTown"
  Else
    TempUFCITY = QPTrim$(DosUnitRec.UFCITY)
  End If
  
  If QPTrim$(DosUnitRec.UFEMPR) = "" Then
    TempUFEMPR = "YourTown"
  Else
    TempUFEMPR = QPTrim$(DosUnitRec.UFEMPR)
  End If
  
  If QPTrim$(DosUnitRec.UFFEDID) = "" Then
    TempUFFEDID = "0"
  Else
    TempUFFEDID = QPTrim$(DosUnitRec.UFFEDID)
  End If
  
  If QPTrim$(DosUnitRec.UFSTAID) = "" Then
    TempUFSTAID = "0"
  Else
    TempUFSTAID = QPTrim$(DosUnitRec.UFSTAID)
  End If
  
  If QPTrim$(DosUnitRec.UFRETID) = "" Then
    TempUFRETID = "0"
  Else
    TempUFRETID = QPTrim$(DosUnitRec.UFRETID)
  End If
  
  If QPTrim$(DosUnitRec.UFRETIDL) = "" Then
    TempUFRETIDL = "0"
  Else
    TempUFRETIDL = QPTrim$(DosUnitRec.UFRETIDL)
  End If
  
  If QPTrim$(DosUnitRec.IMMDNUM) = "" Then
    TempIMMDNUM = "0"
  Else
    TempIMMDNUM = QPTrim$(DosUnitRec.IMMDNUM)
  End If
  
  If QPTrim$(DosUnitRec.IMMONUM) = "" Then
    TempIMMONUM = "0"
  Else
    TempIMMONUM = QPTrim$(DosUnitRec.IMMONUM)
  End If
  
  If DosUnitRec.ESCRTYPE < 0 Then
    TempESCRTYPE = 0
  Else
    TempESCRTYPE = DosUnitRec.ESCRTYPE
  End If
  
  If DosUnitRec.TAXWBASE < 0 Then
    TempTAXWBASE = 0
  Else
    TempTAXWBASE = DosUnitRec.TAXWBASE
  End If
  
  If QPTrim$(DosUnitRec.UFSTATE) = "" Then
    TempUFSTATE = "NC"
  Else
    TempUFSTATE = QPTrim$(DosUnitRec.UFSTATE)
  End If
  
  UnitRecLen = Len(UnitRec)
  FileHandle = FreeFile
  Open "PRData\PRUNIT.DAT" For Random Shared As FileHandle Len = UnitRecLen
  Get FileHandle, 1, UnitRec
  UnitRec.BBTBATCH = TempBBTBATCH
  UnitRec.BBTCNTNO = TempBBTCNTNO
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
  UnitRec.LMT401YN = "N"
  UnitRec.FileVer = "Done" 'once set prohibits this conversion to
  'be run again when accessed by way of the payroll.exe
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
  
'  If Not Exist("PRDATA\PRTRANSH.DAT") Then GoTo NoTransHistorySaved 'for Canton install only
  
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

  For X = 1 To OLDTransHistRecCnt
    Get OldTransHistFileHandle, X, OldTransHistFileRec
    TempTActive(X) = OldTransHistFileRec.TActive
    TempPrevTransRec(X) = OldTransHistFileRec.PrevTransRec
    TempEmpPin(X) = OldTransHistFileRec.EmpPin
    TempPaySFlag(X) = OldTransHistFileRec.PaySFlag
    If OldTransHistFileRec.NetPay < 0 Then
      If OldTransHistFileRec.CheckNum > 0 Then
        OldTransHistFileRec.CheckNum = -OldTransHistFileRec.CheckNum
      End If
    End If
    TempCheckNum(X) = OldTransHistFileRec.CheckNum
    TempPayPdStart(X) = OldTransHistFileRec.PayPdStart
    TempPayPdEnd(X) = OldTransHistFileRec.PayPdEnd
    TempCheckDate(X) = OldTransHistFileRec.CheckDate
    TempPostDate(X) = OldTransHistFileRec.PostDate
    TempPayType(X) = OldTransHistFileRec.PayType
    TempBaseRate(X) = OldRound#(OldTransHistFileRec.BaseRate)
    TempOTRate(X) = OldRound#(OldTransHistFileRec.OTRate)
    If OldTransHistFileRec.RegHrsWork > 0 And OldTransHistFileRec.RegHrsWork < 1 Then
      OldTransHistFileRec.RegHrsWork = 0
    End If
    TempRegHrsWork(X) = OldRound#(OldTransHistFileRec.RegHrsWork)
    TempVacUsed(X) = OldRound#(OldTransHistFileRec.VacUsed)
    TempSickUsed(X) = OldRound#(OldTransHistFileRec.SickUsed)
    TempCompUsed(X) = OldRound#(OldTransHistFileRec.CompUsed)
    TempHOLHOURS(X) = OldRound#(OldTransHistFileRec.HOLHOURS)
    TempPerHours(X) = OldRound#(OldTransHistFileRec.PerHours)
    TempRegHrsPaid(X) = OldRound#(OldTransHistFileRec.RegHrsPaid)
    TempOTHours(X) = OldRound#(OldTransHistFileRec.OTHours)
    TempOTHrsPaid(X) = OldRound#(OldTransHistFileRec.OTHrsPaid)
    TempOT2COMP(X) = OldRound#(OldTransHistFileRec.OT2Comp)
    For y = 1 To 8
      OldTransHistFileRec.TDist(y).DOHrs = OldRound#(OldTransHistFileRec.TDist(y).DOHrs)
      OldTransHistFileRec.TDist(y).DOWage = OldRound#(OldTransHistFileRec.TDist(y).DOWage)
      OldTransHistFileRec.TDist(y).DPct = OldRound#(OldTransHistFileRec.TDist(y).DPct)
      OldTransHistFileRec.TDist(y).DRHrs = OldRound#(OldTransHistFileRec.TDist(y).DRHrs)
      OldTransHistFileRec.TDist(y).DRWage = OldRound#(OldTransHistFileRec.TDist(y).DRWage)
      TempTDist(X, y) = OldTransHistFileRec.TDist(y)
    Next y
    TempTotRegWage(X) = OldRound#(OldTransHistFileRec.TotRegWage)
    TempTotOTWage(X) = OldRound#(OldTransHistFileRec.TotOTWage)
    TempGrossWage(X) = OldRound#(OldTransHistFileRec.GrossWage)
    For y = 1 To 3
      TempEAmt(X, y) = OldRound#(OldTransHistFileRec.EAmt(y))
    Next y
    For y = 1 To 6
      OldTransHistFileRec.EDist(y).EAmt = OldRound#(OldTransHistFileRec.EDist(y).EAmt)
      TempEDist(X, y) = OldTransHistFileRec.EDist(y)
    Next y
    TempTotAdditEarn(X) = OldRound#(OldTransHistFileRec.TotAdditEarn)
    TempGrossPay(X) = OldRound#(OldTransHistFileRec.GrossPay)
    TempSocGrossPay(X) = OldRound#(OldTransHistFileRec.SocGrossPay)
    TempMedGrossPay(X) = OldRound#(OldTransHistFileRec.MedGrossPay)
    TempFedGrossPay(X) = OldRound#(OldTransHistFileRec.FedGrossPay)
    TempStaGrossPay(X) = OldRound#(OldTransHistFileRec.StaGrossPay)
    TempSocTaxAmt(X) = OldRound#(OldTransHistFileRec.SocTaxAmt)
    TempMedTaxAmt(X) = OldRound#(OldTransHistFileRec.MedTaxAmt)
    TempFedTaxAmt(X) = OldRound#(OldTransHistFileRec.FedTaxAmt)
    TempStaTaxAmt(X) = OldRound#(OldTransHistFileRec.StaTaxAmt)
    TempTotTaxAmt(X) = OldRound#(OldTransHistFileRec.TotTaxAmt)
    TempRetireAmt(X) = OldTransHistFileRec.RetireAmt
    For y = 1 To 12
      TempDAmt(X, y) = OldRound#(OldTransHistFileRec.DAmt(y))
    Next y
    TempTotDedAmt(X) = OldRound#(OldTransHistFileRec.TotDedAmt)
    TempEICAmt(X) = OldRound#(OldTransHistFileRec.EICAmt)
    TempNetPay(X) = OldRound#(OldTransHistFileRec.NetPay)
    TempPeriodHistRec(X) = OldTransHistFileRec.PeriodHistRec
    TempMatchRetAmt(X) = OldRound#(OldTransHistFileRec.MatchRetAmt)
    TempMatchSocAmt(X) = OldRound#(OldTransHistFileRec.MatchSocAmt)
    TempMatchMedAmt(X) = OldRound#(OldTransHistFileRec.MatchMedAmt)
    TempRetGrossPay(X) = OldRound#(OldTransHistFileRec.RetGrossPay)
    TempTaxFring(X) = OldRound#(OldTransHistFileRec.TaxFring)
  Next X

  Close OldTransHistFileHandle

  NewTransHistRecLen = Len(NewTransHistFileRec)
  NewTransHistFileHandle = FreeFile
  Open "prdata\" + "PRTRANSH.DAT" For Random Shared As NewTransHistFileHandle Len = NewTransHistRecLen

  For X = 1 To OLDTransHistRecCnt
    NewTransHistFileRec.TActive = TempTActive(X)
    NewTransHistFileRec.PrevTransRec = TempPrevTransRec(X)
    NewTransHistFileRec.EmpPin = TempEmpPin(X)
    NewTransHistFileRec.PaySFlag = TempPaySFlag(X)
    NewTransHistFileRec.CheckNum = TempCheckNum(X)
    NewTransHistFileRec.PayPdStart = TempPayPdStart(X)
    NewTransHistFileRec.PayPdEnd = TempPayPdEnd(X)
    NewTransHistFileRec.CheckDate = TempCheckDate(X)
    NewTransHistFileRec.PostDate = TempPostDate(X)
    NewTransHistFileRec.PayType = TempPayType(X)
    NewTransHistFileRec.BaseRate = TempBaseRate(X)
    NewTransHistFileRec.OTRate = TempOTRate(X)
    NewTransHistFileRec.RegHrsWork = TempRegHrsWork(X)
    NewTransHistFileRec.VacUsed = TempVacUsed(X)
    NewTransHistFileRec.SickUsed = TempSickUsed(X)
    NewTransHistFileRec.CompUsed = TempCompUsed(X)
    NewTransHistFileRec.HOLHOURS = TempHOLHOURS(X)
    NewTransHistFileRec.PerHours = TempPerHours(X)
    NewTransHistFileRec.RegHrsPaid = TempRegHrsPaid(X)
    NewTransHistFileRec.OTHours = TempOTHours(X)
    NewTransHistFileRec.OTHrsPaid = TempOTHrsPaid(X)
    NewTransHistFileRec.OT2Comp = TempOT2COMP(X)
    For y = 1 To 8
      NewTransHistFileRec.TDist(y) = TempTDist(X, y)
    Next y
    NewTransHistFileRec.TotRegWage = TempTotRegWage(X)
    NewTransHistFileRec.TotOTWage = TempTotOTWage(X)
    NewTransHistFileRec.GrossWage = TempGrossWage(X)
    For y = 1 To 3
      NewTransHistFileRec.EAmt(y) = TempEAmt(X, y)
    Next y
    For y = 1 To 6
      NewTransHistFileRec.EDist(y) = TempEDist(X, y)
    Next y
    NewTransHistFileRec.TotAdditEarn = TempTotAdditEarn(X)
    NewTransHistFileRec.GrossPay = TempGrossPay(X)
    NewTransHistFileRec.SocGrossPay = TempSocGrossPay(X)
    NewTransHistFileRec.MedGrossPay = TempMedGrossPay(X)
    NewTransHistFileRec.FedGrossPay = TempFedGrossPay(X)
    NewTransHistFileRec.StaGrossPay = TempStaGrossPay(X)
    NewTransHistFileRec.SocTaxAmt = TempSocTaxAmt(X)
    NewTransHistFileRec.MedTaxAmt = TempMedTaxAmt(X)
    NewTransHistFileRec.FedTaxAmt = TempFedTaxAmt(X)
    NewTransHistFileRec.StaTaxAmt = TempStaTaxAmt(X)
    NewTransHistFileRec.TotTaxAmt = TempTotTaxAmt(X)
    NewTransHistFileRec.RetireAmt = TempRetireAmt(X)
    For y = 1 To 12
      NewTransHistFileRec.DAmt(y) = TempDAmt(X, y)
    Next y
    For y = 13 To 50
      NewTransHistFileRec.DAmt(y) = 0#
    Next y
    
    NewTransHistFileRec.TotDedAmt = TempTotDedAmt(X)
    NewTransHistFileRec.EICAmt = TempEICAmt(X)
    NewTransHistFileRec.NetPay = TempNetPay(X)
    NewTransHistFileRec.PeriodHistRec = TempPeriodHistRec(X)
    NewTransHistFileRec.MatchRetAmt = TempMatchRetAmt(X)
    NewTransHistFileRec.MatchSocAmt = TempMatchSocAmt(X)
    NewTransHistFileRec.MatchMedAmt = TempMatchMedAmt(X)
    NewTransHistFileRec.RetGrossPay = TempRetGrossPay(X)
    NewTransHistFileRec.TaxFring = TempTaxFring(X)
    For y = 1 To 3
      NewTransHistFileRec.Less401k(y) = False
    Next y
    NewTransHistFileRec.Pad1 = ""
    Put NewTransHistFileHandle, X, NewTransHistFileRec
  Next X

  Close NewTransHistFileHandle

'NoTransHistorySaved: For Canton install only
  OldTransWorkRecLen = Len(OldTransWorkFileRec)
  OldTransWorkFileHandle = FreeFile
  Open "prdata\" + "PRTRANST.DAT" For Random Shared As OldTransWorkFileHandle Len = OldTransWorkRecLen
  OLDTransWorkRecCnt = LOF(OldTransWorkFileHandle) / Len(OldTransWorkFileRec)
  Dim TryThis As Double
  For X = 1 To OLDTransWorkRecCnt
    Get OldTransWorkFileHandle, X, OldTransWorkFileRec
    TempTActive(X) = 0 'OldTransWorkFileRec.TActive
    TempPrevTransRec(X) = OldTransWorkFileRec.PrevTransRec
    TempEmpPin(X) = OldTransWorkFileRec.EmpPin
    TempPaySFlag(X) = OldTransWorkFileRec.PaySFlag
    TempCheckNum(X) = OldTransWorkFileRec.CheckNum
    TempPayPdStart(X) = OldTransWorkFileRec.PayPdStart
    TempPayPdEnd(X) = OldTransWorkFileRec.PayPdEnd
    TempCheckDate(X) = OldTransWorkFileRec.CheckDate
    TempPostDate(X) = OldTransWorkFileRec.PostDate
    TempPayType(X) = OldTransWorkFileRec.PayType
    TempBaseRate(X) = OldTransWorkFileRec.BaseRate
    TempOTRate(X) = OldTransWorkFileRec.OTRate
    TempRegHrsWork(X) = OldTransWorkFileRec.RegHrsWork
    TempVacUsed(X) = OldTransWorkFileRec.VacUsed
    TempSickUsed(X) = OldTransWorkFileRec.SickUsed
    TempCompUsed(X) = OldTransWorkFileRec.CompUsed
    TempHOLHOURS(X) = OldTransWorkFileRec.HOLHOURS
    TempPerHours(X) = OldTransWorkFileRec.PerHours
    TempRegHrsPaid(X) = OldTransWorkFileRec.RegHrsPaid
    TempOTHours(X) = OldTransWorkFileRec.OTHours
    TempOTHrsPaid(X) = OldTransWorkFileRec.OTHrsPaid
    TempOT2COMP(X) = OldTransWorkFileRec.OT2Comp
    For y = 1 To 8
      TempTDist(X, y) = OldTransWorkFileRec.TDist(y)
    Next y
    TempTotRegWage(X) = OldTransWorkFileRec.TotRegWage
    TempTotOTWage(X) = OldTransWorkFileRec.TotOTWage
    TempGrossWage(X) = OldTransWorkFileRec.GrossWage
    For y = 1 To 3
      TempEAmt(X, y) = OldTransWorkFileRec.EAmt(y)
    Next y
    For y = 1 To 6
      TempEDist(X, y) = OldTransWorkFileRec.EDist(y)
    Next y
    TempTotAdditEarn(X) = OldTransWorkFileRec.TotAdditEarn
    TempGrossPay(X) = OldTransWorkFileRec.GrossPay
    TempSocGrossPay(X) = OldTransWorkFileRec.SocGrossPay
    TempMedGrossPay(X) = OldTransWorkFileRec.MedGrossPay
    TempFedGrossPay(X) = OldTransWorkFileRec.FedGrossPay
    TempStaGrossPay(X) = OldTransWorkFileRec.StaGrossPay
    TempSocTaxAmt(X) = OldTransWorkFileRec.SocTaxAmt
    TempMedTaxAmt(X) = OldTransWorkFileRec.MedTaxAmt
    TempFedTaxAmt(X) = OldTransWorkFileRec.FedTaxAmt
    TempStaTaxAmt(X) = OldTransWorkFileRec.StaTaxAmt
    TempTotTaxAmt(X) = OldTransWorkFileRec.TotTaxAmt
    TempRetireAmt(X) = OldTransWorkFileRec.RetireAmt
    For y = 1 To 12
      TempDAmt(X, y) = OldTransWorkFileRec.DAmt(y)
    Next y
    TempTotDedAmt(X) = OldTransWorkFileRec.TotDedAmt
    TempEICAmt(X) = OldTransWorkFileRec.EICAmt
    TempNetPay(X) = OldTransWorkFileRec.NetPay
    TempPeriodHistRec(X) = OldTransWorkFileRec.PeriodHistRec
    TempMatchRetAmt(X) = OldTransWorkFileRec.MatchRetAmt
    TempMatchSocAmt(X) = OldTransWorkFileRec.MatchSocAmt
    TempMatchMedAmt(X) = OldTransWorkFileRec.MatchMedAmt
    TempRetGrossPay(X) = OldTransWorkFileRec.RetGrossPay
    TempTaxFring(X) = OldTransWorkFileRec.TaxFring
  Next X

  Close OldTransWorkFileHandle

  NewTransWorkRecLen = Len(NewTransWorkFileRec)
  NewTransWorkFileHandle = FreeFile
  Open "prdata\" + "PRTRANST.DAT" For Random Shared As NewTransWorkFileHandle Len = NewTransWorkRecLen

  For X = 1 To OLDTransWorkRecCnt
    NewTransWorkFileRec.TActive = TempTActive(X)
    NewTransWorkFileRec.PrevTransRec = TempPrevTransRec(X)
    NewTransWorkFileRec.EmpPin = TempEmpPin(X)
    NewTransWorkFileRec.PaySFlag = TempPaySFlag(X)
    NewTransWorkFileRec.CheckNum = TempCheckNum(X)
    NewTransWorkFileRec.PayPdStart = TempPayPdStart(X)
    NewTransWorkFileRec.PayPdEnd = TempPayPdEnd(X)
    NewTransWorkFileRec.CheckDate = TempCheckDate(X)
    NewTransWorkFileRec.PostDate = TempPostDate(X)
    NewTransWorkFileRec.PayType = TempPayType(X)
    NewTransWorkFileRec.BaseRate = TempBaseRate(X)
    NewTransWorkFileRec.OTRate = TempOTRate(X)
    NewTransWorkFileRec.RegHrsWork = TempRegHrsWork(X)
    NewTransWorkFileRec.VacUsed = TempVacUsed(X)
    NewTransWorkFileRec.SickUsed = TempSickUsed(X)
    NewTransWorkFileRec.CompUsed = TempCompUsed(X)
    NewTransWorkFileRec.HOLHOURS = TempHOLHOURS(X)
    NewTransWorkFileRec.PerHours = TempPerHours(X)
    NewTransWorkFileRec.RegHrsPaid = TempRegHrsPaid(X)
    NewTransWorkFileRec.OTHours = TempOTHours(X)
    NewTransWorkFileRec.OTHrsPaid = TempOTHrsPaid(X)
    NewTransWorkFileRec.OT2Comp = TempOT2COMP(X)
    For y = 1 To 8
      NewTransWorkFileRec.TDist(y) = TempTDist(X, y)
    Next y
    NewTransWorkFileRec.TotRegWage = TempTotRegWage(X)
    NewTransWorkFileRec.TotOTWage = TempTotOTWage(X)
    NewTransWorkFileRec.GrossWage = TempGrossWage(X)
    For y = 1 To 3
      NewTransWorkFileRec.EAmt(y) = TempEAmt(X, y)
    Next y
    For y = 1 To 6
      NewTransWorkFileRec.EDist(y) = TempEDist(X, y)
    Next y
    NewTransWorkFileRec.TotAdditEarn = TempTotAdditEarn(X)
    NewTransWorkFileRec.GrossPay = TempGrossPay(X)
    NewTransWorkFileRec.SocGrossPay = TempSocGrossPay(X)
    NewTransWorkFileRec.MedGrossPay = TempMedGrossPay(X)
    NewTransWorkFileRec.FedGrossPay = TempFedGrossPay(X)
    NewTransWorkFileRec.StaGrossPay = TempStaGrossPay(X)
    NewTransWorkFileRec.SocTaxAmt = TempSocTaxAmt(X)
    NewTransWorkFileRec.MedTaxAmt = TempMedTaxAmt(X)
    NewTransWorkFileRec.FedTaxAmt = TempFedTaxAmt(X)
    NewTransWorkFileRec.StaTaxAmt = TempStaTaxAmt(X)
    NewTransWorkFileRec.TotTaxAmt = TempTotTaxAmt(X)
    NewTransWorkFileRec.RetireAmt = TempRetireAmt(X)
    For y = 1 To 12
      NewTransWorkFileRec.DAmt(y) = TempDAmt(X, y)
    Next y
    
    For y = 13 To 50
      NewTransWorkFileRec.DAmt(y) = 0#
    Next y
      
    NewTransWorkFileRec.TotDedAmt = TempTotDedAmt(X)
    NewTransWorkFileRec.EICAmt = TempEICAmt(X)
    NewTransWorkFileRec.NetPay = TempNetPay(X)
    NewTransWorkFileRec.PeriodHistRec = TempPeriodHistRec(X)
    NewTransWorkFileRec.MatchRetAmt = TempMatchRetAmt(X)
    NewTransWorkFileRec.MatchSocAmt = TempMatchSocAmt(X)
    NewTransWorkFileRec.MatchMedAmt = TempMatchMedAmt(X)
    NewTransWorkFileRec.RetGrossPay = TempRetGrossPay(X)
    NewTransWorkFileRec.TaxFring = TempTaxFring(X)
    For y = 1 To 3
      NewTransWorkFileRec.Less401k(y) = False '9/26/03
    Next y
    NewTransWorkFileRec.Pad1 = ""
    Put NewTransWorkFileHandle, X, NewTransWorkFileRec
  Next X
  
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
  
  For X = 1 To NumOfChkRecs
    Get OldChecksHandle, X, OldChkRec
    TempCActive(X) = OldChkRec.CActive
    TempCheckNum(X) = OldChkRec.CheckNum
    TempCheckDate(X) = OldChkRec.CheckDate
    TempEmpName(X) = OldChkRec.EmpName
    TempEmpNo(X) = OldChkRec.EmpNo
    TempEmpSSN(X) = OldChkRec.EmpSSN
    TempEmpAddr1(X) = OldChkRec.EmpAddr1
    TempEmpCity(X) = OldChkRec.EmpCity
    TempEmpState(X) = OldChkRec.EmpState
    TempEmpZip(X) = OldChkRec.EmpZip
    TempPayEndDate(X) = OldChkRec.PayEndDate
    TempBaseRate(X) = OldChkRec.BaseRate
    TempGrossPay(X) = OldChkRec.GrossPay
    TempFedTaxAmt(X) = OldChkRec.FedTaxAmt
    TempStaTaxAmt(X) = OldChkRec.StaTaxAmt
    TempMedTaxAmt(X) = OldChkRec.MedTaxAmt
    TempSocTaxAmt(X) = OldChkRec.SocTaxAmt
    TempTotDedAmt(X) = OldChkRec.TotDedAmt
    TempRetireAmt(X) = OldChkRec.RetireAmt
    TempNetPay(X) = OldChkRec.NetPay
    TempYTDGrossPay(X) = OldChkRec.YTDGrossPay
    TempYTDFederal(X) = OldChkRec.YTDFederal
    TempYTDState(X) = OldChkRec.YTDState
    TempYTDSocial(X) = OldChkRec.YTDSocial
    TempYTDMedicare(X) = OldChkRec.YTDMedicare
    TempYTDTotDed(X) = OldChkRec.YTDTotDed
    TempYTDNetPay(X) = OldChkRec.YTDNetPay
    TempYTDRetire(X) = OldChkRec.YTDRetire
    TempVactBal(X) = OldChkRec.VactBal
    TempSickBal(X) = OldChkRec.SickBal
    TempCompBal(X) = OldChkRec.CompBal
    TempCompEarn(X) = OldChkRec.CompEarn
    TempRegHrsWork(X) = OldChkRec.RegHrsWork
    TempOTHrsPaid(X) = OldChkRec.OTHrsPaid
    TempTotRegWage(X) = OldChkRec.TotRegWage
    TempVacUsed(X) = OldChkRec.VacUsed
    TempSickUsed(X) = OldChkRec.SickUsed
    TempCompUsed(X) = OldChkRec.CompUsed
    TempHolUsed(X) = OldChkRec.HolUsed
    TempPerUsed(X) = OldChkRec.PerUsed
    TempRegHrsPaid(X) = OldChkRec.RegHrsPaid
    TempTotOTWage(X) = OldChkRec.TotOTWage
    For y = 1 To 3
      TempAEarn(X, y).DAmt = OldChkRec.AEarn(y).DAmt
      TempAEarn(X, y).DCode = OldChkRec.AEarn(y).DCode
      TempAEarn(X, y).YTDDAmt = OldChkRec.AEarn(y).YTDDAmt
    Next y
    TempTotAdditEarn(X) = OldChkRec.TotAdditEarn
    TempEICAmt(X) = OldChkRec.EICAmt
    TempTaxFring(X) = OldChkRec.TaxFring
    For y = 1 To 12
      TempCDED(X, y).DAmt = OldChkRec.CDED(y).DAmt
      TempCDED(X, y).DCode = OldChkRec.CDED(y).DCode
      TempCDED(X, y).YTDDAmt = OldChkRec.CDED(y).YTDDAmt
    Next y
    TempDDFlag(X) = OldChkRec.DDFlag
  Next X
  
  Close OldChecksHandle
  
  NewChecksHandle = FreeFile
  Open "prdata\" + "PRCHECKS.DAT" For Random Shared As NewChecksHandle Len = NewChecksRecLen
  
  For X = 1 To NumOfChkRecs
    Get NewChecksHandle, X, NewChkRec
    NewChkRec.CActive = TempCActive(X)
    NewChkRec.CheckNum = TempCheckNum(X)
    NewChkRec.CheckDate = TempCheckDate(X)
    NewChkRec.EmpName = TempEmpName(X)
    NewChkRec.EmpNo = TempEmpNo(X)
    NewChkRec.EmpSSN = TempEmpSSN(X)
    NewChkRec.EmpAddr1 = TempEmpAddr1(X)
    NewChkRec.EmpCity = TempEmpCity(X)
    NewChkRec.EmpState = TempEmpState(X)
    NewChkRec.EmpZip = TempEmpZip(X)
    NewChkRec.PayEndDate = TempPayEndDate(X)
    NewChkRec.BaseRate = TempBaseRate(X)
    NewChkRec.GrossPay = TempGrossPay(X)
    NewChkRec.FedTaxAmt = TempFedTaxAmt(X)
    NewChkRec.StaTaxAmt = TempStaTaxAmt(X)
    NewChkRec.MedTaxAmt = TempMedTaxAmt(X)
    NewChkRec.SocTaxAmt = TempSocTaxAmt(X)
    NewChkRec.TotDedAmt = TempTotDedAmt(X)
    NewChkRec.RetireAmt = TempRetireAmt(X)
    NewChkRec.NetPay = TempNetPay(X)
    NewChkRec.YTDGrossPay = TempYTDGrossPay(X)
    NewChkRec.YTDFederal = TempYTDFederal(X)
    NewChkRec.YTDState = TempYTDState(X)
    NewChkRec.YTDSocial = TempYTDSocial(X)
    NewChkRec.YTDMedicare = TempYTDMedicare(X)
    NewChkRec.YTDTotDed = TempYTDTotDed(X)
    NewChkRec.YTDNetPay = TempYTDNetPay(X)
    NewChkRec.YTDRetire = TempYTDRetire(X)
    NewChkRec.VactBal = TempVactBal(X)
    NewChkRec.SickBal = TempSickBal(X)
    NewChkRec.CompBal = TempCompBal(X)
    NewChkRec.CompEarn = TempCompEarn(X)
    NewChkRec.RegHrsWork = TempRegHrsWork(X)
    NewChkRec.OTHrsPaid = TempOTHrsPaid(X)
    NewChkRec.TotRegWage = TempTotRegWage(X)
    NewChkRec.VacUsed = TempVacUsed(X)
    NewChkRec.SickUsed = TempSickUsed(X)
    NewChkRec.CompUsed = TempCompUsed(X)
    NewChkRec.HolUsed = TempHolUsed(X)
    NewChkRec.PerUsed = TempPerUsed(X)
    NewChkRec.RegHrsPaid = TempRegHrsPaid(X)
    NewChkRec.TotOTWage = TempTotOTWage(X)
    For y = 1 To 3
      NewChkRec.AEarn(y).DAmt = TempAEarn(X, y).DAmt
      NewChkRec.AEarn(y).DCode = TempAEarn(X, y).DCode
      NewChkRec.AEarn(y).YTDDAmt = TempAEarn(X, y).YTDDAmt
    Next y
    NewChkRec.TotAdditEarn = TempTotAdditEarn(X)
    NewChkRec.EICAmt = TempEICAmt(X)
    NewChkRec.TaxFring = TempTaxFring(X)
    For y = 1 To 12
      NewChkRec.CDED(y).DAmt = TempCDED(X, y).DAmt
      NewChkRec.CDED(y).DCode = TempCDED(X, y).DCode
      NewChkRec.CDED(y).YTDDAmt = TempCDED(X, y).YTDDAmt
    Next y
    For y = 13 To 50
      NewChkRec.CDED(y).DAmt = 0
      NewChkRec.CDED(y).DCode = ""
      NewChkRec.CDED(y).YTDDAmt = 0
    Next y
    NewChkRec.DDFlag = TempDDFlag(X)
    
    Put NewChecksHandle, X, NewChkRec
  Next X
  Close NewChecksHandle
  
  
  Dim TempExtrYear As Integer
  Dim TempDeds(0 To 50) As W2DedType
  
  OldW2Handle = FreeFile
  Open "prdata\" + "PRW2SETU.DAT" For Random Shared As OldW2Handle Len = OldW2RecLen
  
  Get OldW2Handle, 1, OldW2SetUp
  TempExtrYear = OldW2SetUp.ExtrYear
  For X = 0 To 12
    TempDeds(X).AMTBOX = OldW2SetUp.Deds(X).AMTBOX
    TempDeds(X).CHKDED = OldW2SetUp.Deds(X).CHKDED
    TempDeds(X).DedCode = OldW2SetUp.Deds(X).DedCode
  Next X
  For X = 13 To 50
    TempDeds(X).AMTBOX = ""
    TempDeds(X).CHKDED = ""
    TempDeds(X).DedCode = ""
  Next X
  
  Close OldW2Handle
  
  NewW2Handle = FreeFile
  Open "prdata\" + "PRW2SETU.DAT" For Random Shared As NewW2Handle Len = NewW2RecLen
  NewW2SetUp.ExtrYear = TempExtrYear
  For X = 0 To 50
    NewW2SetUp.Deds(X).AMTBOX = TempDeds(X).AMTBOX
    NewW2SetUp.Deds(X).CHKDED = TempDeds(X).CHKDED
    NewW2SetUp.Deds(X).DedCode = TempDeds(X).DedCode
  Next X
  
  Put NewW2Handle, 1, NewW2SetUp
  Close NewW2Handle
  
  Dim DosErnHandle As Integer
  Dim DosErnRec As DosErnCodeRecType
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  
  OpenDosErnCodeFile DosErnHandle
  NumOfErns = LOF(DosErnHandle) / Len(DosErnRec)
  If NumOfErns = 0 Then
    GoTo NoErnMatchNeeded
  End If
  
  ReDim TempERNCODE1(1 To NumOfErns) As String * 10
  ReDim TempERNFWT1(1 To NumOfErns) As String * 1
  ReDim TempERNSWT1(1 To NumOfErns) As String * 1
  ReDim TempERNSOC1(1 To NumOfErns) As String * 1
  ReDim TempERNMED1(1 To NumOfErns) As String * 1
  ReDim TempERNRET1(1 To NumOfErns) As String * 1
  For X = 1 To NumOfErns
    Get DosErnHandle, X, DosErnRec
    TempERNCODE1(X) = QPTrim$(DosErnRec.ERNCODE1)
    TempERNFWT1(X) = DosErnRec.ERNFWT1
    TempERNSWT1(X) = DosErnRec.ERNSWT1
    TempERNSOC1(X) = DosErnRec.ERNSOC1
    TempERNMED1(X) = DosErnRec.ERNMED1
    TempERNRET1(X) = DosErnRec.ERNRET1
  Next X
  Close DosErnHandle
  
  OpenErnCodeFile ErnHandle
  For X = 1 To NumOfErns
    ErnRec.ERNCODE1 = QPTrim$(TempERNCODE1(X))
    ErnRec.ERNFWT1 = TempERNFWT1(X)
    ErnRec.ERNSWT1 = TempERNSWT1(X)
    ErnRec.ERNSOC1 = TempERNSOC1(X)
    ErnRec.ERNMED1 = TempERNMED1(X)
    ErnRec.ERNRET1 = TempERNRET1(X)
    ErnRec.EarnYN = "Y" '"Y"es Include
    ErnRec.Pad = ""
    Put ErnHandle, X, ErnRec
  Next X
  Close ErnHandle
NoErnMatchNeeded:
  txtUpdate.Visible = False
  cmdCancel.Enabled = True
  
  txtUpdate.Text = "UPDATE COMPLETE. PRESS OK TO GO TO PAYROLL MAIN MENU"
  cmdOK.Visible = True
  cmdConvert.Visible = False
  cmdCancel.Visible = False
  
  Exit Sub
ErrorHandler:
  frmWarnReloadPRData.Show vbModal
  Call cmdOK_Click
End Sub

Private Sub cmdConvert_Click()
  
  If chkNotFall04.Value = 0 And chkFall04Update.Value = 0 Then
    MsgBox "Please select which conversion should be processed."
  ElseIf chkNotFall04.Value = 1 Then
    Call Convert
  ElseIf chkFall04Update.Value = 1 Then
    Call ConvertFall04
  End If
End Sub

Private Sub cmdOK_Click()
  Close
  Unload frmConvertDos2Win
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

Private Sub ConvertFall04()
  Dim OldChkRec As DosPRCheckRecType
  Dim NewChkRec As PRCheckRecType
  Dim OldChecksHandle As Integer
  Dim NewChecksHandle As Integer
  Dim OldChecksRecLen As Integer
  Dim NewChecksRecLen As Integer
  Dim NumOfChkRecs As Long
  Dim OldPrnSet As DosPRNSetupRecType
  Dim NewPrnSet As PRNSetupRecType
  Dim OldEmpData2FileHandle As Integer
  Dim cnt As Integer
  Dim OldEmpData2FileRec As DosEmpData2Type
  Dim NewEmpData2FileRec As EmpData2TypeFall04
  Dim OldEmpData2RecLen As Integer
  Dim NewEmpData2RecLen As Integer
  Dim OldEmpData3FileRec As DosEmpData3Type
  Dim NewEmpData3FileRec As EmpData3Type
  Dim OldEmpData3RecLen As Integer
  Dim NewEmpData3RecLen As Integer
  Dim NumOfRecs As Integer
  Dim Old3Handle As Integer, New3Handle As Integer
  Dim Cnt2 As Integer
  Dim OldPrnLen As Integer, NewPrnLen As Integer
  Dim OldHandle As Integer, NewHandle As Integer
  Dim TempPrnLen As Integer
  Dim EICHandle As Integer
  Dim EICRec As EICRecType
  Dim EICFileRec As EICRecType
  Dim EICRecLen As Integer
  Dim DosFileHandle As Integer
  Dim DosUnitRec As DosUnitFileRecType
  Dim DosUnitRecLen As Integer
  Dim FileHandle As Integer
  Dim UnitRec As UnitFileRecTypeFall04
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
  Dim y As Integer
  Dim OldTransHistFileHandle As Integer
  Dim OldTransWorkFileHandle As Integer
  Dim NewTransWorkFileHandle As Integer
  Dim RetHandle As Integer
  Dim RetRec As RetireRecType
  Dim RetRecLen As Integer, Nextx As Integer
  Dim OLDTransHistRecCnt As Integer
  Dim OLDTransWorkRecCnt As Integer
  Dim X As Integer, TempDedRec(1 To 12) As Integer
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
  Dim ESCFileName$
  Dim ESCHandle As Integer
  Dim RemitNumb$, EmplrAcct$
  Dim OldLeaveRec As OldLeaveRecType
  Dim OldLvHandle As Integer
  Dim LeaveRec As LeaveRecType
  Dim NewLvHandle As Integer
  Dim NumOfLvRecs As Integer
  Dim USize As Integer
  
  On Error GoTo ErrorHandler
  USize = FileSize(PRData + "PRUNIT.DAT")
  
'  If USize = 398 Then
'    MsgBox "This data has already been converted for the Fall 2004 Update. Conversion attempt aborted."
'    Exit Sub
'  ElseIf USize = 381 Then
'    MsgBox "This conversion program is designed only for use when converting from DOS to Windows. This data only needs the Fall 2004 Update. Please run the CnvtPRFall04.exe. Conversion attempt aborted."
'    Exit Sub
'  End If
  
  InFileNames(1) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames, 1) = False Then
    Exit Sub
  End If
  
  OpenOldLeaveFileName OldLvHandle
  NumOfLvRecs = LOF(OldLvHandle) / Len(OldLeaveRec)
  
  ReDim TempVacMax(1 To NumOfLvRecs) As Double
  ReDim TempVEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  ReDim TempSICKMAX(1 To NumOfLvRecs) As Double
  ReDim TempSEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  
'  If Exist("PRDATA\PRSYS.DAT") Then
'    If FileLen("PRDATA\PRSYS.DAT") <> 337 Then
'      MsgBox "Conversion has already taken place for this customer. Please convert leave table files with the leave table specific conversion program."
'      Close
'      Exit Sub
'    End If
'  End If

  For X = 1 To NumOfLvRecs
    Get OldLvHandle, X, OldLeaveRec
    TempVacMax(X) = OldLeaveRec.VacMax
    TempSICKMAX(X) = OldLeaveRec.SICKMAX
    For y = 1 To 20
      TempVEntry(X, y).EARN = OldLeaveRec.VEntry(y).EARN
      TempSEntry(X, y).EARN = OldLeaveRec.SEntry(y).EARN
      TempVEntry(X, y).YEARS = OldLeaveRec.VEntry(y).YEARS
      TempSEntry(X, y).YEARS = OldLeaveRec.SEntry(y).YEARS
    Next y
  Next X
  Close OldLvHandle
  
  OpenLeaveFileName NewLvHandle
  For X = 1 To NumOfLvRecs
    LeaveRec.VacMax = TempVacMax(X)
    LeaveRec.SICKMAX = TempSICKMAX(X)
    LeaveRec.HolMax = 0
    LeaveRec.PerMax = 0
    For y = 1 To 20
      LeaveRec.VEntry(y).EARN = TempVEntry(X, y).EARN
      LeaveRec.SEntry(y).EARN = TempSEntry(X, y).EARN
      LeaveRec.HEntry(y).EARN = 0
      LeaveRec.PEntry(y).EARN = 0
      LeaveRec.VEntry(y).YEARS = TempVEntry(X, y).YEARS
      LeaveRec.SEntry(y).YEARS = TempSEntry(X, y).YEARS
      LeaveRec.HEntry(y).YEARS = 0
      LeaveRec.PEntry(y).YEARS = 0
    Next y
    Put NewLvHandle, X, LeaveRec
  Next X
  Close NewLvHandle
  
SkipLeaveConversion:
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
    Unload frmConvertDos2Win
    DoEvents
    frmDos2WinFilesMissing.Show
    Exit Sub
  End If
  
  cmdConvert.MousePointer = vbHourglass
  
  OldPrnLen = Len(OldPrnSet)
  NewPrnLen = Len(NewPrnSet)
    
  OldTransHistRecLen = Len(OldTransHistFileRec)
  OldEmpData2RecLen = Len(OldEmpData2FileRec)
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
  
  Open "prdata\" + "PREMP2.DAT" For Random Shared As OldHandle Len = OldEmpData2RecLen
  
  NumOfRecs = LOF(OldHandle) / Len(OldEmpData2FileRec)
  
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
  ReDim TempDed(1 To NumOfRecs, 1 To 12) As EmpDedType
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
    If cnt = 334 Then GoTo SkipMe
    Get OldHandle, cnt, OldEmpData2FileRec
    If QPTrim$(OldEmpData2FileRec.EmpNo) = "" Then
      TEmpNo(cnt) = "0"
    Else
      TEmpNo(cnt) = OldEmpData2FileRec.EmpNo
    End If
    If QPTrim$(OldEmpData2FileRec.EmpSSN) = "" Then
      TEmpSSN(cnt) = "000-00-000"
    Else
      TEmpSSN(cnt) = OldEmpData2FileRec.EmpSSN
    End If
    'Found some situations where employees were zeroed out so they had
    'a value but the program didn't recognize zero as a name
    If QPTrim$(OldEmpData2FileRec.EMPLNAME) = "" Or InStr(OldEmpData2FileRec.EMPLNAME, "0") Then
      TEMPLNAME(cnt) = "NONE"
    Else
      TEMPLNAME(cnt) = UCase(OldEmpData2FileRec.EMPLNAME)
    End If
    If QPTrim$(OldEmpData2FileRec.EMPFNAME) = "" Or InStr(OldEmpData2FileRec.EMPFNAME, "0") Then
      TEMPFNAME(cnt) = "NONE"
    Else
      TEMPFNAME(cnt) = UCase(OldEmpData2FileRec.EMPFNAME)
    End If
    
    If QPTrim$(OldEmpData2FileRec.EmpAddr1) = "" Then
      TEmpAddr1(cnt) = "None"
    Else
      TEmpAddr1(cnt) = OldEmpData2FileRec.EmpAddr1
    End If
    TEMPADDR2(cnt) = OldEmpData2FileRec.EMPADDR2
    
    If QPTrim$(OldEmpData2FileRec.EmpCity) = "" Then
      TEmpCity(cnt) = "None"
    Else
      TEmpCity(cnt) = OldEmpData2FileRec.EmpCity
    End If
    If QPTrim$(OldEmpData2FileRec.EmpState) = "" Then
      TEmpState(cnt) = "None"
    Else
      TEmpState(cnt) = OldEmpData2FileRec.EmpState
    End If
    
    If QPTrim$(OldEmpData2FileRec.EmpZip) = "" Then
      TEmpZip(cnt) = "00000"
    Else
      TEmpZip(cnt) = OldEmpData2FileRec.EmpZip
    End If
    If Val(OldEmpData2FileRec.EMPBDAY) <= -22000 Then
      TEMPBDAY(cnt) = 0
    Else
      TEMPBDAY(cnt) = OldEmpData2FileRec.EMPBDAY
    End If
    TEMPGENDR(cnt) = OldEmpData2FileRec.EMPGENDR
    TEMPRACE(cnt) = OldEmpData2FileRec.EMPRACE
    TEMPRETNO(cnt) = OldEmpData2FileRec.EMPRETNO
    TEMPRETTP(cnt) = OldEmpData2FileRec.EMPRETTP
    If QPTrim$(OldEmpData2FileRec.DRAFTCOD) = "0" Then
      OldEmpData2FileRec.DRAFTCOD = ""
    End If
    If QPTrim$(OldEmpData2FileRec.DRAFTCOD) = "" Then
      TDRAFTCOD(cnt) = ""
      TEMPDDACC(cnt) = ""
      TPRENOTED(cnt) = ""
      TBANKNAME(cnt) = ""
      TBANKLOC(cnt) = ""
      TTRANSIT(cnt) = ""
    Else
      TDRAFTCOD(cnt) = OldEmpData2FileRec.DRAFTCOD
      TEMPDDACC(cnt) = OldEmpData2FileRec.EMPDDACC
      TPRENOTED(cnt) = OldEmpData2FileRec.PRENOTED
      TBANKNAME(cnt) = OldEmpData2FileRec.BANKNAME
      TBANKLOC(cnt) = OldEmpData2FileRec.BANKLOC
      TTRANSIT(cnt) = OldEmpData2FileRec.TRANSIT
    End If
    TEMPJOB(cnt) = OldEmpData2FileRec.EMPJOB
    If QPTrim$(OldEmpData2FileRec.EMPWCCLS) = "" Then
      TEMPWCCLS(cnt) = "0"
    Else
      TEMPWCCLS(cnt) = OldEmpData2FileRec.EMPWCCLS
    End If
    
    TEMPSTATS(cnt) = OldEmpData2FileRec.EMPSTATS
    If OldEmpData2FileRec.EMPBCODE < 0 Then
      TEMPBCODE(cnt) = 0
    Else
      TEMPBCODE(cnt) = OldEmpData2FileRec.EMPBCODE
    End If
    TEMPPTYPE(cnt) = OldEmpData2FileRec.EMPPTYPE
    TEMPPFREQ(cnt) = OldEmpData2FileRec.EMPPFREQ
    If OldEmpData2FileRec.EMPPRATE < 0 Then
      TEMPPRATE(cnt) = 0
    Else
      TEMPPRATE(cnt) = OldEmpData2FileRec.EMPPRATE
    End If
    If OldEmpData2FileRec.EMPORATE < 0 Then
      TEMPORATE(cnt) = 0
    Else
      TEMPORATE(cnt) = OldEmpData2FileRec.EMPORATE
    End If
    If Val(OldEmpData2FileRec.EMPHDATE) <= -10957 Then
      TEMPHDATE(cnt) = 0
    Else
      TEMPHDATE(cnt) = OldEmpData2FileRec.EMPHDATE
    End If
    
    If OldEmpData2FileRec.EMPRDATE < 0 Then
      TEMPRDATE(cnt) = 0
    Else
      TEMPRDATE(cnt) = OldEmpData2FileRec.EMPRDATE
    End If
    
    If OldEmpData2FileRec.EMPTDATE <= -3652 Then
      TEMPTDATE(cnt) = 0
    Else
      TEMPTDATE(cnt) = OldEmpData2FileRec.EMPTDATE
    End If
    If QPTrim$(OldEmpData2FileRec.EMPFEDX) = "" Then
      TEMPFEDX(cnt) = "N"
    Else
      TEMPFEDX(cnt) = OldEmpData2FileRec.EMPFEDX
    End If
    TEMPFEDO2(cnt) = OldEmpData2FileRec.EMPFEDO2
    
    'neutralize situation where there is no value in the Fed Amt/Pct
    'field but there is a value in the Fed figure field
    If Len(QPTrim$(OldEmpData2FileRec.EMPFEDX)) = 0 And OldEmpData2FileRec.EMPFEDO1 > 0 Then
      TEMPFEDO1(cnt) = 0
      GoTo BadFedAmtPct
    End If
    If OldEmpData2FileRec.EMPFEDO1 < -0 Then
      TEMPFEDO1(cnt) = 0
    Else
      TEMPFEDO1(cnt) = OldEmpData2FileRec.EMPFEDO1
    End If
BadFedAmtPct:
    If QPTrim$(OldEmpData2FileRec.EMPFEDS) = "" Then
      TEMPFEDS(cnt) = "S"
    Else
      TEMPFEDS(cnt) = OldEmpData2FileRec.EMPFEDS
    End If
    If OldEmpData2FileRec.EMPFEDA < 0 Then
      TEMPFEDA(cnt) = 0
    Else
      TEMPFEDA(cnt) = OldEmpData2FileRec.EMPFEDA
    End If
    If OldEmpData2FileRec.EMPFEDAA < 0 Then
      TEMPFEDAA(cnt) = 0
    Else
      TEMPFEDAA(cnt) = OldEmpData2FileRec.EMPFEDAA
    End If
    
    If QPTrim$(OldEmpData2FileRec.EMPSTAX) = "" Then
      TEMPSTAX(cnt) = "N"
    Else
      TEMPSTAX(cnt) = OldEmpData2FileRec.EMPSTAX
    End If
    TEMPSTAO2(cnt) = OldEmpData2FileRec.EMPSTAO2
    'neutralize situation where there is no value in the State Amt/Pct
    'field but there is a value in the State figure field
    If Len(QPTrim$(OldEmpData2FileRec.EMPSTAX)) = 0 And OldEmpData2FileRec.EMPSTAO1 > 0 Then
      TEMPSTAO1(cnt) = 0
      GoTo BadStaAmtPct
    End If
    
    If OldEmpData2FileRec.EMPSTAO1 < 0 Then
      TEMPSTAO1(cnt) = 0
    Else
      TEMPSTAO1(cnt) = OldEmpData2FileRec.EMPSTAO1
    End If
BadStaAmtPct:
    If QPTrim$(OldEmpData2FileRec.EMPSTAS) = "" Then
      TEMPSTAS(cnt) = "S"
    Else
      TEMPSTAS(cnt) = OldEmpData2FileRec.EMPSTAS
    End If
    If OldEmpData2FileRec.EMPSTAA < 0 Then
      TEMPSTAA(cnt) = 0
    Else
      TEMPSTAA(cnt) = OldEmpData2FileRec.EMPSTAA
    End If
    If OldEmpData2FileRec.EMPSTAAA < 0 Then
      TEMPSTAAA(cnt) = 0
    Else
      TEMPSTAAA(cnt) = OldEmpData2FileRec.EMPSTAAA
    End If
    If QPTrim$(OldEmpData2FileRec.EMPSOCX) = "" Then
      TEMPSOCX(cnt) = "N"
    Else
      TEMPSOCX(cnt) = OldEmpData2FileRec.EMPSOCX
    End If
    If QPTrim$(OldEmpData2FileRec.EMPMEDX) = "" Then
      TEMPMEDX(cnt) = "N"
    Else
      TEMPMEDX(cnt) = OldEmpData2FileRec.EMPMEDX
    End If
    TempEIC(cnt) = OldEmpData2FileRec.EMPEIC
    For Cnt2 = 1 To 12
      If OldEmpData2FileRec.EmpDed(Cnt2).DAmt < 0 Then
        TempDed(cnt, Cnt2).DAmt = 0
      Else
        TempDed(cnt, Cnt2).DAmt = OldEmpData2FileRec.EmpDed(Cnt2).DAmt
      End If
      If Len(OldEmpData2FileRec.EmpDed(Cnt2).DPct) > 0 And Val(OldEmpData2FileRec.EmpDed(Cnt2).DAmt) <= 0 Then
        TempDed(cnt, Cnt2).DPct = ""
      Else
        TempDed(cnt, Cnt2).DPct = OldEmpData2FileRec.EmpDed(Cnt2).DPct
      End If
    TempDed(cnt, Cnt2).DOTI = OldEmpData2FileRec.EmpDed(Cnt2).DOTI
    Next Cnt2
    TEMPEACT1(cnt) = OldEmpData2FileRec.EMPEACT1
    If OldEmpData2FileRec.EMPEAMT1 < 0 Then
      TEMPEAMT1(cnt) = 0
    Else
      TEMPEAMT1(cnt) = OldEmpData2FileRec.EMPEAMT1
    End If
    
    TEMPEACT2(cnt) = OldEmpData2FileRec.EMPEACT2
    If OldEmpData2FileRec.EMPEAMT2 < 0 Then
      TEMPEAMT2(cnt) = 0
    Else
      TEMPEAMT2(cnt) = OldEmpData2FileRec.EMPEAMT2
    End If
    TEMPEACT3(cnt) = OldEmpData2FileRec.EMPEACT3
    If OldEmpData2FileRec.EMPEAMT3 < -0 Then
      TEMPEAMT3(cnt) = 0
    Else
      TEMPEAMT3(cnt) = OldEmpData2FileRec.EMPEAMT3
    End If
    TEMPHP(cnt) = OldEmpData2FileRec.EMPHP
    For Cnt2 = 1 To 8
      TEDist(cnt, Cnt2).DAcct = OldEmpData2FileRec.EDist(Cnt2).DAcct
      If OldEmpData2FileRec.EDist(Cnt2).DAmt < 0 Then
        TEDist(cnt, Cnt2).DAmt = 0
      Else
        TEDist(cnt, Cnt2).DAmt = OldEmpData2FileRec.EDist(Cnt2).DAmt
      End If
    Next Cnt2
    
    If OldEmpData2FileRec.EMPVACE < -1000 Then
      TEMPVACE(cnt) = 0
    Else
      TEMPVACE(cnt) = OldEmpData2FileRec.EMPVACE
    End If
    If OldEmpData2FileRec.EMPVUSED < -1000 Then
      TEMPVUSED(cnt) = 0
    Else
      TEMPVUSED(cnt) = OldEmpData2FileRec.EMPVUSED
    End If
    If OldEmpData2FileRec.EMPVBAL < -1000 Then
      TEMPVBAL(cnt) = 0
    Else
      TEMPVBAL(cnt) = OldEmpData2FileRec.EMPVBAL
    End If
    
    If OldEmpData2FileRec.EMPSLE < -1000 Then
      TEMPSLE(cnt) = 0
    Else
      TEMPSLE(cnt) = OldEmpData2FileRec.EMPSLE
    End If
    If OldEmpData2FileRec.EMPSLUSE < -1000 Then
      TEMPSLUSE(cnt) = 0
    Else
      TEMPSLUSE(cnt) = OldEmpData2FileRec.EMPSLUSE
    End If
    If OldEmpData2FileRec.EMPSLBAL < -1000 Then
      TEMPSLBAL(cnt) = 0
    Else
      TEMPSLBAL(cnt) = OldEmpData2FileRec.EMPSLBAL
    End If
    
    If OldEmpData2FileRec.EMPCTE < -1000 Then
      TEMPCTE(cnt) = 0
    Else
      TEMPCTE(cnt) = OldEmpData2FileRec.EMPCTE
    End If
    If OldEmpData2FileRec.EMPCTUSE < -1000 Then
      TEMPCTUSE(cnt) = 0
    Else
      TEMPCTUSE(cnt) = OldEmpData2FileRec.EMPCTUSE
    End If
    If OldEmpData2FileRec.EMPCTBAL < -1000 Then
      TEMPCTBAL(cnt) = 0
    Else
      TEMPCTBAL(cnt) = OldEmpData2FileRec.EMPCTBAL
    End If
    
    If OldEmpData2FileRec.PERERN < -1000 Then
      TPERERN(cnt) = 0
    Else
      TPERERN(cnt) = OldEmpData2FileRec.PERERN
    End If
    If OldEmpData2FileRec.PerUsed < -1000 Then
      TPerUsed(cnt) = 0
    Else
      TPerUsed(cnt) = OldEmpData2FileRec.PerUsed
    End If
    If OldEmpData2FileRec.PERBAL < -1000 Then
      TPERBAL(cnt) = 0
    Else
      TPERBAL(cnt) = OldEmpData2FileRec.PERBAL
    End If
    
    If OldEmpData2FileRec.HOLERN < -1000 Then
      THOLERN(cnt) = 0
    Else
      THOLERN(cnt) = OldEmpData2FileRec.HOLERN
    End If
    If OldEmpData2FileRec.HolUsed < -1000 Then
      THolUsed(cnt) = 0
    Else
      THolUsed(cnt) = OldEmpData2FileRec.HolUsed
    End If
    If OldEmpData2FileRec.HOLBAL < -1000 Then
      THOLBAL(cnt) = 0
    Else
      THOLBAL(cnt) = OldEmpData2FileRec.HOLBAL
    End If
    TLeaveTbl(cnt) = OldEmpData2FileRec.LeaveTbl
    If OldEmpData2FileRec.ExcludeESC = "0" Then OldEmpData2FileRec.ExcludeESC = "N"
    TExcludeESC(cnt) = OldEmpData2FileRec.ExcludeESC
    TUseLife(cnt) = OldEmpData2FileRec.UseLife
    If OldEmpData2FileRec.EmpPin < 0 Then
      TEmpPin(cnt) = 0
    Else
      TEmpPin(cnt) = OldEmpData2FileRec.EmpPin
    End If
    TLastTransRec(cnt) = OldEmpData2FileRec.LastTransRec
    TDeleted(cnt) = OldEmpData2FileRec.Deleted
    TLDTDate(cnt) = OldEmpData2FileRec.LDTDate
    TCDTDate(cnt) = OldEmpData2FileRec.CDTDate
    TInprocFlag(cnt) = OldEmpData2FileRec.InprocFlag
    TUnused(cnt) = OldEmpData2FileRec.Unused
SkipMe:
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
    
    For Cnt2 = 1 To 12
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
    If TPERBAL(cnt) > 500000 Then TPERBAL(cnt) = 0
    NewEmpData2FileRec.PERBAL = TPERBAL(cnt)
    NewEmpData2FileRec.HOLERN = THOLERN(cnt)
    NewEmpData2FileRec.HolUsed = THolUsed(cnt)
    If THOLBAL(cnt) > 500000 Then THOLBAL(cnt) = 0
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
    If QPTrim$(TEMPRETTP(cnt)) = "" Then
      NewEmpData2FileRec.YN401K = "N"
    Else
      NewEmpData2FileRec.YN401K = "Y"
    End If
    NewEmpData2FileRec.PrimeDept = "1"
    NewEmpData2FileRec.HomePhone = "(000)-000-0000"
    NewEmpData2FileRec.EmrgncyCntctName = "Nothing Saved"
    NewEmpData2FileRec.EmrgncyCntctPhnNum = "(000)-000-0000"
    NewEmpData2FileRec.EmrgncyCntctRelation = "Nothing Saved"
    'added 9/1/04
    NewEmpData2FileRec.Comment = ""
    
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
  For X = 1 To 12
    TempUseDed(X) = OldPRDefaultRec.UseDed(X)
  Next X
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
  For X = 1 To 12
    NewPRDefaultRec.UseDed(X) = TempUseDed(X)
  Next X
  For X = 13 To 50
    NewPRDefaultRec.UseDed(X) = "N"
  Next X
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
  
  For X = 1 To OldDedCnt
    Get OldDedHandle, X, OldDedFileRec
    TempDCDESC1(X) = OldDedFileRec.DCDESC1
    TempDCACCT1(X) = OldDedFileRec.DCACCT1
    TempDCFWT1(X) = OldDedFileRec.DCFWT1
    TempDCSWT1(X) = OldDedFileRec.DCSWT1
    TempDCSOC1(X) = OldDedFileRec.DCSOC1
    TempDCMED1(X) = OldDedFileRec.DCMED1
  Next X
  Close OldDedHandle
  
  NewDedHandle = FreeFile
  DedRecLen = Len(NewDedFileRec)
  KillFile "prdata\" + "PRDEDCOD.DAT"
  Open "prdata\" + "PRDEDCOD.DAT" For Random Shared As NewDedHandle Len = DedRecLen
  
  Nextx = 1
  For X = 1 To OldDedCnt
    Get NewDedHandle, X, NewDedFileRec
    If QPTrim(TempDCDESC1(X)) = "" Then GoTo NoDesc
    NewDedFileRec.DCDESC1 = QPTrim$(TempDCDESC1(X))
    NewDedFileRec.DCACCT1 = QPTrim$(TempDCACCT1(X))
    NewDedFileRec.DCFWT1 = QPTrim$(TempDCFWT1(X))
    NewDedFileRec.DCSWT1 = QPTrim$(TempDCSWT1(X))
    NewDedFileRec.DCSOC1 = QPTrim$(TempDCSOC1(X))
    NewDedFileRec.DCMED1 = QPTrim$(TempDCMED1(X))
    Put NewDedHandle, Nextx, NewDedFileRec
    Nextx = Nextx + 1
NoDesc:
  Next X
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
    
  DosUnitRecLen = Len(UnitRec)
  DosFileHandle = FreeFile
  Open "PRData\PRUNIT.DAT" For Random Shared As DosFileHandle Len = DosUnitRecLen
  Get DosFileHandle, 1, DosUnitRec
  txtUpdate.Text = "Updating Unit Data"
  DoEvents

  If QPTrim$(DosUnitRec.BBTBATCH) = "" Then
    TempBBTBATCH = "0"
  Else
    TempBBTBATCH = QPTrim$(DosUnitRec.BBTBATCH)
  End If
  Close DosFileHandle
  
  If QPTrim$(DosUnitRec.BBTCNTNO) = "" Then
    TempBBTCNTNO = "0"
  Else
    TempBBTCNTNO = QPTrim$(DosUnitRec.BBTCNTNO)
  End If
  
  If Len(DosUnitRec.FileVer) = 0 Then
    TempFileVer = 0
  Else
     TempFileVer = DosUnitRec.FileVer
  End If
  
  If QPTrim$(DosUnitRec.UFCITY) = "" Then
    TempUFCITY = "YourTown"
  Else
    TempUFCITY = QPTrim$(DosUnitRec.UFCITY)
  End If
  
  If QPTrim$(DosUnitRec.UFEMPR) = "" Then
    TempUFEMPR = "YourTown"
  Else
    TempUFEMPR = QPTrim$(DosUnitRec.UFEMPR)
  End If
  
  If QPTrim$(DosUnitRec.UFFEDID) = "" Then
    TempUFFEDID = "0"
  Else
    TempUFFEDID = QPTrim$(DosUnitRec.UFFEDID)
  End If
  
  If QPTrim$(DosUnitRec.UFSTAID) = "" Then
    TempUFSTAID = "0"
  Else
    TempUFSTAID = QPTrim$(DosUnitRec.UFSTAID)
  End If
  
  If QPTrim$(DosUnitRec.UFRETID) = "" Then
    TempUFRETID = "0"
  Else
    TempUFRETID = QPTrim$(DosUnitRec.UFRETID)
  End If
  
  If QPTrim$(DosUnitRec.UFRETIDL) = "" Then
    TempUFRETIDL = "0"
  Else
    TempUFRETIDL = QPTrim$(DosUnitRec.UFRETIDL)
  End If
  
  If QPTrim$(DosUnitRec.IMMDNUM) = "" Then
    TempIMMDNUM = "0"
  Else
    TempIMMDNUM = QPTrim$(DosUnitRec.IMMDNUM)
  End If
  
  If QPTrim$(DosUnitRec.IMMONUM) = "" Then
    TempIMMONUM = "0"
  Else
    TempIMMONUM = QPTrim$(DosUnitRec.IMMONUM)
  End If
  
  If DosUnitRec.ESCRTYPE < 0 Then
    TempESCRTYPE = 0
  Else
    TempESCRTYPE = DosUnitRec.ESCRTYPE
  End If
  
  If DosUnitRec.TAXWBASE < 0 Then
    TempTAXWBASE = 0
  Else
    TempTAXWBASE = DosUnitRec.TAXWBASE
  End If
  
  If QPTrim$(DosUnitRec.UFSTATE) = "" Then
    TempUFSTATE = "NC"
  Else
    TempUFSTATE = QPTrim$(DosUnitRec.UFSTATE)
  End If
  
  UnitRecLen = Len(UnitRec)
  FileHandle = FreeFile
  Open "PRData\PRUNIT.DAT" For Random Shared As FileHandle Len = UnitRecLen
  Get FileHandle, 1, UnitRec
  UnitRec.BBTBATCH = TempBBTBATCH
  UnitRec.BBTCNTNO = TempBBTCNTNO
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
  UnitRec.LMT401YN = "N"
  UnitRec.FileVer = "Fall04"
  'added 9/1/04
  UnitRec.SSNOnCheck = "Y"
  UnitRec.GMatch401K = 0
  UnitRec.LMatch401K = 0
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
  
'  If Not Exist("PRDATA\PRTRANSH.DAT") Then GoTo NoTransHistorySaved 'for Canton install only
  
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

  For X = 1 To OLDTransHistRecCnt
    Get OldTransHistFileHandle, X, OldTransHistFileRec
    TempTActive(X) = OldTransHistFileRec.TActive
    TempPrevTransRec(X) = OldTransHistFileRec.PrevTransRec
    TempEmpPin(X) = OldTransHistFileRec.EmpPin
    TempPaySFlag(X) = OldTransHistFileRec.PaySFlag
    If OldTransHistFileRec.NetPay < 0 Then
      If OldTransHistFileRec.CheckNum > 0 Then
        OldTransHistFileRec.CheckNum = -OldTransHistFileRec.CheckNum
      End If
    End If
    TempCheckNum(X) = OldTransHistFileRec.CheckNum
    TempPayPdStart(X) = OldTransHistFileRec.PayPdStart
    TempPayPdEnd(X) = OldTransHistFileRec.PayPdEnd
    TempCheckDate(X) = OldTransHistFileRec.CheckDate
    TempPostDate(X) = OldTransHistFileRec.PostDate
    TempPayType(X) = OldTransHistFileRec.PayType
    TempBaseRate(X) = OldRound#(OldTransHistFileRec.BaseRate)
    TempOTRate(X) = OldRound#(OldTransHistFileRec.OTRate)
    If OldTransHistFileRec.RegHrsWork > 0 And OldTransHistFileRec.RegHrsWork < 1 Then
      OldTransHistFileRec.RegHrsWork = 0
    End If
    TempRegHrsWork(X) = OldRound#(OldTransHistFileRec.RegHrsWork)
    TempVacUsed(X) = OldRound#(OldTransHistFileRec.VacUsed)
    TempSickUsed(X) = OldRound#(OldTransHistFileRec.SickUsed)
    TempCompUsed(X) = OldRound#(OldTransHistFileRec.CompUsed)
    TempHOLHOURS(X) = OldRound#(OldTransHistFileRec.HOLHOURS)
    TempPerHours(X) = OldRound#(OldTransHistFileRec.PerHours)
    TempRegHrsPaid(X) = OldRound#(OldTransHistFileRec.RegHrsPaid)
    TempOTHours(X) = OldRound#(OldTransHistFileRec.OTHours)
    TempOTHrsPaid(X) = OldRound#(OldTransHistFileRec.OTHrsPaid)
    TempOT2COMP(X) = OldRound#(OldTransHistFileRec.OT2Comp)
    For y = 1 To 8
      OldTransHistFileRec.TDist(y).DOHrs = OldRound#(OldTransHistFileRec.TDist(y).DOHrs)
      OldTransHistFileRec.TDist(y).DOWage = OldRound#(OldTransHistFileRec.TDist(y).DOWage)
      OldTransHistFileRec.TDist(y).DPct = OldRound#(OldTransHistFileRec.TDist(y).DPct)
      OldTransHistFileRec.TDist(y).DRHrs = OldRound#(OldTransHistFileRec.TDist(y).DRHrs)
      OldTransHistFileRec.TDist(y).DRWage = OldRound#(OldTransHistFileRec.TDist(y).DRWage)
      TempTDist(X, y) = OldTransHistFileRec.TDist(y)
    Next y
    TempTotRegWage(X) = OldRound#(OldTransHistFileRec.TotRegWage)
    TempTotOTWage(X) = OldRound#(OldTransHistFileRec.TotOTWage)
    TempGrossWage(X) = OldRound#(OldTransHistFileRec.GrossWage)
    For y = 1 To 3
      TempEAmt(X, y) = OldRound#(OldTransHistFileRec.EAmt(y))
    Next y
    For y = 1 To 6
      OldTransHistFileRec.EDist(y).EAmt = OldRound#(OldTransHistFileRec.EDist(y).EAmt)
      TempEDist(X, y) = OldTransHistFileRec.EDist(y)
    Next y
    TempTotAdditEarn(X) = OldRound#(OldTransHistFileRec.TotAdditEarn)
    TempGrossPay(X) = OldRound#(OldTransHistFileRec.GrossPay)
    TempSocGrossPay(X) = OldRound#(OldTransHistFileRec.SocGrossPay)
    TempMedGrossPay(X) = OldRound#(OldTransHistFileRec.MedGrossPay)
    TempFedGrossPay(X) = OldRound#(OldTransHistFileRec.FedGrossPay)
    TempStaGrossPay(X) = OldRound#(OldTransHistFileRec.StaGrossPay)
    TempSocTaxAmt(X) = OldRound#(OldTransHistFileRec.SocTaxAmt)
    TempMedTaxAmt(X) = OldRound#(OldTransHistFileRec.MedTaxAmt)
    TempFedTaxAmt(X) = OldRound#(OldTransHistFileRec.FedTaxAmt)
    TempStaTaxAmt(X) = OldRound#(OldTransHistFileRec.StaTaxAmt)
    TempTotTaxAmt(X) = OldRound#(OldTransHistFileRec.TotTaxAmt)
    TempRetireAmt(X) = OldTransHistFileRec.RetireAmt
    For y = 1 To 12
      TempDAmt(X, y) = OldRound#(OldTransHistFileRec.DAmt(y))
    Next y
    TempTotDedAmt(X) = OldRound#(OldTransHistFileRec.TotDedAmt)
    TempEICAmt(X) = OldRound#(OldTransHistFileRec.EICAmt)
    TempNetPay(X) = OldRound#(OldTransHistFileRec.NetPay)
    TempPeriodHistRec(X) = OldTransHistFileRec.PeriodHistRec
    TempMatchRetAmt(X) = OldRound#(OldTransHistFileRec.MatchRetAmt)
    TempMatchSocAmt(X) = OldRound#(OldTransHistFileRec.MatchSocAmt)
    TempMatchMedAmt(X) = OldRound#(OldTransHistFileRec.MatchMedAmt)
    TempRetGrossPay(X) = OldRound#(OldTransHistFileRec.RetGrossPay)
    TempTaxFring(X) = OldRound#(OldTransHistFileRec.TaxFring)
  Next X

  Close OldTransHistFileHandle

  NewTransHistRecLen = Len(NewTransHistFileRec)
  NewTransHistFileHandle = FreeFile
  Open "prdata\" + "PRTRANSH.DAT" For Random Shared As NewTransHistFileHandle Len = NewTransHistRecLen

  For X = 1 To OLDTransHistRecCnt
    NewTransHistFileRec.TActive = TempTActive(X)
    NewTransHistFileRec.PrevTransRec = TempPrevTransRec(X)
    NewTransHistFileRec.EmpPin = TempEmpPin(X)
    NewTransHistFileRec.PaySFlag = TempPaySFlag(X)
    NewTransHistFileRec.CheckNum = TempCheckNum(X)
    NewTransHistFileRec.PayPdStart = TempPayPdStart(X)
    NewTransHistFileRec.PayPdEnd = TempPayPdEnd(X)
    NewTransHistFileRec.CheckDate = TempCheckDate(X)
    NewTransHistFileRec.PostDate = TempPostDate(X)
    NewTransHistFileRec.PayType = TempPayType(X)
    NewTransHistFileRec.BaseRate = TempBaseRate(X)
    NewTransHistFileRec.OTRate = TempOTRate(X)
    NewTransHistFileRec.RegHrsWork = TempRegHrsWork(X)
    NewTransHistFileRec.VacUsed = TempVacUsed(X)
    NewTransHistFileRec.SickUsed = TempSickUsed(X)
    NewTransHistFileRec.CompUsed = TempCompUsed(X)
    NewTransHistFileRec.HOLHOURS = TempHOLHOURS(X)
    NewTransHistFileRec.PerHours = TempPerHours(X)
    NewTransHistFileRec.RegHrsPaid = TempRegHrsPaid(X)
    NewTransHistFileRec.OTHours = TempOTHours(X)
    NewTransHistFileRec.OTHrsPaid = TempOTHrsPaid(X)
    NewTransHistFileRec.OT2Comp = TempOT2COMP(X)
    For y = 1 To 8
      NewTransHistFileRec.TDist(y) = TempTDist(X, y)
    Next y
    NewTransHistFileRec.TotRegWage = TempTotRegWage(X)
    NewTransHistFileRec.TotOTWage = TempTotOTWage(X)
    NewTransHistFileRec.GrossWage = TempGrossWage(X)
    For y = 1 To 3
      NewTransHistFileRec.EAmt(y) = TempEAmt(X, y)
    Next y
    For y = 1 To 6
      NewTransHistFileRec.EDist(y) = TempEDist(X, y)
    Next y
    NewTransHistFileRec.TotAdditEarn = TempTotAdditEarn(X)
    NewTransHistFileRec.GrossPay = TempGrossPay(X)
    NewTransHistFileRec.SocGrossPay = TempSocGrossPay(X)
    NewTransHistFileRec.MedGrossPay = TempMedGrossPay(X)
    NewTransHistFileRec.FedGrossPay = TempFedGrossPay(X)
    NewTransHistFileRec.StaGrossPay = TempStaGrossPay(X)
    NewTransHistFileRec.SocTaxAmt = TempSocTaxAmt(X)
    NewTransHistFileRec.MedTaxAmt = TempMedTaxAmt(X)
    NewTransHistFileRec.FedTaxAmt = TempFedTaxAmt(X)
    NewTransHistFileRec.StaTaxAmt = TempStaTaxAmt(X)
    NewTransHistFileRec.TotTaxAmt = TempTotTaxAmt(X)
    NewTransHistFileRec.RetireAmt = TempRetireAmt(X)
    For y = 1 To 12
      NewTransHistFileRec.DAmt(y) = TempDAmt(X, y)
    Next y
    For y = 13 To 50
      NewTransHistFileRec.DAmt(y) = 0#
    Next y
    
    NewTransHistFileRec.TotDedAmt = TempTotDedAmt(X)
    NewTransHistFileRec.EICAmt = TempEICAmt(X)
    NewTransHistFileRec.NetPay = TempNetPay(X)
    NewTransHistFileRec.PeriodHistRec = TempPeriodHistRec(X)
    NewTransHistFileRec.MatchRetAmt = TempMatchRetAmt(X)
    NewTransHistFileRec.MatchSocAmt = TempMatchSocAmt(X)
    NewTransHistFileRec.MatchMedAmt = TempMatchMedAmt(X)
    NewTransHistFileRec.RetGrossPay = TempRetGrossPay(X)
    NewTransHistFileRec.TaxFring = TempTaxFring(X)
    For y = 1 To 3
      NewTransHistFileRec.Less401k(y) = False
    Next y
    NewTransHistFileRec.Pad1 = ""
    Put NewTransHistFileHandle, X, NewTransHistFileRec
  Next X

  Close NewTransHistFileHandle

'NoTransHistorySaved: For Canton install only
  OldTransWorkRecLen = Len(OldTransWorkFileRec)
  OldTransWorkFileHandle = FreeFile
  Open "prdata\" + "PRTRANST.DAT" For Random Shared As OldTransWorkFileHandle Len = OldTransWorkRecLen
  OLDTransWorkRecCnt = LOF(OldTransWorkFileHandle) / Len(OldTransWorkFileRec)
  Dim TryThis As Double
  For X = 1 To OLDTransWorkRecCnt
    Get OldTransWorkFileHandle, X, OldTransWorkFileRec
    TempTActive(X) = 0 'OldTransWorkFileRec.TActive
    TempPrevTransRec(X) = OldTransWorkFileRec.PrevTransRec
    TempEmpPin(X) = OldTransWorkFileRec.EmpPin
    TempPaySFlag(X) = OldTransWorkFileRec.PaySFlag
    TempCheckNum(X) = OldTransWorkFileRec.CheckNum
    TempPayPdStart(X) = OldTransWorkFileRec.PayPdStart
    TempPayPdEnd(X) = OldTransWorkFileRec.PayPdEnd
    TempCheckDate(X) = OldTransWorkFileRec.CheckDate
    TempPostDate(X) = OldTransWorkFileRec.PostDate
    TempPayType(X) = OldTransWorkFileRec.PayType
    TempBaseRate(X) = OldTransWorkFileRec.BaseRate
    TempOTRate(X) = OldTransWorkFileRec.OTRate
    TempRegHrsWork(X) = OldTransWorkFileRec.RegHrsWork
    TempVacUsed(X) = OldTransWorkFileRec.VacUsed
    TempSickUsed(X) = OldTransWorkFileRec.SickUsed
    TempCompUsed(X) = OldTransWorkFileRec.CompUsed
    TempHOLHOURS(X) = OldTransWorkFileRec.HOLHOURS
    TempPerHours(X) = OldTransWorkFileRec.PerHours
    TempRegHrsPaid(X) = OldTransWorkFileRec.RegHrsPaid
    TempOTHours(X) = OldTransWorkFileRec.OTHours
    TempOTHrsPaid(X) = OldTransWorkFileRec.OTHrsPaid
    TempOT2COMP(X) = OldTransWorkFileRec.OT2Comp
    For y = 1 To 8
      TempTDist(X, y) = OldTransWorkFileRec.TDist(y)
    Next y
    TempTotRegWage(X) = OldTransWorkFileRec.TotRegWage
    TempTotOTWage(X) = OldTransWorkFileRec.TotOTWage
    TempGrossWage(X) = OldTransWorkFileRec.GrossWage
    For y = 1 To 3
      TempEAmt(X, y) = OldTransWorkFileRec.EAmt(y)
    Next y
    For y = 1 To 6
      TempEDist(X, y) = OldTransWorkFileRec.EDist(y)
    Next y
    TempTotAdditEarn(X) = OldTransWorkFileRec.TotAdditEarn
    TempGrossPay(X) = OldTransWorkFileRec.GrossPay
    TempSocGrossPay(X) = OldTransWorkFileRec.SocGrossPay
    TempMedGrossPay(X) = OldTransWorkFileRec.MedGrossPay
    TempFedGrossPay(X) = OldTransWorkFileRec.FedGrossPay
    TempStaGrossPay(X) = OldTransWorkFileRec.StaGrossPay
    TempSocTaxAmt(X) = OldTransWorkFileRec.SocTaxAmt
    TempMedTaxAmt(X) = OldTransWorkFileRec.MedTaxAmt
    TempFedTaxAmt(X) = OldTransWorkFileRec.FedTaxAmt
    TempStaTaxAmt(X) = OldTransWorkFileRec.StaTaxAmt
    TempTotTaxAmt(X) = OldTransWorkFileRec.TotTaxAmt
    TempRetireAmt(X) = OldTransWorkFileRec.RetireAmt
    For y = 1 To 12
      TempDAmt(X, y) = OldTransWorkFileRec.DAmt(y)
    Next y
    TempTotDedAmt(X) = OldTransWorkFileRec.TotDedAmt
    TempEICAmt(X) = OldTransWorkFileRec.EICAmt
    TempNetPay(X) = OldTransWorkFileRec.NetPay
    TempPeriodHistRec(X) = OldTransWorkFileRec.PeriodHistRec
    TempMatchRetAmt(X) = OldTransWorkFileRec.MatchRetAmt
    TempMatchSocAmt(X) = OldTransWorkFileRec.MatchSocAmt
    TempMatchMedAmt(X) = OldTransWorkFileRec.MatchMedAmt
    TempRetGrossPay(X) = OldTransWorkFileRec.RetGrossPay
    TempTaxFring(X) = OldTransWorkFileRec.TaxFring
  Next X

  Close OldTransWorkFileHandle

  NewTransWorkRecLen = Len(NewTransWorkFileRec)
  NewTransWorkFileHandle = FreeFile
  Open "prdata\" + "PRTRANST.DAT" For Random Shared As NewTransWorkFileHandle Len = NewTransWorkRecLen

  For X = 1 To OLDTransWorkRecCnt
    NewTransWorkFileRec.TActive = TempTActive(X)
    NewTransWorkFileRec.PrevTransRec = TempPrevTransRec(X)
    NewTransWorkFileRec.EmpPin = TempEmpPin(X)
    NewTransWorkFileRec.PaySFlag = TempPaySFlag(X)
    NewTransWorkFileRec.CheckNum = TempCheckNum(X)
    NewTransWorkFileRec.PayPdStart = TempPayPdStart(X)
    NewTransWorkFileRec.PayPdEnd = TempPayPdEnd(X)
    NewTransWorkFileRec.CheckDate = TempCheckDate(X)
    NewTransWorkFileRec.PostDate = TempPostDate(X)
    NewTransWorkFileRec.PayType = TempPayType(X)
    NewTransWorkFileRec.BaseRate = TempBaseRate(X)
    NewTransWorkFileRec.OTRate = TempOTRate(X)
    NewTransWorkFileRec.RegHrsWork = TempRegHrsWork(X)
    NewTransWorkFileRec.VacUsed = TempVacUsed(X)
    NewTransWorkFileRec.SickUsed = TempSickUsed(X)
    NewTransWorkFileRec.CompUsed = TempCompUsed(X)
    NewTransWorkFileRec.HOLHOURS = TempHOLHOURS(X)
    NewTransWorkFileRec.PerHours = TempPerHours(X)
    NewTransWorkFileRec.RegHrsPaid = TempRegHrsPaid(X)
    NewTransWorkFileRec.OTHours = TempOTHours(X)
    NewTransWorkFileRec.OTHrsPaid = TempOTHrsPaid(X)
    NewTransWorkFileRec.OT2Comp = TempOT2COMP(X)
    For y = 1 To 8
      NewTransWorkFileRec.TDist(y) = TempTDist(X, y)
    Next y
    NewTransWorkFileRec.TotRegWage = TempTotRegWage(X)
    NewTransWorkFileRec.TotOTWage = TempTotOTWage(X)
    NewTransWorkFileRec.GrossWage = TempGrossWage(X)
    For y = 1 To 3
      NewTransWorkFileRec.EAmt(y) = TempEAmt(X, y)
    Next y
    For y = 1 To 6
      NewTransWorkFileRec.EDist(y) = TempEDist(X, y)
    Next y
    NewTransWorkFileRec.TotAdditEarn = TempTotAdditEarn(X)
    NewTransWorkFileRec.GrossPay = TempGrossPay(X)
    NewTransWorkFileRec.SocGrossPay = TempSocGrossPay(X)
    NewTransWorkFileRec.MedGrossPay = TempMedGrossPay(X)
    NewTransWorkFileRec.FedGrossPay = TempFedGrossPay(X)
    NewTransWorkFileRec.StaGrossPay = TempStaGrossPay(X)
    NewTransWorkFileRec.SocTaxAmt = TempSocTaxAmt(X)
    NewTransWorkFileRec.MedTaxAmt = TempMedTaxAmt(X)
    NewTransWorkFileRec.FedTaxAmt = TempFedTaxAmt(X)
    NewTransWorkFileRec.StaTaxAmt = TempStaTaxAmt(X)
    NewTransWorkFileRec.TotTaxAmt = TempTotTaxAmt(X)
    NewTransWorkFileRec.RetireAmt = TempRetireAmt(X)
    For y = 1 To 12
      NewTransWorkFileRec.DAmt(y) = TempDAmt(X, y)
    Next y
    
    For y = 13 To 50
      NewTransWorkFileRec.DAmt(y) = 0#
    Next y
      
    NewTransWorkFileRec.TotDedAmt = TempTotDedAmt(X)
    NewTransWorkFileRec.EICAmt = TempEICAmt(X)
    NewTransWorkFileRec.NetPay = TempNetPay(X)
    NewTransWorkFileRec.PeriodHistRec = TempPeriodHistRec(X)
    NewTransWorkFileRec.MatchRetAmt = TempMatchRetAmt(X)
    NewTransWorkFileRec.MatchSocAmt = TempMatchSocAmt(X)
    NewTransWorkFileRec.MatchMedAmt = TempMatchMedAmt(X)
    NewTransWorkFileRec.RetGrossPay = TempRetGrossPay(X)
    NewTransWorkFileRec.TaxFring = TempTaxFring(X)
    For y = 1 To 3
      NewTransWorkFileRec.Less401k(y) = False '9/26/03
    Next y
    NewTransWorkFileRec.Pad1 = ""
    Put NewTransWorkFileHandle, X, NewTransWorkFileRec
  Next X
  
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
  
  For X = 1 To NumOfChkRecs
    Get OldChecksHandle, X, OldChkRec
    TempCActive(X) = OldChkRec.CActive
    TempCheckNum(X) = OldChkRec.CheckNum
    TempCheckDate(X) = OldChkRec.CheckDate
    TempEmpName(X) = OldChkRec.EmpName
    TempEmpNo(X) = OldChkRec.EmpNo
    TempEmpSSN(X) = OldChkRec.EmpSSN
    TempEmpAddr1(X) = OldChkRec.EmpAddr1
    TempEmpCity(X) = OldChkRec.EmpCity
    TempEmpState(X) = OldChkRec.EmpState
    TempEmpZip(X) = OldChkRec.EmpZip
    TempPayEndDate(X) = OldChkRec.PayEndDate
    TempBaseRate(X) = OldChkRec.BaseRate
    TempGrossPay(X) = OldChkRec.GrossPay
    TempFedTaxAmt(X) = OldChkRec.FedTaxAmt
    TempStaTaxAmt(X) = OldChkRec.StaTaxAmt
    TempMedTaxAmt(X) = OldChkRec.MedTaxAmt
    TempSocTaxAmt(X) = OldChkRec.SocTaxAmt
    TempTotDedAmt(X) = OldChkRec.TotDedAmt
    TempRetireAmt(X) = OldChkRec.RetireAmt
    TempNetPay(X) = OldChkRec.NetPay
    TempYTDGrossPay(X) = OldChkRec.YTDGrossPay
    TempYTDFederal(X) = OldChkRec.YTDFederal
    TempYTDState(X) = OldChkRec.YTDState
    TempYTDSocial(X) = OldChkRec.YTDSocial
    TempYTDMedicare(X) = OldChkRec.YTDMedicare
    TempYTDTotDed(X) = OldChkRec.YTDTotDed
    TempYTDNetPay(X) = OldChkRec.YTDNetPay
    TempYTDRetire(X) = OldChkRec.YTDRetire
    TempVactBal(X) = OldChkRec.VactBal
    TempSickBal(X) = OldChkRec.SickBal
    TempCompBal(X) = OldChkRec.CompBal
    TempCompEarn(X) = OldChkRec.CompEarn
    TempRegHrsWork(X) = OldChkRec.RegHrsWork
    TempOTHrsPaid(X) = OldChkRec.OTHrsPaid
    TempTotRegWage(X) = OldChkRec.TotRegWage
    TempVacUsed(X) = OldChkRec.VacUsed
    TempSickUsed(X) = OldChkRec.SickUsed
    TempCompUsed(X) = OldChkRec.CompUsed
    TempHolUsed(X) = OldChkRec.HolUsed
    TempPerUsed(X) = OldChkRec.PerUsed
    TempRegHrsPaid(X) = OldChkRec.RegHrsPaid
    TempTotOTWage(X) = OldChkRec.TotOTWage
    For y = 1 To 3
      TempAEarn(X, y).DAmt = OldChkRec.AEarn(y).DAmt
      TempAEarn(X, y).DCode = OldChkRec.AEarn(y).DCode
      TempAEarn(X, y).YTDDAmt = OldChkRec.AEarn(y).YTDDAmt
    Next y
    TempTotAdditEarn(X) = OldChkRec.TotAdditEarn
    TempEICAmt(X) = OldChkRec.EICAmt
    TempTaxFring(X) = OldChkRec.TaxFring
    For y = 1 To 12
      TempCDED(X, y).DAmt = OldChkRec.CDED(y).DAmt
      TempCDED(X, y).DCode = OldChkRec.CDED(y).DCode
      TempCDED(X, y).YTDDAmt = OldChkRec.CDED(y).YTDDAmt
    Next y
    TempDDFlag(X) = OldChkRec.DDFlag
  Next X
  
  Close OldChecksHandle
  
  NewChecksHandle = FreeFile
  Open "prdata\" + "PRCHECKS.DAT" For Random Shared As NewChecksHandle Len = NewChecksRecLen
  
  For X = 1 To NumOfChkRecs
    Get NewChecksHandle, X, NewChkRec
    NewChkRec.CActive = TempCActive(X)
    NewChkRec.CheckNum = TempCheckNum(X)
    NewChkRec.CheckDate = TempCheckDate(X)
    NewChkRec.EmpName = TempEmpName(X)
    NewChkRec.EmpNo = TempEmpNo(X)
    NewChkRec.EmpSSN = TempEmpSSN(X)
    NewChkRec.EmpAddr1 = TempEmpAddr1(X)
    NewChkRec.EmpCity = TempEmpCity(X)
    NewChkRec.EmpState = TempEmpState(X)
    NewChkRec.EmpZip = TempEmpZip(X)
    NewChkRec.PayEndDate = TempPayEndDate(X)
    NewChkRec.BaseRate = TempBaseRate(X)
    NewChkRec.GrossPay = TempGrossPay(X)
    NewChkRec.FedTaxAmt = TempFedTaxAmt(X)
    NewChkRec.StaTaxAmt = TempStaTaxAmt(X)
    NewChkRec.MedTaxAmt = TempMedTaxAmt(X)
    NewChkRec.SocTaxAmt = TempSocTaxAmt(X)
    NewChkRec.TotDedAmt = TempTotDedAmt(X)
    NewChkRec.RetireAmt = TempRetireAmt(X)
    NewChkRec.NetPay = TempNetPay(X)
    NewChkRec.YTDGrossPay = TempYTDGrossPay(X)
    NewChkRec.YTDFederal = TempYTDFederal(X)
    NewChkRec.YTDState = TempYTDState(X)
    NewChkRec.YTDSocial = TempYTDSocial(X)
    NewChkRec.YTDMedicare = TempYTDMedicare(X)
    NewChkRec.YTDTotDed = TempYTDTotDed(X)
    NewChkRec.YTDNetPay = TempYTDNetPay(X)
    NewChkRec.YTDRetire = TempYTDRetire(X)
    NewChkRec.VactBal = TempVactBal(X)
    NewChkRec.SickBal = TempSickBal(X)
    NewChkRec.CompBal = TempCompBal(X)
    NewChkRec.CompEarn = TempCompEarn(X)
    NewChkRec.RegHrsWork = TempRegHrsWork(X)
    NewChkRec.OTHrsPaid = TempOTHrsPaid(X)
    NewChkRec.TotRegWage = TempTotRegWage(X)
    NewChkRec.VacUsed = TempVacUsed(X)
    NewChkRec.SickUsed = TempSickUsed(X)
    NewChkRec.CompUsed = TempCompUsed(X)
    NewChkRec.HolUsed = TempHolUsed(X)
    NewChkRec.PerUsed = TempPerUsed(X)
    NewChkRec.RegHrsPaid = TempRegHrsPaid(X)
    NewChkRec.TotOTWage = TempTotOTWage(X)
    For y = 1 To 3
      NewChkRec.AEarn(y).DAmt = TempAEarn(X, y).DAmt
      NewChkRec.AEarn(y).DCode = TempAEarn(X, y).DCode
      NewChkRec.AEarn(y).YTDDAmt = TempAEarn(X, y).YTDDAmt
    Next y
    NewChkRec.TotAdditEarn = TempTotAdditEarn(X)
    NewChkRec.EICAmt = TempEICAmt(X)
    NewChkRec.TaxFring = TempTaxFring(X)
    For y = 1 To 12
      NewChkRec.CDED(y).DAmt = TempCDED(X, y).DAmt
      NewChkRec.CDED(y).DCode = TempCDED(X, y).DCode
      NewChkRec.CDED(y).YTDDAmt = TempCDED(X, y).YTDDAmt
    Next y
    For y = 13 To 50
      NewChkRec.CDED(y).DAmt = 0
      NewChkRec.CDED(y).DCode = ""
      NewChkRec.CDED(y).YTDDAmt = 0
    Next y
    NewChkRec.DDFlag = TempDDFlag(X)
    
    Put NewChecksHandle, X, NewChkRec
  Next X
  Close NewChecksHandle
  
  
  Dim TempExtrYear As Integer
  Dim TempDeds(0 To 50) As W2DedType
  
  OldW2Handle = FreeFile
  Open "prdata\" + "PRW2SETU.DAT" For Random Shared As OldW2Handle Len = OldW2RecLen
  
  Get OldW2Handle, 1, OldW2SetUp
  TempExtrYear = OldW2SetUp.ExtrYear
  For X = 0 To 12
    TempDeds(X).AMTBOX = OldW2SetUp.Deds(X).AMTBOX
    TempDeds(X).CHKDED = OldW2SetUp.Deds(X).CHKDED
    TempDeds(X).DedCode = OldW2SetUp.Deds(X).DedCode
  Next X
  For X = 13 To 50
    TempDeds(X).AMTBOX = ""
    TempDeds(X).CHKDED = ""
    TempDeds(X).DedCode = ""
  Next X
  
  Close OldW2Handle
  
  NewW2Handle = FreeFile
  Open "prdata\" + "PRW2SETU.DAT" For Random Shared As NewW2Handle Len = NewW2RecLen
  NewW2SetUp.ExtrYear = TempExtrYear
  For X = 0 To 50
    NewW2SetUp.Deds(X).AMTBOX = TempDeds(X).AMTBOX
    NewW2SetUp.Deds(X).CHKDED = TempDeds(X).CHKDED
    NewW2SetUp.Deds(X).DedCode = TempDeds(X).DedCode
  Next X
  
  Put NewW2Handle, 1, NewW2SetUp
  Close NewW2Handle
  
  Dim DosErnHandle As Integer
  Dim DosErnRec As DosErnCodeRecType
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  
  OpenDosErnCodeFile DosErnHandle
  NumOfErns = LOF(DosErnHandle) / Len(DosErnRec)
  If NumOfErns = 0 Then
    GoTo NoErnMatchNeeded
  End If
  
  ReDim TempERNCODE1(1 To NumOfErns) As String * 10
  ReDim TempERNFWT1(1 To NumOfErns) As String * 1
  ReDim TempERNSWT1(1 To NumOfErns) As String * 1
  ReDim TempERNSOC1(1 To NumOfErns) As String * 1
  ReDim TempERNMED1(1 To NumOfErns) As String * 1
  ReDim TempERNRET1(1 To NumOfErns) As String * 1
  For X = 1 To NumOfErns
    Get DosErnHandle, X, DosErnRec
    TempERNCODE1(X) = QPTrim$(DosErnRec.ERNCODE1)
    TempERNFWT1(X) = DosErnRec.ERNFWT1
    TempERNSWT1(X) = DosErnRec.ERNSWT1
    TempERNSOC1(X) = DosErnRec.ERNSOC1
    TempERNMED1(X) = DosErnRec.ERNMED1
    TempERNRET1(X) = DosErnRec.ERNRET1
  Next X
  Close DosErnHandle
  
  OpenErnCodeFile ErnHandle
  For X = 1 To NumOfErns
    ErnRec.ERNCODE1 = QPTrim$(TempERNCODE1(X))
    ErnRec.ERNFWT1 = TempERNFWT1(X)
    ErnRec.ERNSWT1 = TempERNSWT1(X)
    ErnRec.ERNSOC1 = TempERNSOC1(X)
    ErnRec.ERNMED1 = TempERNMED1(X)
    ErnRec.ERNRET1 = TempERNRET1(X)
    ErnRec.EarnYN = "Y" '"Y"es Include
    ErnRec.Pad = ""
    Put ErnHandle, X, ErnRec
  Next X
  Close ErnHandle
NoErnMatchNeeded:
  txtUpdate.Visible = False
  cmdCancel.Enabled = True
  
  Call InsertVoidFieldsIntoTrans
  
  txtUpdate.Text = "UPDATE COMPLETE. PRESS OK TO GO TO PAYROLL MAIN MENU"
  cmdOK.Visible = True
  cmdConvert.Visible = False
  cmdCancel.Visible = False
  
  Exit Sub
ErrorHandler:
  frmWarnReloadPRData.Show vbModal
  Call cmdOK_Click
End Sub

