VERSION 5.00
Begin VB.Form frmBankReconMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconciliation "
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmBankRecon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrintUncanceled 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Print &Uncanceled Checks List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   3612
   End
   Begin VB.CommandButton cmdSelectChkCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Select Checks to Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   3612
   End
   Begin VB.CommandButton cmdRemoveCancelChks 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Remove Canceled Checks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   3612
   End
   Begin VB.CommandButton cmdAddOutstandCks 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Add Outstanding Checks to File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   3612
   End
   Begin VB.CommandButton cmdSortOutstandChks 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Sor&t Outstanding Checks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      HelpContextID   =   36
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintCancelChks 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Print &Canceled Checks List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      HelpContextID   =   33
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitBankReconMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Bank Reconciliation Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   3612
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BANK RECONCILIATION MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   7092
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmBankReconMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim GLSetup As GLSetupRecType

Private Sub cmdAddOutstandCks_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Not Exist("APCHKINF.DAT") Then
    If Exist("crchek.opn") Then
      FileHandle = FreeFile
      Open "crchek.opn" For Input As FileHandle
      Line Input #FileHandle, WhosOnFirst$
      Close FileHandle
      MsgBox "The Check Reconciliation File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
    Else
      FileHandle = FreeFile
      Open "crchek.opn" For Output As FileHandle
      Print #FileHandle, ComputerName$
      Close FileHandle
      frmChkRecAdd.Show
      Unload frmBankReconMenu
    End If
  Else
    MsgBox "An UnPosted AP Check File Exists, Please Wait Until Checks Are Posted Before Editing The Check Reconciliation File.", vbOKOnly, "Please Wait"
  End If
End Sub

Private Sub cmdExitBankReconMenu_Click()
  frmGLMainMenu.Show
  Unload frmBankReconMenu
End Sub

Private Sub cmdPrintCancelChks_Click()
  frmReportOpt.Show 1
  FrmShowPctComp.Label1 = "Creating Canceled Check Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmBankReconMenu
  If rptopt = 1 Then
    PrintChkRecList 1, frmBankReconMenu, , , 999
  ElseIf rptopt = 2 Then
    PrintChkRecList2 1, frmBankReconMenu, , , 999
  Else
    Unload FrmShowPctComp
    ActivateControls frmBankReconMenu
    FrmShowPctComp.Out = False
  End If
End Sub

Private Sub cmdPrintUncanceled_Click()
  frmPrnOutstandCks.Show
End Sub

Private Sub cmdRemoveCancelChks_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Not Exist("APCHKINF.DAT") Then
    If Exist("crchek.opn") Then
      FileHandle = FreeFile
      Open "crchek.opn" For Input As FileHandle
      Line Input #FileHandle, WhosOnFirst$
      Close FileHandle
      MsgBox "The Check Reconciliation File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
    Else
      FileHandle = FreeFile
      Open "crchek.opn" For Output As FileHandle
      Print #FileHandle, ComputerName$
      Close FileHandle
      frmRemCanChks.Show
      Unload frmBankReconMenu
    End If
  Else
    MsgBox "An UnPosted AP Check File Exists, Please Wait Until Checks Are Posted Before Editing The Check Reconciliation File.", vbOKOnly, "Please Wait"
  End If

End Sub

Private Sub cmdSelectChkCancel_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Not Exist("APCHKINF.DAT") Then
    If Exist("crchek.opn") Then
      FileHandle = FreeFile
      Open "crchek.opn" For Input As FileHandle
      Line Input #FileHandle, WhosOnFirst$
      Close FileHandle
      MsgBox "The Check Reconciliation File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
    Else
      FileHandle = FreeFile
      Open "crchek.opn" For Output As FileHandle
      Print #FileHandle, ComputerName$
      Close FileHandle
      frmChkRecCancel.Show
      Unload frmBankReconMenu
    End If
  Else
    MsgBox "An UnPosted AP Check File Exists, Please Wait Until Checks Are Posted Before Editing The Check Reconciliation File.", vbOKOnly, "Please Wait"
  End If
End Sub

Private Sub cmdSortOutstandChks_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Not Exist("APCHKINF.DAT") Then
    If Exist("crchek.opn") Then
      FileHandle = FreeFile
      Open "crchek.opn" For Input As FileHandle
      Line Input #FileHandle, WhosOnFirst$
      Close FileHandle
      MsgBox "The Check Reconciliation File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
    Else
      FileHandle = FreeFile
      Open "crchek.opn" For Output As FileHandle
      Print #FileHandle, ComputerName$
      Close FileHandle
      SortCheckFile
      Kill "crchek.opn"
      MsgBox "Outstanding checks sorted.", vbOKOnly, "Complete"
    End If
  Else
    MsgBox "An UnPosted AP Check File Exists, Please Wait Until Checks Are Posted Before Editing The Check Reconciliation File.", vbOKOnly, "Please Wait"
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitBankReconMenu_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpBankReconciliation
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitBankReconMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Public Sub PrintChkRecList(WhatKind%, formname As Form, Optional datethru As Integer, Optional Src As Integer, Optional bankc As Integer)
  Dim ReportFile As String, RptTitle As String, SrchDate As Integer
  Dim ChkType As Integer, BankNum As Integer, TotalChks As Double
  Dim NumChks As Integer, RecLen As Integer, FileNum As Integer
  Dim NumRecs As Integer, PRNFileNum As Integer, cnt As Integer, Newrp As String
  Dim T As String, ToPrint As String
  Dim OSChek As OSChekRecType
  Newrp = "CR"
  GetRPTName Newrp
  ReportFile$ = Newrp
  If WhatKind = 0 Then
    RptTitle$ = "Outstanding Checks Report"
    SrchDate = datethru
    If Src = 2 Then
      ChkType = 3
    Else
      ChkType = Src
    End If
    If bankc = 0 Then
      BankNum = 999
    Else
      BankNum = bankc
    End If
  Else
    RptTitle$ = " Canceled Checks Report  "
    SrchDate = 32767
    BankNum = 999
    ChkType = 3
  End If
  GoTo ProcessReport

ProcessReport:

   'Report is sent to the following file which is passed to fileview for
   'screen output or printed using the BLPrint routine

   TotalChks# = 0  'Problem in totals using single precision
   NumChks = 0

   RecLen = Len(OSChek)
   FileNum = FreeFile
   Open "crchek.dat" For Random Access Read Write Shared As FileNum Len = RecLen
   NumRecs = LOF(FileNum) \ RecLen
   PRNFileNum = FreeFile
   Open ReportFile$ For Output As #PRNFileNum

   'ShowProcessingScrn RptTitle$
   For cnt = 1 To NumRecs
      FrmShowPctComp.ShowPctComp cnt, NumRecs
      If FrmShowPctComp.Out = True Then
        Close
        Unload FrmShowPctComp
        ActivateControls frmBankReconMenu
        FrmShowPctComp.Out = False
        GoTo CancelExit
      End If

      Get FileNum, cnt, OSChek
       '''' If OSChek.ChkNum = 23193 Then Stop
        If ChkType = 3 Then GoTo jumpall
        If ChkType = 1 Then BankNum = 999

        If OSChek.Src = ChkType Then

jumpall:
        If OSChek.Bankcode = BankNum Or BankNum = 999 Then
          If OSChek.Cleared = WhatKind Then

            If OSChek.chkdate <= SrchDate Then

               NumChks = NumChks + 1
           
               TotalChks# = Round#(TotalChks# + OSChek.Amt)
               Select Case OSChek.Src
                 Case 0
                   T$ = "AP"
                 Case 1
                   T$ = "PR"
               End Select

               ToPrint$ = ""
               ToPrint$ = Str$(OSChek.ChkNum)
               ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", OSChek.chkdate, "12-31-1979"), "mm/dd/yy")
               ToPrint$ = ToPrint$ + "~" + T$
               ToPrint$ = ToPrint$ + "~" + QPTrim(OSChek.Desc)
               ToPrint$ = ToPrint$ + "~" + Using$("##,###,###.##", Str$(OSChek.Amt))
               ToPrint$ = ToPrint$ + "~" + Str$(OSChek.Bankcode)
               Print #PRNFileNum, ToPrint$
            End If

         End If
       End If
     End If
   Next
  Close
  Load frmLoadingRpt
  ActivateControls frmBankReconMenu
  ARptOutstChks.totcks = Using$("#####", Str$(NumChks))
  ARptOutstChks.totamt = Using$("##,###,###.##", Str$(TotalChks#))
  ARptOutstChks.txtDate = Now
  ARptOutstChks.txtTown = GLUserName$
  ARptOutstChks.Title.Caption = RptTitle$
  ARptOutstChks.GetName ReportFile$
  ARptOutstChks.startrpt

   
  ' ViewPrint ReportFile$, RptTitle$
'Kill ReportFile$     'Clean up after ourselves

ChkListExit:
Exit Sub
Getout:
'UNLOCK FileNum
Close
Exit Sub
CancelExit:
Exit Sub
End Sub
Public Sub PrintChkRecList2(WhatKind%, formname As Form, Optional datethru As Integer, Optional Src As Integer, Optional bankc As Integer)
  Dim ReportFile As String, RptTitle As String, SrchDate As Integer
  Dim ChkType As Integer, BankNum As Integer, TotalChks As Double
  Dim NumChks As Integer, RecLen As Integer, FileNum As Integer
  Dim NumRecs As Integer, PRNFileNum As Integer, cnt As Integer, Newrp As String
  Dim T As String, ToPrint As String
  Dim OSChek As OSChekRecType
  Newrp = "CR"
  GetRPTName Newrp
  ReportFile$ = Newrp
  If WhatKind = 0 Then
    RptTitle$ = "Outstanding Checks Report"
    SrchDate = datethru
    If Src = 2 Then
      ChkType = 3
    Else
      ChkType = Src
    End If
    If bankc = 0 Then
      BankNum = 999
    Else
      BankNum = bankc
    End If
  Else
    RptTitle$ = " Canceled Checks Report  "
    SrchDate = 32767
    BankNum = 999
    ChkType = 3
  End If
  GoTo ProcessReport

ProcessReport:

   'Report is sent to the following file which is passed to fileview for
   'screen output or printed using the BLPrint routine

   TotalChks# = 0  'Problem in totals using single precision
   NumChks = 0

   RecLen = Len(OSChek)
   FileNum = FreeFile
   Open "crchek.dat" For Random Access Read Write Shared As FileNum Len = RecLen
   NumRecs = LOF(FileNum) \ RecLen
   PRNFileNum = FreeFile
   Open ReportFile$ For Output As #PRNFileNum

   'ShowProcessingScrn RptTitle$
   GoSub DoHeading
   For cnt = 1 To NumRecs
      FrmShowPctComp.ShowPctComp cnt, NumRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        ActivateControls frmBankReconMenu
        Unload FrmShowPctComp
        GoTo CancelExit
      End If

      Get FileNum, cnt, OSChek
        If ChkType = 3 Then GoTo jumpall
        If ChkType = 1 Then BankNum = 999

        If OSChek.Src = ChkType Then

jumpall:
        If OSChek.Bankcode = BankNum Or BankNum = 999 Then
          If OSChek.Cleared = WhatKind Then

            If OSChek.chkdate <= SrchDate Then

               NumChks = NumChks + 1
           
               TotalChks# = Round#(TotalChks# + OSChek.Amt)
               Select Case OSChek.Src
                 Case 0
                   T$ = "AP"
                 Case 1
                   T$ = "PR"
               End Select

               ToPrint$ = Space$(80)
               LSet ToPrint$ = Str$(OSChek.ChkNum)
               Mid$(ToPrint$, 10) = Format(DateAdd("d", OSChek.chkdate, "12-31-1979"), "mm/dd/yy")
               Mid$(ToPrint$, 20) = T$
               Mid$(ToPrint$, 25) = QPTrim(OSChek.Desc)
               Mid$(ToPrint$, 60) = Using$("##,###,###.##", Str$(OSChek.Amt))
               Mid$(ToPrint$, 78) = Str$(OSChek.Bankcode)
               Print #PRNFileNum, ToPrint$
            End If

         End If
       End If
     End If
   Next

   Print #PRNFileNum, "" 'Add a blank line after last line

   ToPrint$ = Space$(80)
   LSet ToPrint$ = Using$("#####", Str$(NumChks))
   Mid$(ToPrint$, 8) = "Checks listed totaling: "
   Mid$(ToPrint$, 60) = Using$("##,###,###.##", Str$(TotalChks#))

   Print #PRNFileNum, ToPrint$
   'UNLOCK FileNum
   Print #PRNFileNum, Chr$(12)
   Close
   ActivateControls frmBankReconMenu
   ViewPrint ReportFile$, RptTitle$
Kill ReportFile$     'Clean up after ourselves
Close
ChkListExit:
Exit Sub
DoHeading:
  ToPrint$ = Space$(80)
  Print #PRNFileNum, GLUserName$
  Print #PRNFileNum, RptTitle$
  Mid$(ToPrint$, 1) = "Check"
  Mid$(ToPrint$, 10) = "Date"
  Mid$(ToPrint$, 20) = "Type"
  Mid$(ToPrint$, 28) = "Description"
  Mid$(ToPrint$, 64) = "Amount"
  Mid$(ToPrint$, 77) = "Bank"
  Print #PRNFileNum, ToPrint$
  Print #PRNFileNum, String(80, "=")
  Return
Getout:
'UNLOCK FileNum
Close
Exit Sub
CancelExit:
Exit Sub
End Sub

Public Sub SortCheckFile()
  Dim ccnt As Long, OSChekFile As Integer, OSChkDate As String
  Dim NumOSChks As Long, lb As Long, UB As Long, ChkRecLen As Integer
  Dim ChkFile As Integer, Numchkrecs As Long
  Dim OSChek As OSChekRecType
  FrmShowPctComp.Label1 = "Sorting Outstanding Checks"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmBankReconMenu
  
  OpenOSChekFile OSChekFile, NumOSChks

  ChkRecLen = Len(OSChek)
  ChkFile = FreeFile
  Open "crchk1.dat" For Output As ChkFile
  Close ChkFile
  ChkFile = FreeFile
  Open "crchk1.dat" For Random Shared As ChkFile Len = ChkRecLen

  Numchkrecs = LOF(OSChekFile) \ ChkRecLen

  If Numchkrecs < 1 Then
    Close
    
    MsgBox "No Checks to Sort.", vbOKOnly, "NO Checks"
    Exit Sub
  End If

  ReDim Index(1 To Numchkrecs) As OSChekSrtType

  For ccnt = 1 To Numchkrecs
    Get OSChekFile, ccnt, OSChek
    Index(ccnt).ChkNum = OSChek.ChkNum
    Index(ccnt).RecNo = ccnt
  Next
  lb = LBound(Index)
  UB = UBound(Index)
  QSort Index(), lb, UB

  For ccnt = 1 To Numchkrecs
    FrmShowPctComp.ShowPctComp ccnt, Numchkrecs
    Get OSChekFile, Index(ccnt).RecNo, OSChek
    Put ChkFile, ccnt, OSChek
  Next
  Close OSChekFile
  Close ChkFile
  Call MainLog("SortChekrecFile.")
  KillFileD "crchek.dat"
  Name "crchk1.dat" As "crchek.dat"
  ActivateControls frmBankReconMenu
End Sub

Private Sub QSort(Idxbuff() As OSChekSrtType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As OSChekSrtType
  Dim Temp2 As OSChekSrtType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).ChkNum < Temp.ChkNum
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.ChkNum < Idxbuff(lngCurHigh).ChkNum
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub

