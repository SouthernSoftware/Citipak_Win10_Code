VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmACHBankDraftMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACH Bank Draft Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmACHBankDraftMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdEmp2Draft 
      Height          =   495
      Left            =   3975
      TabIndex        =   0
      Top             =   2925
      Width           =   3600
      _Version        =   131072
      _ExtentX        =   6350
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmACHBankDraftMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDraftTransFile 
      Height          =   495
      Left            =   3975
      TabIndex        =   1
      Top             =   3735
      Width           =   3600
      _Version        =   131072
      _ExtentX        =   6350
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmACHBankDraftMenu.frx":0AB7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrenoteFile 
      Height          =   495
      Left            =   3975
      TabIndex        =   2
      Top             =   4530
      Width           =   3600
      _Version        =   131072
      _ExtentX        =   6350
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmACHBankDraftMenu.frx":0CAA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintDraftEmpList 
      Height          =   495
      Left            =   3975
      TabIndex        =   3
      Top             =   5340
      Width           =   3600
      _Version        =   131072
      _ExtentX        =   6350
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmACHBankDraftMenu.frx":0E98
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   3975
      TabIndex        =   4
      Top             =   6150
      Width           =   3600
      _Version        =   131072
      _ExtentX        =   6350
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmACHBankDraftMenu.frx":1088
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Index           =   2
      Left            =   2101
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Index           =   1
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   1095
      Index           =   0
      Left            =   1500
      Top             =   895
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACH Bank Draft Menu"
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
      Left            =   2820
      TabIndex        =   5
      Top             =   1248
      Width           =   6012
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7884
      Y2              =   7884
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.243
      Y2              =   7892.757
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2147.351
      Y2              =   7866.486
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7884
      Y2              =   7884
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2220
      Top             =   2211
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8712
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   768
      Width           =   8652
   End
End
Attribute VB_Name = "frmACHBankDraftMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdDraftTransFile_Click()
  InFileNames(1) = "PRDATA\PRDRAFTI.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT" 'these files
  'are needed here to keep the program from crashing
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  frmCreateDraftFile.Show
  DoEvents
  Unload frmACHBankDraftMenu
End Sub

Private Sub cmdEmp2Draft_Click()
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call PrintText
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call PrintGraphics
  Else
    Exit Sub
  End If

End Sub

Private Sub PrintGraphics()
  Dim Image1$
  Dim EmpRecSize As Integer
  Dim PPDFLen As Integer, RptName$
  Dim RptTitle$, RecNo As Long
  Dim RptHandle As Integer
  Dim EHandle As Integer
  Dim NumOfRecs As Long
  Dim PPDFFile As Integer, PPDate$
  Dim TNet#, EmpCnt As Integer
  Dim UHandle As Integer
  Dim Page As Integer
  Dim dlm$
  
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  Image1$ = "#,##0.00"

  ReDim Unit(1) As UnitFileRecType
  ReDim PPDFInfo(1) As PRPPDraftInfoType

  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 30
  ReDim ETitle(1) As String * 12

  EmpRecSize = Len(Emp2Rec(1))
  PPDFLen = Len(PPDFInfo(1))
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle

  RptTitle$ = "Pay Period Draft Report"

  RptName$ = "PRRPTS\PPDFG.RPT"

  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  OpenEmpData2File EHandle
  OpenPPDraftInfo PPDFFile
  NumOfRecs = LOF(PPDFFile) \ Len(PPDFInfo(1))
  
  FrmShowPctComp.Label1 = RptTitle$
  FrmShowPctComp.Show
  If NumOfRecs = 0 Then Unload FrmShowPctComp
  Get #PPDFFile, 1, PPDFInfo(1)
  PPDate$ = MakeRegDate(PPDFInfo(1).DraftDate)

  For RecNo = 1 To NumOfRecs
    Get #PPDFFile, RecNo, PPDFInfo(1)
    If CLng(PPDFInfo(1).EmpRec) = 0 Then GoTo SkipIt
    Get EHandle, CLng(PPDFInfo(1).EmpRec), Emp2Rec(1)
    EmpCnt = EmpCnt + 1
    GoSub PrintEmp2DraftData
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
SkipIt:
  Next RecNo
  
  Unload FrmShowPctComp

  Close EHandle
  Close RptHandle
  Close PPDFFile
  arEmp2Draft.Show
  frmLoadingRpt.Show
  Exit Sub

PrintEmp2DraftData:
  TNet# = OldRound#(TNet# + PPDFInfo(1).NetPay)
  EName(1) = QPTrim$(Emp2Rec(1).EmpLName) + ", " + QPTrim$(Emp2Rec(1).EmpFName)
  '                       0                        1
  Print #RptHandle, Unit(1).UFEMPR; dlm; MakeRegDate(PPDFInfo(1).DraftDate); dlm;
  '                        2                 3                    4
  Print #RptHandle, Emp2Rec(1).EmpNo; dlm; EName(1); dlm; Emp2Rec(1).DRAFTCOD; dlm;
  '                          5                                 6
  Print #RptHandle, Emp2Rec(1).EMPDDACC; dlm; Using$("###,##0.00", PPDFInfo(1).NetPay) '; dlm;
Return

End Sub

Private Sub cmdExit_Click()
   frmPayrollProcessingMenu.Show
   DoEvents
   Unload frmACHBankDraftMenu
   MainLog ("ACH Bank Draft Menu exited.")
End Sub

Private Sub cmdPrenoteFile_Click()
  InFileNames(1) = "PRDATA\PRDRAFTI.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  
  frmCreatePreNoteFiles.Show
  DoEvents
  Unload frmACHBankDraftMenu

End Sub

Private Sub cmdPrintDraftEmpList_Click()
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call PrintDraftEmpListT
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call PrintDraftEmpListG
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintDraftEmpListG()
  Dim Image1$
  Dim EmpRecSize As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Long
  Dim RptTitle$, RptName$
  Dim RptHandle As Integer
  Dim RecNo As Long
  Dim UHandle As Integer
  Dim EHandle As Integer
  Dim EmpCnt As Integer
  Dim IdxNHandle As Integer
  Dim x As Integer
  Dim Page As Integer
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim BankName As String * 26
  Dim dlm$
  
'  Me.HelpContextID =
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  If PPDRec.PACTIVE = 0 Then
     frmWarnNOPPD.Show
     DoEvents
     Exit Sub
  End If

  Image1$ = "#,##0.00"

  ReDim Unit(1) As UnitFileRecType

  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 23
  ReDim ETitle(1) As String * 12

  EmpRecSize = Len(Emp2Rec(1))
  
  IdxRecLen = 2

  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) \ IdxRecLen
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  RptTitle$ = "Employee Draft Listing"
  FrmShowPctComp.Label1 = RptTitle$
  FrmShowPctComp.Show ' , Me

  RptName$ = "PRRPTS\EMPDFLSTG.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle

  OpenEmpData2File EHandle

  For RecNo = 1 To NumOfRecs
    Get EHandle, CLng(IdxBuff(RecNo)), Emp2Rec(1)
    If Not Emp2Rec(1).Deleted Then
      If Emp2Rec(1).DRAFTCOD = "C" Or Emp2Rec(1).DRAFTCOD = "S" Then
        EmpCnt = EmpCnt + 1
        GoSub PrintEmpListData
      End If
    End If

SkipThisOne:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close EHandle
  Close RptHandle
  arEmpDraftList.Show
  frmLoadingRpt.Show
  Exit Sub


PrintEmpListData:
  EName(1) = QPTrim$(Emp2Rec(1).EmpLName) + ", " + QPTrim$(Emp2Rec(1).EmpFName)
  '                        0               1                  2
  Print #RptHandle, Unit(1).UFEMPR; dlm; Date$; dlm; Emp2Rec(1).EmpNo; dlm;
  '                    3                        4                                  5
  Print #RptHandle, EName(1); dlm; QPTrim$(Emp2Rec(1).DRAFTCOD); dlm; QPTrim$(Emp2Rec(1).EMPDDACC); dlm;
  '                     6                             7                         8                            9
  Print #RptHandle, Emp2Rec(1).TRANSIT; dlm; Emp2Rec(1).PRENOTED; dlm; Emp2Rec(1).BankName; dlm; QPTrim$(Emp2Rec(1).BANKLOC)
  Return


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%B"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim FileHandle As Integer
    
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpACHBankDraftM
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
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmACHBankDraftMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim Image1$
  Dim Dash As String * 78
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim EmpRecSize As Integer
  Dim PPDFLen As Integer, RptName$
  Dim RptTitle$, RecNo As Long
  Dim RptHandle As Integer
  Dim EHandle As Integer
  Dim NumOfRecs As Long
  Dim PPDFFile As Integer, PPDate$, FF$
  Dim TNet#, EmpCnt As Integer
  Dim UHandle As Integer
  Dim Page As Integer
  
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  Image1$ = "#,##0.00"

  ReDim Unit(1) As UnitFileRecType
  ReDim PPDFInfo(1) As PRPPDraftInfoType

  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 30
  ReDim ETitle(1) As String * 12

  MaxLines = 55
  
  FF$ = Chr$(12)
  LineCnt = 0
  Dash = String$(78, "-")
  EmpRecSize = Len(Emp2Rec(1))
  PPDFLen = Len(PPDFInfo(1))
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle

  RptTitle$ = "Pay Period Draft Report"

  RptName$ = "PRRPTS\PPDF.RPT"

  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  OpenEmpData2File EHandle
  OpenPPDraftInfo PPDFFile
  NumOfRecs = LOF(PPDFFile) \ Len(PPDFInfo(1))
  
  FrmShowPctComp.Label1 = RptTitle$
  FrmShowPctComp.Show
  If NumOfRecs = 0 Then Unload FrmShowPctComp
  Get #PPDFFile, 1, PPDFInfo(1)
  PPDate$ = MakeRegDate(PPDFInfo(1).DraftDate)

  GoSub PrintEmp2DraftHeader
  
  For RecNo = 1 To NumOfRecs
    Get #PPDFFile, RecNo, PPDFInfo(1)
    If CLng(PPDFInfo(1).EmpRec) = 0 Then GoTo SkipIt
    Get EHandle, CLng(PPDFInfo(1).EmpRec), Emp2Rec(1)
    EmpCnt = EmpCnt + 1
    GoSub PrintEmp2DraftData
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEmp2DraftHeader
    End If
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Exit Sub
    End If
SkipIt:
  Next RecNo
  
  Unload FrmShowPctComp

  GoSub PrintEmp2DraftTotals

  Close EHandle
  Close RptHandle
  Close PPDFFile

  ViewPrint RptName$, RptTitle$, True
  Exit Sub

PrintEmp2DraftHeader:
  Page = Page + 1
  Print #RptHandle, Unit(1).UFEMPR
  Print #RptHandle, "Pay Period Draft Report"
  Print #RptHandle, "Draft Date: " + PPDate$ + "                                            Page: " + CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "    Number     Employee Name                  Bank Acct No.       Draft Amt."
          
  Print #RptHandle, Dash$
  LineCnt = 6
Return
PrintEmp2DraftData:
  TNet# = OldRound#(TNet# + PPDFInfo(1).NetPay)
  ENumb(1) = Emp2Rec(1).EmpNo
  EName(1) = QPTrim$(Emp2Rec(1).EmpLName) + ", " + QPTrim$(Emp2Rec(1).EmpFName)
  ETitle(1) = Using$("###,##0.00", PPDFInfo(1).NetPay)
  Print #RptHandle, ENumb(1) + "    " + EName(1) + " " + Emp2Rec(1).DRAFTCOD + " " + Emp2Rec(1).EMPDDACC + " " + ETitle(1)
  LineCnt = LineCnt + 1
Return

PrintEmp2DraftTotals:
  Print #RptHandle, Dash$
  Print #RptHandle, "Total Employees: " + CStr(EmpCnt)
  Print #RptHandle, "    Draft Total: " + Using$("###,##0.00", TNet#)
  Print #RptHandle, FF$
  Return

End Sub

Private Sub PrintDraftEmpListT()
  
  Dim Image1$
  Dim Dash As String * 78
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim EmpRecSize As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Long
  Dim RptTitle$, RptName$
  Dim RptHandle As Integer
  Dim RecNo As Long
  Dim UHandle As Integer
  Dim EHandle As Integer
  Dim EmpCnt As Integer
  Dim FF$
  Dim IdxNHandle As Integer
  Dim x As Integer
  Dim Page As Integer
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim BankName As String * 26
  
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  If PPDRec.PACTIVE = 0 Then
     frmWarnNOPPD.Show
     DoEvents
     Exit Sub
  End If

  Image1$ = "#,##0.00"

  ReDim Unit(1) As UnitFileRecType

  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 23
  ReDim ETitle(1) As String * 12

  MaxLines = 55

  LineCnt = 0
  Dash = String$(78, "-")
  EmpRecSize = Len(Emp2Rec(1))
  
  FF$ = Chr$(12)
  IdxRecLen = 2

  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) \ IdxRecLen
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  RptTitle$ = "Employee Draft Listing"
  FrmShowPctComp.Label1 = RptTitle$
  FrmShowPctComp.Show ' , Me

  RptName$ = "PRRPTS\EMPDFLST.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle

  OpenEmpData2File EHandle
  GoSub PrintEmpListHeader

  For RecNo = 1 To NumOfRecs
    Get EHandle, CLng(IdxBuff(RecNo)), Emp2Rec(1)
    If Not Emp2Rec(1).Deleted Then
      If Emp2Rec(1).DRAFTCOD = "C" Or Emp2Rec(1).DRAFTCOD = "S" Then
        EmpCnt = EmpCnt + 1
        GoSub PrintEmpListData
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintEmpListHeader
        End If
      End If
    End If

SkipThisOne:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  GoSub PrintEmpListTotals
  Close EHandle
  Close RptHandle

  ViewPrint RptName$, RptTitle$, True
  Exit Sub

PrintEmpListHeader:
  Page = Page + 1
  Print #RptHandle, Unit(1).UFEMPR
  Print #RptHandle, "Employee Draft Listing"
  Print #RptHandle, "Report Date: " + Date$ + "                                            Page: " + CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "    Number     Employee Name           Bank Acct No.         Transit No."
  Print #RptHandle, " Prenoted?     Bank Name                                Bank Location"
  Print #RptHandle, Dash$
  LineCnt = 7
  Return

PrintEmpListData:
  ENumb(1) = Emp2Rec(1).EmpNo
  EName(1) = QPTrim$(Emp2Rec(1).EmpLName) + ", " + QPTrim$(Emp2Rec(1).EmpFName)
  BankName = Emp2Rec(1).BankName
  ETitle(1) = Emp2Rec(1).TRANSIT
  Print #RptHandle, ENumb(1) + "    " + EName(1) + " " + QPTrim$(Emp2Rec(1).DRAFTCOD) + " " + QPTrim$(Emp2Rec(1).EMPDDACC) + "           " + ETitle(1)
  If Emp2Rec(1).PRENOTED = "Y" Then
    ENumb(1) = "Y"
  Else
    ENumb(1) = "N"
  End If '
  Print #RptHandle, "    " + ENumb(1) + BankName, Tab(55), QPTrim$(Emp2Rec(1).BANKLOC) '8/28...added BankName   QPTrim$(Emp2Rec(1).BankName)
  Print #RptHandle,
  LineCnt = LineCnt + 3
  Return

PrintEmpListTotals:
  Print #RptHandle, Dash$
  Print #RptHandle, "Total Employees: " + CStr(EmpCnt)
  Print #RptHandle, FF$
  Return

End Sub
