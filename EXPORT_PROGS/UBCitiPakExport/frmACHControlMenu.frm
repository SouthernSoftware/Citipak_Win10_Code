VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmACHControlMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACH Bank Draft Control Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmACHControlMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdSetup 
      Height          =   491
      Left            =   4005
      TabIndex        =   0
      ToolTipText     =   "Press to bring up the screen where payroll bank draft information is entered."
      Top             =   3343
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   866
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
      ButtonDesigner  =   "frmACHControlMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrenote 
      Height          =   491
      Left            =   4005
      TabIndex        =   1
      ToolTipText     =   "Press to begin processing the payroll bank draft prenote file."
      Top             =   4147
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   866
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
      ButtonDesigner  =   "frmACHControlMenu.frx":0AF0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   491
      Left            =   4005
      TabIndex        =   2
      ToolTipText     =   "Press to print a report of all employees earmarked for payroll drafting."
      Top             =   4938
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   866
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
      ButtonDesigner  =   "frmACHControlMenu.frx":0D16
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   491
      Left            =   4005
      TabIndex        =   3
      ToolTipText     =   "Press to escape to the 'Control Maintenance' menu."
      Top             =   5752
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   866
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
      ButtonDesigner  =   "frmACHControlMenu.frx":0F3E
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.243
      Y2              =   7892.757
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7884
      Y2              =   7884
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2151.243
      Y2              =   7880.108
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7884
      Y2              =   7884
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACH BANK DRAFT MAINTENANCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   4
      Top             =   1248
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1097
      Index           =   1
      Left            =   1500
      Top             =   896
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   2101
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
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
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
End
Attribute VB_Name = "frmACHControlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdSetup_Click()
  frmACHDraftInfo.Show
  DoEvents
  Unload frmACHControlMenu
End Sub

Private Sub cmdExit_Click()
  KillFile "achcont.dat"
  frmControlFileMaint.Show
  DoEvents
  Unload frmACHControlMenu
End Sub

Private Sub cmdList_Click()
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

Private Sub cmdPrenote_Click()
  InFileNames(1) = "PRDATA\PRDRAFTI.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  
  frmCreatePreNoteFiles.Show
  DoEvents
  Unload frmACHControlMenu

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%M"
      Call cmdExit_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim One As Integer
  Dim DHandle As Integer
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  
  One = 1
  DHandle = FreeFile
  Open "achcont.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  Me.HelpContextID = hlpACHBankDraftM

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
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
      MainLog ("Payroll.exe terminated via menu bar on frmControlFileMaint.")
      KillFile "achcont.dat"
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  cmdSetup.Enabled = False
  cmdPrenote.Enabled = False
  cmdList.Enabled = False
  cmdExit.Enabled = False
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = False
      End If
  Next cnt
    EnableCloseButton Me.hwnd, False
     
End Sub

Private Sub ActivateControls()
  Dim cmdButton As CommandButton
  Dim x As Control
  Dim cnt As Integer
  
  cmdSetup.Enabled = True
  cmdPrenote.Enabled = True
  cmdList.Enabled = True
  cmdExit.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
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
  
'  OpenPPDefaultFile PHandle
'  Get PHandle, 1, PPDRec
'  Close PHandle
'  If PPDRec.PACTIVE = 0 Then
'     frmWarnNOPPD.Show
'     DoEvents
'     Exit Sub
'  End If

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
  
  If EmpCnt = 0 Then
    MsgBox "There are no employees currently participating in electronic funds transfer."
    Close
    Exit Sub
  End If
  
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
   
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PRUNIT.DAT"
  InFileNames(2) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  
'  OpenPPDefaultFile PHandle
'  Get PHandle, 1, PPDRec
'  Close PHandle
'  If PPDRec.PACTIVE = 0 Then
'     frmWarnNOPPD.Show
'     DoEvents
'     Exit Sub
'  End If

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
  If EmpCnt = 0 Then
    MsgBox "There are no employees currently participating in electronic funds transfer."
    Close
    Exit Sub
  End If
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

