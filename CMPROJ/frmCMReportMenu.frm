VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMReportMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Management Reports"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmCMReportMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitMenu 
      Caption         =   "E&XIT Menu"
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
      Left            =   3846
      TabIndex        =   3
      Top             =   5688
      Width           =   4524
   End
   Begin VB.CommandButton cmdPaymentJournal 
      Caption         =   "Print Cash Management &Journal"
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
      Left            =   3846
      TabIndex        =   0
      Top             =   3288
      Width           =   4524
   End
   Begin VB.CommandButton cmdPaymentInq 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Payment &Inquiry"
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
      Left            =   3846
      TabIndex        =   1
      Top             =   4092
      Width           =   4524
   End
   Begin VB.CommandButton cmdMiscCodes 
      BackColor       =   &H008F8265&
      Caption         =   "Print &Miscellaneous Codes"
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
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   4884
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/14/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2388
      X2              =   3348
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CASH MANAGEMENT REPORTS"
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
      Left            =   3348
      TabIndex        =   5
      Top             =   1152
      Width           =   5292
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
End
Attribute VB_Name = "frmCMReportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
'Dim FormOver As clsFormOverRider
Private Temp_Class As Resize_Class

Private Sub cmdExitMenu_Click()
  Load frmCMMainMenu
  DoEvents
  frmCMMainMenu.Show
  Unload Me
End Sub

Private Sub cmdMiscCodes_Click()
  frmReportOpt.Show 1
  DeActivateControls Me
  If rptopt = 1 Then
    'do the graphics
   PrintMiscCodeList 1
  ElseIf rptopt = 2 Then
    'do the text
   PrintMiscCodeList 2
   ActivateControls Me
  Else
    ActivateControls Me
  End If
End Sub

Private Sub cmdPaymentInq_Click()
  Load frmRptCMInquiry
  DoEvents
  frmRptCMInquiry.Show
  Unload Me
End Sub

Private Sub cmdPaymentJournal_Click()
  Load frmRptCMJournal
  DoEvents
  frmRptCMJournal.Show
  Unload Me
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Call CMLog("Close via CM ReportMenu" + PWUser$)
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdPaymentJournal.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub PrintMiscCodeList(rptopt)
  Dim Dash As String, Report As String, MaxLine As Integer, MiscCodeRecLen As Integer
  Dim MFile As Integer, NumOfMiscRecs As Long, RptFile As Integer, cnt As Long
  Dash$ = String$(80, "-")
  FF$ = Chr$(12)
  MaxLine = 56
  PageNo = 0
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MFile) \ MiscCodeRecLen

  If NumOfMiscRecs = 0 Then
    Close MFile
    MsgBox "No Codes to Print.", vbOKOnly, "No Codes"
    GoTo ExitCodePrint
  End If

'  FOR Cnt = 1 TO NumOfMiscRecs
'    GET MFile, Cnt, MiscCodeRec(1)
'    GLAcct$ = QPTrim$(MiscCodeRec(1).GlAcctNumb)
'    PerPos = INSTR(GLAcct$, ".")
'    DO WHILE PerPos > 0
'      GLAcct$ = LEFT$(GLAcct$, PerPos - 1) + MID$(GLAcct$, PerPos + 1)
'      PerPos = INSTR(GLAcct$, ".")
'    LOOP
'    LSET MiscCodeRec(1).GlAcctNumb = GLAcct$
'    PUT MFile, Cnt, MiscCodeRec(1)
'  NEXT
  Report$ = UBPath$ + "MiscCode.rpt"
  RptFile = FreeFile
  Open Report$ For Output As RptFile
  GoSub PrintCodeHeader
  For cnt = 1 To NumOfMiscRecs
    Get MFile, cnt, MiscCodeRec(1)
    Print #RptFile, cnt; Tab(8); MiscCodeRec(1).MiscCode; Tab(18); MiscCodeRec(1).Description; Tab(50); MiscCodeRec(1).GlAcctNumb;
    If QPTrim$(MiscCodeRec(1).InactiveFlag) = "Y" Then
      Print #RptFile, Tab(68); "Yes"
    Else
      Print #RptFile,
    End If
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLine - 5 Then
      Print #RptFile, FF$
      GoSub PrintCodeHeader
    End If
  Next
  If rptopt = 2 Then
    Print #RptFile, FF$
  End If
  Close
  If rptopt = 2 Then
    ViewPrint Report$, "Miscellaneous Code Listing"
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmCMReportMenu
    ARptLineRpt.GetName Report$
    ARptLineRpt.startrpt
  End If
  'Kill Report$
  Exit Sub

PrintCodeHeader:
  PageNo = PageNo + 1
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, "Miscellaneous Payment Codes Listing"; Tab(70); "Page:"; PageNo
  Print #RptFile, "        Code         Description                  GL Account      Inactive"
  Print #RptFile, Dash$
  LineCnt = 5
Return

ExitCodePrint:
End Sub

