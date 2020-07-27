VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBStatReportsMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility Billing Statistical Reports"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmUBStatReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdConsumpRange 
      Caption         =   "Consumption by Range Analysis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3828
      TabIndex        =   7
      Top             =   5572
      Width           =   4524
   End
   Begin VB.CommandButton cmdTopTen 
      Caption         =   "Top Ten"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3828
      TabIndex        =   3
      Top             =   4857
      Width           =   4524
   End
   Begin VB.CommandButton cmdConsmpRpt2 
      Caption         =   "Consumption by Rate,Book,Cycle,Group"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3828
      TabIndex        =   2
      Top             =   4142
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitUBStatReportsMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3828
      TabIndex        =   4
      Top             =   6288
      Width           =   4524
   End
   Begin VB.CommandButton cmdCustCountRate 
      Caption         =   "Customer Count by &Rate Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3828
      TabIndex        =   1
      Top             =   3427
      Width           =   4524
   End
   Begin VB.CommandButton cmdConsRateCode 
      BackColor       =   &H008F8265&
      Caption         =   "&Consumption by Rate Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   3828
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2712
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "10:00 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "3/16/2006"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      Index           =   1
      X1              =   8976
      X2              =   8976
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8976
      X2              =   9696
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8856
      X2              =   9816
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8856
      X2              =   9816
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8856
      X2              =   8856
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9816
      X2              =   9816
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1776
      Top             =   744
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UTILITY BILLING STATISTICAL REPORTS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3012
      TabIndex        =   5
      Top             =   1104
      Width           =   6156
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8976
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8856
      Top             =   1824
      Width           =   972
   End
End
Attribute VB_Name = "frmUBStatReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdConsmpRpt2_Click()
  frmRptSCnsmpRateCodeN.Show
  Unload Me
End Sub

Private Sub cmdConsRateCode_Click()
'use default values of screen settings
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBRATE.DAT") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO RateCode FILE"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO RATE CODE FILE!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
  If Not Exist(UBPath$ + "UBCust.dat") Or Not Exist(UBPath$ + "UBTrans.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO Cust Info"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO CUSTOMER INFORMATION!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If
    frmRptSCnsmpRateCode.GetRptInfo 1
    frmRptSCnsmpRateCode.Show
    Unload frmUBStatReportsMenu
  
End Sub

Private Sub CmdConsumpRange_Click()
  frmRptSCnsmpRange.Show
  Unload Me
End Sub

Private Sub cmdCustCountRate_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist(UBPath$ + "UBCust.dat") Or Not Exist(UBPath$ + "UBRATE.dat") Then
    frmMsgDialog.RetLabel = "-2"
    UBLog "ERROR: NO Cust Info"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "NO CUSTOMER INFORMATION!"
    MsgText(3) = ""
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
    Exit Sub
  End If

  frmReportOpt.Show 1
  DeActivateControls Me
  If rptopt = 1 Then
    'do the graphics
    CustbyRate2
  ElseIf rptopt = 2 Then
    'do the text
    CustbyRate
    ActivateControls Me
  Else
    ActivateControls Me
  End If
End Sub

Private Sub cmdExitUBStatReportsMenu_Click()
frmUBMainMenu.Show
Unload frmUBStatReportsMenu
End Sub
'Removed this option - only used by MOWASA
'Private Sub cmdIrrigation_Click()
'  frmRptSCnsmpRateCode.GetRptInfo 2
'  frmRptSCnsmpRateCode.Caption = "Irrigation Consumption By Rate Code"
'  frmRptSCnsmpRateCode.Label1 = "Irrigation Consumption By Rate Code"
'  frmRptSCnsmpRateCode.Show
'  Unload frmUBStatReportsMenu
'End Sub

Private Sub cmdTopTen_Click()
  frmRptSCnsmpTopTen.Show
  Unload Me
End Sub


Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  Me.HelpContextID = hlpStatistical
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitUBStatReportsMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via StatReportsMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitUBStatRep_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub cmdExitUBStatRep_Click()
  frmUBMainMenu.Show
  Unload frmUBStatReportsMenu
End Sub
Private Sub CustbyRate()
  Dim NumOfRateRecs As Long, UBRateTblRecLen As Integer
  Dim UBFile As Integer, cnt As Long, Dash80 As String
  Dim UBCustRecLen As Integer, RCnt As Integer, ZCnt As Integer
  Dim UBCust As Integer, NumOfRecs As Long, ReportFile As String
  Dim UBRpt As Integer, ThisCode As String, CustomerCnt As Long
  NumOfRateRecs = GetNumRateRecs
  ReDim UBRateTblRec(1 To NumOfRateRecs) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))

  UBFile = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As UBFile Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get UBFile, cnt, UBRateTblRec(cnt)
  Next
  Close

  ReDim CustCnt(1 To NumOfRateRecs, 1 To 2) As Integer

  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = String$(80, "-")

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReportFile$ = UBPath$ + "UBCBRATE.RPT"
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen

  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt


  FrmShowPctComp.Label1 = "Creating Customer Count By Rate Report"
  FrmShowPctComp.Show


  GoSub DOCustRateHeader

  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitCustRateReport
    End If

    Get UBCust, cnt, UBCustRec(1)
    If UBCustRec(1).Status = "A" Then
      '*************************************
      '   Main body of Printing goes here
      For RCnt = 1 To 15
        'If RCnt = 5 Then GoTo SkipThisRev
        ThisCode$ = QPTrim$(UBCustRec(1).serv(RCnt).Ratecode)
        If Len(ThisCode$) > 0 Then
          For ZCnt = 1 To NumOfRateRecs
            If QPTrim$(UBRateTblRec(ZCnt).Ratecode) = ThisCode$ Then
              CustCnt(ZCnt, 1) = CustCnt(ZCnt, 1) + 1
              Exit For
            End If
          Next
        End If
SkipThisRev:
      Next
'Irrigation column not used
'      ThisCode$ = QPTrim$(UBCustRec(1).Serv(5).RATECODE)
'      If Len(ThisCode$) > 0 Then
'        For ZCnt = 1 To NumOfRateRecs
'          If QPTrim$(UBRateTblRec(ZCnt).RATECODE) = ThisCode$ Then
'            CustCnt(ZCnt, 2) = CustCnt(ZCnt, 2) + 1
'            Exit For
'          End If
'        Next
'      End If
      'LineCnt = LineCnt + 1
      CustomerCnt = CustomerCnt + 1
      '*************************************
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
    End If
    'ShowPctComp cnt, NumOfRecs
  Next

  For cnt = 1 To NumOfRateRecs
    Print #UBRpt, UBRateTblRec(cnt).Ratecode; Tab(12); UBRateTblRec(cnt).RATEDESC; Tab(51); Using("####", CustCnt(cnt, 1)); Tab(68) '; Using("####", CustCnt(cnt, 2))
  Next

  GoSub DOCustRateFooter

  Close

  Erase UBCustRec
'  If Not AbortFlag Then
'    PrintRptFile , , 1, RetCode,
'  End If
  ViewPrint ReportFile$, "Customer By Rate Code Report."
  'KillFile "UBCOLIST.RPT"

ExitCustRateReport:

  Exit Sub

DOCustRateHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Tab(26); "Customers By Rate Code Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$;
  Print #UBRpt, ""
  Print #UBRpt, "Rate        Description                    Customer Count " '      Irrigation"
  Print #UBRpt, Dash80$
  LineCnt = 6
  Return

DOCustRateFooter:
  Print #UBRpt, ""
  Print #UBRpt, "Total Customers Count: "; Using("##,###", CustomerCnt)
  Return


End Sub
Private Sub CustbyRate2()
  Dim NumOfRateRecs As Long, UBRateTblRecLen As Integer, ReportFile As String
  Dim UBFile As Integer, cnt As Long, ZCnt As Integer, UBRpt As Integer
  Dim UBCustRecLen As Integer, RCnt As Integer, ThisCode As String
  Dim UBCust As Integer, NumOfRecs As Long, CustomerCnt As Long
  NumOfRateRecs = GetNumRateRecs
  ReDim UBRateTblRec(1 To NumOfRateRecs) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))

  UBFile = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As UBFile Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get UBFile, cnt, UBRateTblRec(cnt)
  Next
  Close

  ReDim CustCnt(1 To NumOfRateRecs, 1 To 2) As Integer

  '***************

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))


  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen
  ReportFile$ = UBPath$ + "UBCBRATE.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt


  FrmShowPctComp.Label1 = "Creating Customer Count By Rate Report"
  FrmShowPctComp.Show


  For cnt = 1 To NumOfRecs
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitCustRateReport
    End If

    Get UBCust, cnt, UBCustRec(1)
    
    If UBCustRec(1).Status = "A" Then
      '*************************************
      '   Main body of Printing goes here
      For RCnt = 1 To 15
       ' If RCnt = 5 Then GoTo SkipThisRev
        ThisCode$ = QPTrim$(UBCustRec(1).serv(RCnt).Ratecode)
        If Len(ThisCode$) > 0 Then
          For ZCnt = 1 To NumOfRateRecs
            If QPTrim$(UBRateTblRec(ZCnt).Ratecode) = ThisCode$ Then
''''            If cnt = 129 Then Stop
''''              If QPTrim$(UBRateTblRec(ZCnt).Ratecode) = "50" Then
''''                Print #UBRpt, ThisCode$ + "~" + Str(cnt) + "~" + " "
''''              End If
              CustCnt(ZCnt, 1) = CustCnt(ZCnt, 1) + 1
              Exit For
            End If
          Next
        End If
SkipThisRev:
      Next
'Irrigation column not used
'      ThisCode$ = QPTrim$(UBCustRec(1).Serv(5).RATECODE)
'      If Len(ThisCode$) > 0 Then
'        For ZCnt = 1 To NumOfRateRecs
'          If QPTrim$(UBRateTblRec(ZCnt).RATECODE) = ThisCode$ Then
'            CustCnt(ZCnt, 2) = CustCnt(ZCnt, 2) + 1
'            Exit For
'          End If
'        Next
'      End If
      'LineCnt = LineCnt + 1
      CustomerCnt = CustomerCnt + 1
      '*************************************
'      If AskAbandonPrint% Then
'        AbortFlag = True
'        Exit For
'      End If
    End If
    'ShowPctComp cnt, NumOfRecs
  Next

  For cnt = 1 To NumOfRateRecs
    Print #UBRpt, UBRateTblRec(cnt).Ratecode + "~" + UBRateTblRec(cnt).RATEDESC + "~" + Using("####", CustCnt(cnt, 1))
  Next

  'GoSub DOCustRateFooter

  Close

  Erase UBCustRec
'  ViewPrint "UBCBRATE.RPT", "Customer By Rate Code Report."
  'KillFile "UBCOLIST.RPT"
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBStatReportsMenu
    ARptSCustCountRate.txtDate = Now
    ARptSCustCountRate.txtTown = TOWNNAME$
    ARptSCustCountRate.Title = "Customer By Rate Code Report."
    ARptSCustCountRate.totCust = Using("##,###", CustomerCnt)
    ARptSCustCountRate.GetName ReportFile$
    ARptSCustCountRate.startrpt

ExitCustRateReport:
  ActivateControls Me
  Exit Sub

'DOCustRateHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(26); "Customers By Rate Code Report"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, "Report Date: "; Date$;
'  Print #UBRpt, ""
'  Print #UBRpt, "Rate        Description                    Customer Count" '      Irrigation"
'  Print #UBRpt,
'  Linecnt = 6
'  Return
'
'DOCustRateFooter:
'  Print #UBRpt, ""
'  Print #UBRpt, "Total Customers Count: "; Using("##,###", CustomerCnt)
'  Return


End Sub

