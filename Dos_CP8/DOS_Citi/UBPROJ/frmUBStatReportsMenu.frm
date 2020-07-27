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
   Icon            =   "frmUBStatReportsMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitUBStatReportsMenu 
      Caption         =   "E&xit Statistical Reports Menu"
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
      Left            =   3846
      TabIndex        =   2
      Top             =   5736
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
      Left            =   3846
      TabIndex        =   1
      Top             =   4176
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
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3408
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "3:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/17/2003"
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2400
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2496
      X2              =   2496
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9840
      X2              =   9840
      Y1              =   2304
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8880
      X2              =   9840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8880
      X2              =   8880
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8880
      X2              =   9840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2496
      X2              =   3216
      Y1              =   8304
      Y2              =   8304
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
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
      Left            =   3030
      TabIndex        =   3
      Top             =   1440
      Width           =   6156
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5916
      Index           =   0
      Left            =   2496
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
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

Private Sub cmdConsRateCode_Click()
'use default values of screen settings
  frmRptSCnsmpRateCode.GetRptInfo 1
  frmRptSCnsmpRateCode.Show
  Unload frmUBStatReportsMenu
End Sub

Private Sub cmdCustCountRate_Click()
  frmReportOpt.Show 1
  DeActivateControls Me
  If rptopt = 1 Then
    'do the graphics
    CustbyRate2
  ElseIf rptopt = 2 Then
    'do the text
    CustbyRate
  End If
  ActivateControls Me
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

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
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
    If cmdExitUBStatReportsMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
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
  Dim UBCust As Integer, NumOfRecs As Long
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


  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen

  UBRpt = FreeFile
  Open UBPath$ + "UBCBRATE.RPT" For Output As UBRpt


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
        If RCnt = 5 Then GoTo SkipThisRev
        ThisCode$ = QPTrim$(UBCustRec(1).Serv(RCnt).RATECODE)
        If Len(ThisCode$) > 0 Then
          For ZCnt = 1 To NumOfRateRecs
            If QPTrim$(UBRateTblRec(ZCnt).RATECODE) = ThisCode$ Then
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
    Print #UBRpt, UBRateTblRec(cnt).RATECODE; Tab(12); UBRateTblRec(cnt).RATEDESC; Tab(51); Using("####", CustCnt(cnt, 1)); Tab(68) '; Using("####", CustCnt(cnt, 2))
  Next

  GoSub DOCustRateFooter

  Close

  Erase UBCustRec
'  If Not AbortFlag Then
'    PrintRptFile , , 1, RetCode,
'  End If
  ViewPrint "UBCBRATE.RPT", "Customer By Rate Code Report."
  'KillFile "UBCOLIST.RPT"

ExitCustRateReport:

  Exit Sub

DOCustRateHeader:
  PageNo = PageNo + 1
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(26); "Customers By Rate Code Report"; Tab(70); "Page #"; PageNo
  Print #UBRpt, "Report Date: "; Date$;
  Print #UBRpt, ""
  Print #UBRpt, "Rate        Description                    Customer Count " '      Irrigation"
  Print #UBRpt, Dash80$
  Linecnt = 6
  Return

DOCustRateFooter:
  Print #UBRpt, ""
  Print #UBRpt, "Total Customers Count: "; Using("##,###", CustomerCnt)
  Return


End Sub
Private Sub CustbyRate2()
  Dim NumOfRateRecs As Long, UBRateTblRecLen As Integer
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

  UBRpt = FreeFile
  Open UBPath$ + "UBCBRATE.RPT" For Output As UBRpt


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
        If RCnt = 5 Then GoTo SkipThisRev
        ThisCode$ = QPTrim$(UBCustRec(1).Serv(RCnt).RATECODE)
        If Len(ThisCode$) > 0 Then
          For ZCnt = 1 To NumOfRateRecs
            If QPTrim$(UBRateTblRec(ZCnt).RATECODE) = ThisCode$ Then
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
    Print #UBRpt, UBRateTblRec(cnt).RATECODE + "~" + UBRateTblRec(cnt).RATEDESC + "~" + Using("####", CustCnt(cnt, 1))
  Next

  'GoSub DOCustRateFooter

  Close

  Erase UBCustRec
'  ViewPrint "UBCBRATE.RPT", "Customer By Rate Code Report."
  'KillFile "UBCOLIST.RPT"
    Load frmLoadingRpt
    ARptSCustCountRate.txtDate = Now
    ARptSCustCountRate.txtTown = TownName$
    ARptSCustCountRate.Title = "Customer By Rate Code Report."
    ARptSCustCountRate.totCust = Using("##,###", CustomerCnt)
    ARptSCustCountRate.GetName "UBCBRATE.RPT"
    ARptSCustCountRate.startrpt

ExitCustRateReport:

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

