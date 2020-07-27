VERSION 5.00
Begin VB.Form frmLvTblWrtsvlBchConvert 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Wrightsville Beach Leave Table"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmLvTblWrtsvlBchConvert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   2976
      TabIndex        =   2
      Top             =   5316
      Width           =   1932
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "F10  &Proceed With Conversion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   5280
      TabIndex        =   1
      Top             =   5316
      Width           =   3756
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "AN ATTEMPT IS BEING MADE TO CONVERT WINDOWS DATA THAT HAS ALREADY BEEN CONVERTED. CONVERSION ABORTED."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1356
      Left            =   1500
      TabIndex        =   4
      Top             =   3744
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "CONVERSION COMPLETED SUCCESSFULLY!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   828
      Left            =   3384
      TabIndex        =   3
      Top             =   4020
      Width           =   5244
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "THIS PROGRAM IS USED SOLELY TO CONVERT THE LEAVE TABLES FOR WRIGHTSVILLE BEACH. ANY OTHER USE WILL CAUSE DATA DAMAGE."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   972
      Left            =   1008
      TabIndex        =   0
      Top             =   2592
      Width           =   9996
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   4572
      Left            =   180
      Top             =   2148
      Width           =   11292
   End
End
Attribute VB_Name = "frmLvTblWrtsvlBchConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdConvert_Click()
  Dim OldLeaveRec As OldLeaveRecType
  Dim OldHandle As Integer
  Dim LeaveRec As LeaveRecType
  Dim NewHandle As Integer
  Dim NumOfLvRecs As Integer
  Dim x As Integer
  Dim y As Integer
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim DosErnHandle As Integer
  Dim DosErnRec As DosErnCodeRecType
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim OldTransHistRec As OldTransRecType
  Dim OldTransHistHandle As Integer
  Dim OldTransWorkRec As OldTransRecType
  Dim OldTransWorkHandle As Integer
  Dim TransHistRec As TransRecType
  Dim TransHistHandle As Integer
  Dim TransWorkRec As TransRecType
  Dim TransWorkHandle As Integer
  Dim NumOfTransHistRecs As Double
  Dim NumOfTransWorkRecs As Double
  Dim Emp2Rec As EmpData2Type
  Dim EmpHandle As Integer
  Dim NumOfEmpRecs As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  If QPTrim$(UnitRec.FileVer) = "Done" Then
    Label7.Visible = True
    Close UHandle
    Exit Sub
  Else
    UnitRec.FileVer = "Done"
    Put UHandle, 1, UnitRec
    Close UHandle
  End If
  'no need for UnitRec.LMT401YN addition to UnitFileRecType...
  'already exists at Wrightsville Beach
  'no need to convert the TransRec type field
  '.Less401k because Wrightsville Beach already has it
  OpenOldLeaveFileName OldHandle
  NumOfLvRecs = LOF(OldHandle) / Len(OldLeaveRec)
  
  ReDim TempVacMax(1 To NumOfLvRecs) As Double
  ReDim TempVEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  ReDim TempSICKMAX(1 To NumOfLvRecs) As Double
  ReDim TempSEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
'  ReDim TempHolMax(1 To NumOfLvRecs) As Double
'  ReDim TempHEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
'  ReDim TempPerMax(1 To NumOfLvRecs) As Double
'  ReDim TempPEntry(1 To NumOfLvRecs, 1 To 20) As LeaveEntryType
  
  For x = 1 To NumOfLvRecs
    Get OldHandle, x, OldLeaveRec
    TempVacMax(x) = OldLeaveRec.VACMAX
    TempSICKMAX(x) = OldLeaveRec.SICKMAX
'    TempHolMax(x) = OldLeaveRec.HolMAX
'    TempPerMax(x) = OldLeaveRec.PerMAX
    For y = 1 To 17
      TempVEntry(x, y).EARN = OldLeaveRec.VEntry(y).EARN
      TempSEntry(x, y).EARN = OldLeaveRec.SEntry(y).EARN
'      TempHEntry(x, y).EARN = OldLeaveRec.HEntry(y).EARN
'      TempPEntry(x, y).EARN = OldLeaveRec.PEntry(y).EARN
      TempVEntry(x, y).YEARS = OldLeaveRec.VEntry(y).YEARS
      TempSEntry(x, y).YEARS = OldLeaveRec.SEntry(y).YEARS
'      TempHEntry(x, y).YEARS = OldLeaveRec.HEntry(y).YEARS
'      TempPEntry(x, y).YEARS = OldLeaveRec.PEntry(y).YEARS
    Next y
  Next x
  Close OldHandle
  
  OpenLeaveFileName NewHandle
  For x = 1 To NumOfLvRecs
    LeaveRec.VACMAX = TempVacMax(x)
    LeaveRec.SICKMAX = TempSICKMAX(x)
'    LeaveRec.HolMAX = TempHolMax(x)
'    LeaveRec.PerMAX = TempPerMax(x)
    For y = 1 To 20
'      If y < 18 Then
        LeaveRec.VEntry(y).EARN = TempVEntry(x, y).EARN
        LeaveRec.SEntry(y).EARN = TempSEntry(x, y).EARN
        LeaveRec.HEntry(y).EARN = 0 'TempHEntry(x, y).EARN
        LeaveRec.PEntry(y).EARN = 0 'TempPEntry(x, y).EARN
        LeaveRec.VEntry(y).YEARS = TempVEntry(x, y).YEARS
        LeaveRec.SEntry(y).YEARS = TempSEntry(x, y).YEARS
        LeaveRec.HEntry(y).YEARS = 0 'TempHEntry(x, y).YEARS
        LeaveRec.PEntry(y).YEARS = 0 'TempPEntry(x, y).YEARS
'      Else
'        LeaveRec.VEntry(y).EARN = 0
'        LeaveRec.SEntry(y).EARN = 0
'        LeaveRec.HEntry(y).EARN = 0
'        LeaveRec.PEntry(y).EARN = 0
'        LeaveRec.VEntry(y).YEARS = 0
'        LeaveRec.SEntry(y).YEARS = 0
'        LeaveRec.HEntry(y).YEARS = 0
'        LeaveRec.PEntry(y).YEARS = 0
'      End If
    Next y
    Put NewHandle, x, LeaveRec
  Next x
  
  Close NewHandle
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
  For x = 1 To NumOfErns
    Get DosErnHandle, x, DosErnRec
    TempERNCODE1(x) = QPTrim$(DosErnRec.ERNCODE1)
    TempERNFWT1(x) = DosErnRec.ERNFWT1
    TempERNSWT1(x) = DosErnRec.ERNSWT1
    TempERNSOC1(x) = DosErnRec.ERNSOC1
    TempERNMED1(x) = DosErnRec.ERNMED1
    TempERNRET1(x) = DosErnRec.ERNRET1
  Next x
  Close DosErnHandle
  
  OpenErnCodeFile ErnHandle
  For x = 1 To NumOfErns
    ErnRec.ERNCODE1 = QPTrim$(TempERNCODE1(x))
    ErnRec.ERNFWT1 = TempERNFWT1(x)
    ErnRec.ERNSWT1 = TempERNSWT1(x)
    ErnRec.ERNSOC1 = TempERNSOC1(x)
    ErnRec.ERNMED1 = TempERNMED1(x)
    ErnRec.ERNRET1 = TempERNRET1(x)
    ErnRec.EarnYN = "Y" '"Y"es include
    ErnRec.Pad = ""
    Put ErnHandle, x, ErnRec
  Next x
  Close ErnHandle
    
NoErnMatchNeeded:

  OpenOldTransHistFile OldTransHistHandle
  NumOfTransHistRecs = LOF(OldTransHistHandle) / Len(OldTransHistRec)
  
  ReDim TempTActive(1 To NumOfTransHistRecs) As Integer
  ReDim TempPrevTransRec(1 To NumOfTransHistRecs) As Integer
  ReDim TempEmpPin(1 To NumOfTransHistRecs) As Integer
  ReDim TempPaySFlag(1 To NumOfTransHistRecs) As String * 1
  ReDim TempCheckNum(1 To NumOfTransHistRecs) As Long
  ReDim TempPayPdStart(1 To NumOfTransHistRecs) As Integer
  ReDim TempPayPdEnd(1 To NumOfTransHistRecs) As Integer
  ReDim TempCheckDate(1 To NumOfTransHistRecs) As Integer
  ReDim TempPostDate(1 To NumOfTransHistRecs) As Integer
  ReDim TempPayType(1 To NumOfTransHistRecs) As String * 1
  ReDim TempBaseRate(1 To NumOfTransHistRecs) As Double
  ReDim TempOTRate(1 To NumOfTransHistRecs) As Double
  ReDim TempRegHrsWork(1 To NumOfTransHistRecs) As Double
  ReDim TempVacUsed(1 To NumOfTransHistRecs) As Double
  ReDim TempSickUsed(1 To NumOfTransHistRecs) As Double
  ReDim TempCompUsed(1 To NumOfTransHistRecs) As Double

  ReDim TempHOLHOURS(1 To NumOfTransHistRecs) As Double
  ReDim TempPerHours(1 To NumOfTransHistRecs) As Double

  ReDim TempRegHrsPaid(1 To NumOfTransHistRecs) As Double
  ReDim TempOTHours(1 To NumOfTransHistRecs) As Double
  ReDim TempOTHrsPaid(1 To NumOfTransHistRecs) As Double
  ReDim TempOT2Comp(1 To NumOfTransHistRecs) As Double
  ReDim TempTDist(1 To NumOfTransHistRecs, 1 To 8) As TransWageDistType
                              
  ReDim TempTotRegWage(1 To NumOfTransHistRecs) As Double
  ReDim TempTotOTWage(1 To NumOfTransHistRecs) As Double
  ReDim TempGrossWage(1 To NumOfTransHistRecs) As Double
  ReDim TempEAmt(1 To NumOfTransHistRecs, 1 To 3) As Double
                                     
  ReDim TempEDist(1 To NumOfTransHistRecs, 1 To 6) As TransEarnDistType
  ReDim TempTotAdditEarn(1 To NumOfTransHistRecs) As Double
  ReDim TempGrossPay(1 To NumOfTransHistRecs) As Double
  ReDim TempSocGrossPay(1 To NumOfTransHistRecs) As Double
  ReDim TempMedGrossPay(1 To NumOfTransHistRecs) As Double
  ReDim TempFedGrossPay(1 To NumOfTransHistRecs) As Double
  ReDim TempStaGrossPay(1 To NumOfTransHistRecs) As Double
  ReDim TempSocTaxAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempMedTaxAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempFedTaxAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempStaTaxAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempTotTaxAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempRetireAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempDAmt(1 To NumOfTransHistRecs, 1 To 50) As Double
  ReDim TempTotDedAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempEICAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempNetPay(1 To NumOfTransHistRecs) As Double
  ReDim TempPeriodHistRec(1 To NumOfTransHistRecs) As Integer
  ReDim TempMatchRetAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempMatchSocAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempMatchMedAmt(1 To NumOfTransHistRecs) As Double
  ReDim TempRetGrossPay(1 To NumOfTransHistRecs) As Double
  ReDim TempTaxFring(1 To NumOfTransHistRecs) As Double
  ReDim TempLess401k(1 To NumOfTransHistRecs) As Boolean
  ReDim TempPad1(1 To NumOfTransHistRecs) As String
  
  For x = 1 To NumOfTransHistRecs
    Get OldTransHistHandle, x, OldTransHistRec
    TempTActive(x) = OldTransHistRec.TActive
    TempPrevTransRec(x) = OldTransHistRec.PrevTransRec
    TempEmpPin(x) = OldTransHistRec.EmpPin
    TempPaySFlag(x) = QPTrim$(OldTransHistRec.PaySFlag)
    TempCheckNum(x) = OldTransHistRec.CheckNum
    TempPayPdStart(x) = OldTransHistRec.PayPdStart
    TempPayPdEnd(x) = OldTransHistRec.PayPdEnd
    TempCheckDate(x) = OldTransHistRec.CheckDate
    TempPostDate(x) = OldTransHistRec.PostDate
    TempPayType(x) = QPTrim$(OldTransHistRec.PayType)
    TempBaseRate(x) = OldTransHistRec.BaseRate
    TempOTRate(x) = OldTransHistRec.OTRate
    TempRegHrsWork(x) = OldTransHistRec.RegHrsWork
    TempVacUsed(x) = OldTransHistRec.VacUsed
    TempSickUsed(x) = OldTransHistRec.SickUsed
    TempCompUsed(x) = OldTransHistRec.CompUsed
  
    TempHOLHOURS(x) = OldTransHistRec.HOLHOURS
    TempPerHours(x) = OldTransHistRec.PerHours
  
    TempRegHrsPaid(x) = OldTransHistRec.RegHrsPaid
    TempOTHours(x) = OldTransHistRec.OTHours
    TempOTHrsPaid(x) = OldTransHistRec.OTHrsPaid
    TempOT2Comp(x) = OldTransHistRec.OT2Comp
    For y = 1 To 8
      TempTDist(x, y).DAcct = QPTrim$(OldTransHistRec.TDist(y).DAcct)
      TempTDist(x, y).DRHrs = OldTransHistRec.TDist(y).DRHrs
      TempTDist(x, y).DOHrs = OldTransHistRec.TDist(y).DOHrs
      TempTDist(x, y).DPct = OldTransHistRec.TDist(y).DPct
      TempTDist(x, y).DRWage = OldTransHistRec.TDist(y).DRWage
      TempTDist(x, y).DOWage = OldTransHistRec.TDist(y).DOWage
    Next y
                                
    TempTotRegWage(x) = OldTransHistRec.TotRegWage
    TempTotOTWage(x) = OldTransHistRec.TotOTWage
    TempGrossWage(x) = OldTransHistRec.GrossWage
    For y = 1 To 3
      TempEAmt(x, y) = OldTransHistRec.EAmt(y)
    Next y
    For y = 1 To 6
      TempEDist(x, y).EAcct = QPTrim$(OldTransHistRec.EDist(y).EAcct)
      TempEDist(x, y).EAmt = OldTransHistRec.EDist(y).EAmt
    Next y
    
    TempTotAdditEarn(x) = OldTransHistRec.TotAdditEarn
    TempGrossPay(x) = OldTransHistRec.GrossPay
    TempSocGrossPay(x) = OldTransHistRec.SocGrossPay
    TempMedGrossPay(x) = OldTransHistRec.MedGrossPay
    TempFedGrossPay(x) = OldTransHistRec.FedGrossPay
    TempStaGrossPay(x) = OldTransHistRec.StaGrossPay
    TempSocTaxAmt(x) = OldTransHistRec.SocTaxAmt
    TempMedTaxAmt(x) = OldTransHistRec.MedTaxAmt
    TempFedTaxAmt(x) = OldTransHistRec.FedTaxAmt
    TempStaTaxAmt(x) = OldTransHistRec.StaTaxAmt
    TempTotTaxAmt(x) = OldTransHistRec.TotTaxAmt
    TempRetireAmt(x) = OldTransHistRec.RetireAmt
    For y = 1 To 50
      TempDAmt(x, y) = OldTransHistRec.DAmt(y)
    Next y
    TempTotDedAmt(x) = OldTransHistRec.TotDedAmt
    TempEICAmt(x) = OldTransHistRec.EICAmt
    TempNetPay(x) = OldTransHistRec.NetPay
    TempPeriodHistRec(x) = OldTransHistRec.PeriodHistRec
    TempMatchRetAmt(x) = OldTransHistRec.MatchRetAmt
    TempMatchSocAmt(x) = OldTransHistRec.MatchSocAmt
    TempMatchMedAmt(x) = OldTransHistRec.MatchMedAmt
    TempRetGrossPay(x) = OldTransHistRec.RetGrossPay
    TempTaxFring(x) = OldTransHistRec.TaxFring
    TempLess401k(x) = OldTransHistRec.Less401k
    TempPad1(x) = QPTrim$(OldTransHistRec.Pad1)
  Next x
  Close OldTransHistHandle
  
  OpenTransHistFile TransHistHandle
  For x = 1 To NumOfTransHistRecs
    TransHistRec.TActive = TempTActive(x)
    TransHistRec.PrevTransRec = TempPrevTransRec(x)
    TransHistRec.EmpPin = TempEmpPin(x)
    TransHistRec.PaySFlag = QPTrim$(TempPaySFlag(x))
    TransHistRec.CheckNum = TempCheckNum(x)
    TransHistRec.PayPdStart = TempPayPdStart(x)
    TransHistRec.PayPdEnd = TempPayPdEnd(x)
    TransHistRec.CheckDate = TempCheckDate(x)
    TransHistRec.PostDate = TempPostDate(x)
    TransHistRec.PayType = QPTrim$(TempPayType(x))
    TransHistRec.BaseRate = TempBaseRate(x)
    TransHistRec.OTRate = TempOTRate(x)
    TransHistRec.RegHrsWork = TempRegHrsWork(x)
    TransHistRec.VacUsed = TempVacUsed(x)
    TransHistRec.SickUsed = TempSickUsed(x)
    TransHistRec.CompUsed = TempCompUsed(x)
  
    TransHistRec.HOLHOURS = TempHOLHOURS(x)
    TransHistRec.PerHours = TempPerHours(x)
  
    TransHistRec.RegHrsPaid = TempRegHrsPaid(x)
    TransHistRec.OTHours = TempOTHours(x)
    TransHistRec.OTHrsPaid = TempOTHrsPaid(x)
    TransHistRec.OT2Comp = TempOT2Comp(x)
    For y = 1 To 8
      TransHistRec.TDist(y).DAcct = TempTDist(x, y).DAcct
      TransHistRec.TDist(y).DRHrs = TempTDist(x, y).DRHrs
      TransHistRec.TDist(y).DOHrs = TempTDist(x, y).DOHrs
      TransHistRec.TDist(y).DPct = TempTDist(x, y).DPct
      TransHistRec.TDist(y).DRWage = TempTDist(x, y).DRWage
      TransHistRec.TDist(y).DOWage = TempTDist(x, y).DOWage
    Next y
                                
    TransHistRec.TotRegWage = TempTotRegWage(x)
    TransHistRec.TotOTWage = TempTotOTWage(x)
    TransHistRec.GrossWage = TempGrossWage(x)
    For y = 1 To 3
      TransHistRec.EAmt(y) = TempEAmt(x, y)
    Next y
    For y = 1 To 6
      TransHistRec.EDist(y).EAcct = QPTrim$(TempEDist(x, y).EAcct)
      TransHistRec.EDist(y).EAmt = TempEDist(x, y).EAmt
    Next y
    
    TransHistRec.TotAdditEarn = TempTotAdditEarn(x)
    TransHistRec.GrossPay = TempGrossPay(x)
    TransHistRec.SocGrossPay = TempSocGrossPay(x)
    TransHistRec.MedGrossPay = TempMedGrossPay(x)
    TransHistRec.FedGrossPay = TempFedGrossPay(x)
    TransHistRec.StaGrossPay = TempStaGrossPay(x)
    TransHistRec.SocTaxAmt = TempSocTaxAmt(x)
    TransHistRec.MedTaxAmt = TempMedTaxAmt(x)
    TransHistRec.FedTaxAmt = TempFedTaxAmt(x)
    TransHistRec.StaTaxAmt = TempStaTaxAmt(x)
    TransHistRec.TotTaxAmt = TempTotTaxAmt(x)
    TransHistRec.RetireAmt = TempRetireAmt(x)
    For y = 1 To 50
      TransHistRec.DAmt(y) = TempDAmt(x, y)
    Next y
    TransHistRec.TotDedAmt = TempTotDedAmt(x)
    TransHistRec.EICAmt = TempEICAmt(x)
    TransHistRec.NetPay = TempNetPay(x)
    TransHistRec.PeriodHistRec = TempPeriodHistRec(x)
    TransHistRec.MatchRetAmt = TempMatchRetAmt(x)
    TransHistRec.MatchSocAmt = TempMatchSocAmt(x)
    TransHistRec.MatchMedAmt = TempMatchMedAmt(x)
    TransHistRec.RetGrossPay = TempRetGrossPay(x)
    TransHistRec.TaxFring = TempTaxFring(x)
    'exclusive with Wrightsville Beach...if an employee
    'is paid solely out of alternate earnings then by default
    'none of his wages for that pay period get matching 401K
    'amounts...so in this conversion we look for any pay transaction
    'showing zero Regular Wage Amounts and assign true to all
    '3 alternate earnings .Less401k fields indicating that they
    'are not to be matched
    If TransHistRec.TotRegWage > 0 Then
      For y = 1 To 3
        TransHistRec.Less401k(y) = False
      Next y
    Else
      For y = 1 To 3
        TransHistRec.Less401k(y) = True
      Next y
    End If
    TransHistRec.Pad1 = QPTrim$(TempPad1(x))
    Put TransHistHandle, x, TransHistRec
  Next x
  Close TransHistHandle
  
  OpenOldTransWorkFile OldTransWorkHandle
  NumOfTransWorkRecs = LOF(OldTransWorkHandle) / Len(OldTransWorkRec)
  
  ReDim TempTActive(1 To NumOfTransWorkRecs) As Integer
  ReDim TempPrevTransRec(1 To NumOfTransWorkRecs) As Integer
  ReDim TempEmpPin(1 To NumOfTransWorkRecs) As Integer
  ReDim TempPaySFlag(1 To NumOfTransWorkRecs) As String * 1
  ReDim TempCheckNum(1 To NumOfTransWorkRecs) As Long
  ReDim TempPayPdStart(1 To NumOfTransWorkRecs) As Integer
  ReDim TempPayPdEnd(1 To NumOfTransWorkRecs) As Integer
  ReDim TempCheckDate(1 To NumOfTransWorkRecs) As Integer
  ReDim TempPostDate(1 To NumOfTransWorkRecs) As Integer
  ReDim TempPayType(1 To NumOfTransWorkRecs) As String * 1
  ReDim TempBaseRate(1 To NumOfTransWorkRecs) As Double
  ReDim TempOTRate(1 To NumOfTransWorkRecs) As Double
  ReDim TempRegHrsWork(1 To NumOfTransWorkRecs) As Double
  ReDim TempVacUsed(1 To NumOfTransWorkRecs) As Double
  ReDim TempSickUsed(1 To NumOfTransWorkRecs) As Double
  ReDim TempCompUsed(1 To NumOfTransWorkRecs) As Double

  ReDim TempHOLHOURS(1 To NumOfTransWorkRecs) As Double
  ReDim TempPerHours(1 To NumOfTransWorkRecs) As Double

  ReDim TempRegHrsPaid(1 To NumOfTransWorkRecs) As Double
  ReDim TempOTHours(1 To NumOfTransWorkRecs) As Double
  ReDim TempOTHrsPaid(1 To NumOfTransWorkRecs) As Double
  ReDim TempOT2Comp(1 To NumOfTransWorkRecs) As Double
  ReDim TempTDist(1 To NumOfTransWorkRecs, 1 To 8) As TransWageDistType
                              
  ReDim TempTotRegWage(1 To NumOfTransWorkRecs) As Double
  ReDim TempTotOTWage(1 To NumOfTransWorkRecs) As Double
  ReDim TempGrossWage(1 To NumOfTransWorkRecs) As Double
  ReDim TempEAmt(1 To NumOfTransWorkRecs, 1 To 3) As Double
                                     
  ReDim TempEDist(1 To NumOfTransWorkRecs, 1 To 6) As TransEarnDistType
  ReDim TempTotAdditEarn(1 To NumOfTransWorkRecs) As Double
  ReDim TempGrossPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempSocGrossPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempMedGrossPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempFedGrossPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempStaGrossPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempSocTaxAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempMedTaxAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempFedTaxAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempStaTaxAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempTotTaxAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempRetireAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempDAmt(1 To NumOfTransWorkRecs, 1 To 50) As Double
  ReDim TempTotDedAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempEICAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempNetPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempPeriodWorkRec(1 To NumOfTransWorkRecs) As Integer
  ReDim TempMatchRetAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempMatchSocAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempMatchMedAmt(1 To NumOfTransWorkRecs) As Double
  ReDim TempRetGrossPay(1 To NumOfTransWorkRecs) As Double
  ReDim TempTaxFring(1 To NumOfTransWorkRecs) As Double
  ReDim TempLess401k(1 To NumOfTransWorkRecs) As Boolean
  ReDim TempPad1(1 To NumOfTransWorkRecs) As String
  
  For x = 1 To NumOfTransWorkRecs
    Get OldTransWorkHandle, x, OldTransWorkRec
    TempTActive(x) = OldTransWorkRec.TActive
    TempPrevTransRec(x) = OldTransWorkRec.PrevTransRec
    TempEmpPin(x) = OldTransWorkRec.EmpPin
    TempPaySFlag(x) = QPTrim$(OldTransWorkRec.PaySFlag)
    TempCheckNum(x) = OldTransWorkRec.CheckNum
    TempPayPdStart(x) = OldTransWorkRec.PayPdStart
    TempPayPdEnd(x) = OldTransWorkRec.PayPdEnd
    TempCheckDate(x) = OldTransWorkRec.CheckDate
    TempPostDate(x) = OldTransWorkRec.PostDate
    TempPayType(x) = QPTrim$(OldTransWorkRec.PayType)
    TempBaseRate(x) = OldTransWorkRec.BaseRate
    TempOTRate(x) = OldTransWorkRec.OTRate
    TempRegHrsWork(x) = OldTransWorkRec.RegHrsWork
    TempVacUsed(x) = OldTransWorkRec.VacUsed
    TempSickUsed(x) = OldTransWorkRec.SickUsed
    TempCompUsed(x) = OldTransWorkRec.CompUsed
  
    TempHOLHOURS(x) = OldTransWorkRec.HOLHOURS
    TempPerHours(x) = OldTransWorkRec.PerHours
  
    TempRegHrsPaid(x) = OldTransWorkRec.RegHrsPaid
    TempOTHours(x) = OldTransWorkRec.OTHours
    TempOTHrsPaid(x) = OldTransWorkRec.OTHrsPaid
    TempOT2Comp(x) = OldTransWorkRec.OT2Comp
    For y = 1 To 8
      TempTDist(x, y).DAcct = QPTrim$(OldTransWorkRec.TDist(y).DAcct)
      TempTDist(x, y).DRHrs = OldTransWorkRec.TDist(y).DRHrs
      TempTDist(x, y).DOHrs = OldTransWorkRec.TDist(y).DOHrs
      TempTDist(x, y).DPct = OldTransWorkRec.TDist(y).DPct
      TempTDist(x, y).DRWage = OldTransWorkRec.TDist(y).DRWage
      TempTDist(x, y).DOWage = OldTransWorkRec.TDist(y).DOWage
    Next y
                                
    TempTotRegWage(x) = OldTransWorkRec.TotRegWage
    TempTotOTWage(x) = OldTransWorkRec.TotOTWage
    TempGrossWage(x) = OldTransWorkRec.GrossWage
    For y = 1 To 3
      TempEAmt(x, y) = OldTransWorkRec.EAmt(y)
    Next y
    For y = 1 To 6
      TempEDist(x, y).EAcct = QPTrim$(OldTransWorkRec.EDist(y).EAcct)
      TempEDist(x, y).EAmt = OldTransWorkRec.EDist(y).EAmt
    Next y
    
    TempTotAdditEarn(x) = OldTransWorkRec.TotAdditEarn
    TempGrossPay(x) = OldTransWorkRec.GrossPay
    TempSocGrossPay(x) = OldTransWorkRec.SocGrossPay
    TempMedGrossPay(x) = OldTransWorkRec.MedGrossPay
    TempFedGrossPay(x) = OldTransWorkRec.FedGrossPay
    TempStaGrossPay(x) = OldTransWorkRec.StaGrossPay
    TempSocTaxAmt(x) = OldTransWorkRec.SocTaxAmt
    TempMedTaxAmt(x) = OldTransWorkRec.MedTaxAmt
    TempFedTaxAmt(x) = OldTransWorkRec.FedTaxAmt
    TempStaTaxAmt(x) = OldTransWorkRec.StaTaxAmt
    TempTotTaxAmt(x) = OldTransWorkRec.TotTaxAmt
    TempRetireAmt(x) = OldTransWorkRec.RetireAmt
    For y = 1 To 50
      TempDAmt(x, y) = OldTransWorkRec.DAmt(y)
    Next y
    TempTotDedAmt(x) = OldTransWorkRec.TotDedAmt
    TempEICAmt(x) = OldTransWorkRec.EICAmt
    TempNetPay(x) = OldTransWorkRec.NetPay
    TempPeriodHistRec(x) = OldTransWorkRec.PeriodHistRec
    TempMatchRetAmt(x) = OldTransWorkRec.MatchRetAmt
    TempMatchSocAmt(x) = OldTransWorkRec.MatchSocAmt
    TempMatchMedAmt(x) = OldTransWorkRec.MatchMedAmt
    TempRetGrossPay(x) = OldTransWorkRec.RetGrossPay
    TempTaxFring(x) = OldTransWorkRec.TaxFring
    TempLess401k(x) = OldTransWorkRec.Less401k
    TempPad1(x) = QPTrim$(OldTransWorkRec.Pad1)
  Next x
  Close OldTransWorkHandle
  
  OpenTransWorkFile TransWorkHandle
  For x = 1 To NumOfTransWorkRecs
    TransWorkRec.TActive = TempTActive(x)
    TransWorkRec.PrevTransRec = TempPrevTransRec(x)
    TransWorkRec.EmpPin = TempEmpPin(x)
    TransWorkRec.PaySFlag = QPTrim$(TempPaySFlag(x))
    TransWorkRec.CheckNum = TempCheckNum(x)
    TransWorkRec.PayPdStart = TempPayPdStart(x)
    TransWorkRec.PayPdEnd = TempPayPdEnd(x)
    TransWorkRec.CheckDate = TempCheckDate(x)
    TransWorkRec.PostDate = TempPostDate(x)
    TransWorkRec.PayType = QPTrim$(TempPayType(x))
    TransWorkRec.BaseRate = TempBaseRate(x)
    TransWorkRec.OTRate = TempOTRate(x)
    TransWorkRec.RegHrsWork = TempRegHrsWork(x)
    TransWorkRec.VacUsed = TempVacUsed(x)
    TransWorkRec.SickUsed = TempSickUsed(x)
    TransWorkRec.CompUsed = TempCompUsed(x)
  
    TransWorkRec.HOLHOURS = TempHOLHOURS(x)
    TransWorkRec.PerHours = TempPerHours(x)
  
    TransWorkRec.RegHrsPaid = TempRegHrsPaid(x)
    TransWorkRec.OTHours = TempOTHours(x)
    TransWorkRec.OTHrsPaid = TempOTHrsPaid(x)
    TransWorkRec.OT2Comp = TempOT2Comp(x)
    For y = 1 To 8
      TransWorkRec.TDist(y).DAcct = TempTDist(x, y).DAcct
      TransWorkRec.TDist(y).DRHrs = TempTDist(x, y).DRHrs
      TransWorkRec.TDist(y).DOHrs = TempTDist(x, y).DOHrs
      TransWorkRec.TDist(y).DPct = TempTDist(x, y).DPct
      TransWorkRec.TDist(y).DRWage = TempTDist(x, y).DRWage
      TransWorkRec.TDist(y).DOWage = TempTDist(x, y).DOWage
    Next y
                                
    TransWorkRec.TotRegWage = TempTotRegWage(x)
    TransWorkRec.TotOTWage = TempTotOTWage(x)
    TransWorkRec.GrossWage = TempGrossWage(x)
    For y = 1 To 3
      TransWorkRec.EAmt(y) = TempEAmt(x, y)
    Next y
    For y = 1 To 6
      TransWorkRec.EDist(y).EAcct = QPTrim$(TempEDist(x, y).EAcct)
      TransWorkRec.EDist(y).EAmt = TempEDist(x, y).EAmt
    Next y
    
    TransWorkRec.TotAdditEarn = TempTotAdditEarn(x)
    TransWorkRec.GrossPay = TempGrossPay(x)
    TransWorkRec.SocGrossPay = TempSocGrossPay(x)
    TransWorkRec.MedGrossPay = TempMedGrossPay(x)
    TransWorkRec.FedGrossPay = TempFedGrossPay(x)
    TransWorkRec.StaGrossPay = TempStaGrossPay(x)
    TransWorkRec.SocTaxAmt = TempSocTaxAmt(x)
    TransWorkRec.MedTaxAmt = TempMedTaxAmt(x)
    TransWorkRec.FedTaxAmt = TempFedTaxAmt(x)
    TransWorkRec.StaTaxAmt = TempStaTaxAmt(x)
    TransWorkRec.TotTaxAmt = TempTotTaxAmt(x)
    TransWorkRec.RetireAmt = TempRetireAmt(x)
    For y = 1 To 50
      TransWorkRec.DAmt(y) = TempDAmt(x, y)
    Next y
    TransWorkRec.TotDedAmt = TempTotDedAmt(x)
    TransWorkRec.EICAmt = TempEICAmt(x)
    TransWorkRec.NetPay = TempNetPay(x)
    TransWorkRec.PeriodHistRec = TempPeriodHistRec(x)
    TransWorkRec.MatchRetAmt = TempMatchRetAmt(x)
    TransWorkRec.MatchSocAmt = TempMatchSocAmt(x)
    TransWorkRec.MatchMedAmt = TempMatchMedAmt(x)
    TransWorkRec.RetGrossPay = TempRetGrossPay(x)
    TransWorkRec.TaxFring = TempTaxFring(x)
    For y = 1 To 3
      TransWorkRec.Less401k(y) = False
    Next y
    TransWorkRec.Pad1 = QPTrim$(TempPad1(x))
    Put TransWorkHandle, x, TransWorkRec
  Next x
  Close TransWorkHandle
  
  OpenEmpData2File EmpHandle
  NumOfEmpRecs = LOF(EmpHandle) / Len(Emp2Rec)
  For x = 1 To NumOfEmpRecs
    Get EmpHandle, x, Emp2Rec
    If QPTrim$(Emp2Rec.EMPRETTP) = "" Then
      Emp2Rec.YN401K = "N"
      Put EmpHandle, x, Emp2Rec
    End If
  Next x
  Close EmpHandle
  
  Label2.Visible = True
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Label2.Visible = False
  Label7.Visible = False
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

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
      SendKeys "%x"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub


