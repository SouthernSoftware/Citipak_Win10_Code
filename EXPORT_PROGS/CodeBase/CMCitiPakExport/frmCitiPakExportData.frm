VERSION 5.00
Begin VB.Form frmCitiPakExportData 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Citi-Pak Export"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmCitiPakExportData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.ListBox lstReqFilesGL 
         Enabled         =   0   'False
         Height          =   1815
         ItemData        =   "frmCitiPakExportData.frx":08CA
         Left            =   -17789
         List            =   "frmCitiPakExportData.frx":08D1
         TabIndex        =   19
         Top             =   -17369
         Width           =   2745
      End
      Begin VB.ListBox lstMissingFilesListGL 
         Enabled         =   0   'False
         Height          =   1815
         ItemData        =   "frmCitiPakExportData.frx":08D8
         Left            =   -17789
         List            =   "frmCitiPakExportData.frx":08DF
         TabIndex        =   18
         Top             =   -17369
         Width           =   2745
      End
      Begin VB.ListBox lstMissingFilesListPayroll 
         Height          =   1815
         ItemData        =   "frmCitiPakExportData.frx":08E6
         Left            =   6675
         List            =   "frmCitiPakExportData.frx":08ED
         TabIndex        =   11
         Top             =   1290
         Width           =   2745
      End
      Begin VB.ListBox lstReqFilesPayroll 
         Height          =   1815
         ItemData        =   "frmCitiPakExportData.frx":08F4
         Left            =   1440
         List            =   "frmCitiPakExportData.frx":08FB
         TabIndex        =   10
         Top             =   1350
         Width           =   2745
      End
      Begin VB.CommandButton cmdProcess 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Export Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7785
         TabIndex        =   17
         Top             =   4500
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Modules"
         Height          =   2415
         Left            =   180
         TabIndex        =   1
         Top             =   3675
         Width           =   10935
         Begin VB.CheckBox chkTaxBilling 
            Caption         =   "Tax Billing"
            Height          =   495
            Left            =   3360
            TabIndex        =   15
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox chkCashMgt 
            Caption         =   "Payment central"
            Height          =   495
            Left            =   1440
            TabIndex        =   14
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox chkUB 
            Caption         =   "Utility Billing"
            Height          =   495
            Left            =   1440
            TabIndex        =   8
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkBL 
            Caption         =   "Business Liscense"
            Height          =   495
            Left            =   3360
            TabIndex        =   7
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox chkAp 
            Caption         =   "AP"
            Height          =   495
            Left            =   1440
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkVehDec 
            Caption         =   "Vehical Decals"
            Height          =   495
            Left            =   3360
            TabIndex        =   5
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkFixedAssets 
            Caption         =   "Fixed Assets"
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox chkGeneralLedger 
            Caption         =   "GL"
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkPayroll 
            Caption         =   "Payroll"
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Required files"
         Enabled         =   0   'False
         Height          =   210
         Left            =   -17204
         TabIndex        =   9
         Top             =   -15764
         Width           =   2160
      End
      Begin VB.Label Label3 
         Caption         =   "Missing Files"
         Enabled         =   0   'False
         Height          =   210
         Left            =   -17204
         TabIndex        =   20
         Top             =   -15764
         Width           =   2160
      End
      Begin VB.Label Label2 
         Caption         =   "Missing Files"
         Height          =   210
         Left            =   3210
         TabIndex        =   13
         Top             =   555
         Width           =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "Required files"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   555
         Width           =   2160
      End
      Begin VB.Label lblInfo 
         Caption         =   " "
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   960
         Width           =   10575
      End
   End
End
Attribute VB_Name = "frmCitiPakExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Dim ErrorCode As Integer


Private Sub chkCashMgt_Click()
   Call FillFilesToConvertList
End Sub


'its here dale

Private Sub cmdProcess_Click()
  Dim RptNamex As String
  Dim RptHandlex As Integer
  Dim ThisFile As String
  lblInfo.FontSize = 12
  ErrorCode = 0

  ValidateFilesExists

  If ErrorCode = 0 Then
    If chkPayroll.Value = 1 Then
      chkPayroll.ForeColor = vbBlue
      Call validateDedCodes
      Call ValidateErnCodes
      If ErrorCode = 0 Then
        Call ProcessPayrollTransHist
        Call ProcessEmployeeData
        Call ProcessPrSys
        Call ProcessEIC1RecType
        Call ProcessRetireRecType
        Call ProcessUnitFileRecType
        Call ProcessDraftInfo
        Call ProcessErnCodeRecType
        Call ProcessDedCodeRecType
        Call ProcessAccrualDates
        Call ProcessPayRateType
        Call ProcessPRMessRecType
        Call ProcessOrbitEmpData
        Call ProcessVoidedCheckType
        Call ProcessW2ElectronicsSubRa
        Call ProcessK401DedType
        Call ProcessLeaveBenefits
      End If
    End If
  
    If chkGeneralLedger.Value = 1 Then
      chkGeneralLedger.ForeColor = vbBlue
      Call ProcessGlFundRecType
      Call ProcessGlAcctRec
      Call ProcessGLTrans
      Call ProcessGLBudgetTrans
      Call ProcessGlAcctRecForBudgetPrep
    End If
  
    If chkFixedAssets.Value = 1 Then
      Call ProcessFAData
    End If
    
    If chkVehDec.Value = 1 Then
      Call ProcessDCData
    End If
    
    If chkCashMgt.Value = 1 Then
      Call ProcessCMData
    End If
    
    cmdProcess.Enabled = False
    
    lblInfo.Caption = "You may Now exit, export completed"
  
  End If
  Close RptHandlex

End Sub
Private Sub ProcessPayrollTransHist()
 
End Sub

Private Sub ReadEmpInfoArray(ByRef empList() As EmpListArrayType)
 
End Sub

Private Sub ProcessEmployeeData()
 
End Sub
Private Sub ProcessPrSys()

End Sub

Private Sub ProcessEIC1RecType()

End Sub

Private Sub ProcessRetireRecType()
 
End Sub

Private Sub ProcessUnitFileRecType()
End Sub
Private Sub ProcessDraftInfo()
End Sub

Private Sub ProcessErnCodeRecType()
End Sub
Private Sub ProcessDedCodeRecType()
End Sub
Private Sub ProcessAccrualDates()
End Sub

Private Sub ProcessPayRateType()
 End Sub
 Private Sub ProcessPRMessRecType()
 End Sub
 
 Private Sub ProcessOrbitEmpData()
 End Sub
Private Sub ProcessVoidedCheckType()
  
  
End Sub

Private Sub ProcessW2ElectronicsSubRa()
End Sub
 
Private Sub validateDedCodes()
End Sub

Private Sub ValidateErnCodes()
  
End Sub
 
Private Sub chkFixedAssets_Click()
    Call FillFilesToConvertList
End Sub

Private Sub chkPayroll_Click()
    Call FillFilesToConvertList
End Sub

Private Sub chkVehDec_Click()
    Call FillFilesToConvertList
End Sub

Private Sub FillFilesToConvertList()
  lstReqFilesPayroll.Clear
  If chkPayroll.Value = 1 Then
    lstReqFilesPayroll.AddItem ("PREMP1.DAT                Payroll")
    lstReqFilesPayroll.AddItem ("PREMP2.DAT")
    lstReqFilesPayroll.AddItem ("PREMP3.DAT")
    lstReqFilesPayroll.AddItem ("PRTRANSH.DAT")
    lstReqFilesPayroll.AddItem ("PRSYS.DAT")
    lstReqFilesPayroll.AddItem ("PREICTBL.DAT")
    lstReqFilesPayroll.AddItem ("PRRETIRE.DAT")
    lstReqFilesPayroll.AddItem ("PRUNIT.DAT")
    lstReqFilesPayroll.AddItem ("PRDRAFTI.DAT")
    lstReqFilesPayroll.AddItem ("PRERNCOD.DAT")
    lstReqFilesPayroll.AddItem ("PRDEDCOD.DAT")
    lstReqFilesPayroll.AddItem ("PRACCRUE.DAT")
    lstReqFilesPayroll.AddItem ("PAYRATE.DAT")
    lstReqFilesPayroll.AddItem ("EMPMESS.DAT")
    lstReqFilesPayroll.AddItem ("OrbEmpData.DAT")
    lstReqFilesPayroll.AddItem ("TEMPVOID.DAT")
    lstReqFilesPayroll.AddItem ("W2ESUBRA.DAT")
  End If
  If chkFixedAssets = 1 Then
    lstReqFilesPayroll.AddItem (FASetUpFileName + "      Fixed Assets")
    lstReqFilesPayroll.AddItem (FAItemFileName)
    lstReqFilesPayroll.AddItem (FAAssetCodeName)
    lstReqFilesPayroll.AddItem (FADeptCodeName)
    lstReqFilesPayroll.AddItem (FAFundCodeName)
    lstReqFilesPayroll.AddItem (FADprHistFileName)
  End If
  If chkVehDec = 1 Then
    lstReqFilesPayroll.AddItem (DCSetupFile + "     Vehicle Decal")
    lstReqFilesPayroll.AddItem (DCTranFile)
    lstReqFilesPayroll.AddItem (DCCustFile)
    lstReqFilesPayroll.AddItem (DCVCodeFile)
    lstReqFilesPayroll.AddItem (DCVehFile)
  End If
  
  If chkCashMgt.Value = 1 Then
    lstReqFilesPayroll.AddItem (CMTranFile + "   Cash Management")
  End If

End Sub

Private Sub ProcessGlFundRecType()
 End Sub
 
 Private Sub ProcessGlAcctRec()
  
 End Sub
 Private Sub ProcessGLTrans()
End Sub
Private Sub ProcessGLBudgetTrans()
End Sub
 Private Sub ProcessGlAcctRecForBudgetPrep()
End Sub
Private Sub ProcessK401DedType()
 End Sub
 Private Sub ProcessLeaveBenefits()
 End Sub
Private Sub chkGeneralLedger_Click()
  If chkGeneralLedger.Value = 1 Then
    lstReqFilesGL.Clear
    lstReqFilesGL.AddItem ("GlFund.dat")
    lstReqFilesGL.AddItem ("GLAcct.dat")
    lstReqFilesGL.AddItem ("GlTrans.dat")
  Else
    lstReqFilesGL.Clear
  End If
End Sub
 
 Private Sub ValidateFilesExists()
 Dim bolModuleSelected As Boolean
 bolModuleSelected = False
 lstMissingFilesListPayroll.Clear
 
 'Validate Payroll Files if the user selects the payroll module
  
'Validate GL files exists
  
'Validate FA files exists
  If chkFixedAssets.Value = 1 Then
    bolModuleSelected = True
    If Not Exist(FAData + FASetUpFileName) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (FASetUpFileName)
    End If
    If Not Exist(FAData + FAItemFileName) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (FAItemFileName)
    End If
    If Not Exist(FAData + FAAssetCodeName) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (FAAssetCodeName)
    End If
    If Not Exist(FAData + FADeptCodeName) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (FADeptCodeName)
    End If
    If Not Exist(FAData + FAFundCodeName) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (FAFundCodeName)
    End If
    'If Not Exist(FAData + FADprHistFileName) = True Then
    '  ErrorCode = 1
    '  lstMissingFilesListPayroll.AddItem (FADprHistFileName)
    'End If
  End If

'Validate DC files exists
  If chkVehDec.Value = 1 Then
    bolModuleSelected = True
    If Not Exist(DCData + DCCustFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (DCCustFile)
    End If
    If Not Exist(DCData + DCTranFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (DCTranFile)
    End If
    If Not Exist(DCData + DCSetupFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (DCSetupFile)
    End If
    If Not Exist(DCData + DCVCodeFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (DCVCodeFile)
    End If
    If Not Exist(DCData + DCVehFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (DCVehFile)
    End If
  End If
 
'ADD CM HERE
  If chkCashMgt.Value = 1 Then
    bolModuleSelected = True
    If Not Exist(CMTranFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (CMTranFile)
    End If
    If Not Exist(CMCodeFile) = True Then
      ErrorCode = 1
      lstMissingFilesListPayroll.AddItem (CMCodeFile)
    End If
  
  End If

  'Validate modules were selected
  If bolModuleSelected = False Then
     ErrorCode = 1
     lblInfo.Caption = "No modules were selected to export"
     lstReqFilesPayroll.Clear
     lstMissingFilesListPayroll.Clear
     lstReqFilesGL.Clear
     lstMissingFilesListGL.Clear
  End If
  
  'Display messege if now files were found
  If ErrorCode = 1 And bolModuleSelected = True Then
    lblInfo.Caption = "Required files were not found, export is aborted"
  End If
End Sub


Private Sub Form_Load()
  cmdProcess.Enabled = True
End Sub

